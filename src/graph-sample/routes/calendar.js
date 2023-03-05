// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const router = require('express-promise-router').default();
const graph = require('../graph.js');
const dateFns = require('date-fns');
const zonedTimeToUtc = require('date-fns-tz/zonedTimeToUtc');
const iana = require('windows-iana');
const { body, validationResult } = require('express-validator');
const validator = require('validator');
const { sendMail } = require('../graph.js');

/* GET /calendar */
// <GetRouteSnippet>
router.get('/',
  async function(req, res) {
    if (!req.session.userId) {
      // Redirect unauthenticated requests to home page
      res.redirect('/');
    } else {
      const params = {
        active: { calendar: true }
      };

      // Get the user
      const user = req.app.locals.users[req.session.userId];
      // Convert user's Windows time zone ("Pacific Standard Time")
      // to IANA format ("America/Los_Angeles")
      const timeZoneId = iana.findIana(user.timeZone)[0];
      console.log(`Time zone: ${timeZoneId.valueOf()}`);

      // Calculate the start and end of the current week
      // Get midnight on the start of the current week in the user's timezone,
      // but in UTC. For example, for Pacific Standard Time, the time value would be
      // 07:00:00Z
      var monthStart = zonedTimeToUtc(dateFns.startOfMonth(new Date()), timeZoneId.valueOf());
      var monthEnd = dateFns.addDays(monthStart, 31);
      console.log(`Start: ${dateFns.formatISO(monthStart)}`);

      try {
        // Get the events
        const events = await graph.getCalendarView(
          req.app.locals.msalClient,
          req.session.userId,
          dateFns.formatISO(monthStart),
          dateFns.formatISO(monthEnd),
          user.timeZone);

        //console.log(events)

        // Assign the events to the view parameters
        params.events = events.value;
      } catch (err) {
        req.flash('error_msg', {
          message: 'Could not fetch events',
          debug: JSON.stringify(err, Object.getOwnPropertyNames(err))
        });
      }
      const arr = await graph.getSchedule(req.app.locals.msalClient,
        req.session.userId);
      params.array = arr;
      console.log(params);
      res.render('../views/calendar.pug', {array:arr});
    }
  }
);
// </GetRouteSnippet>

// <GetEventFormSnippet>
/* GET /calendar/new */
router.get('/new',
  function(req, res) {
    if (!req.session.userId) {
      // Redirect unauthenticated requests to home page
      res.redirect('/');
    } else {
      res.locals.newEvent = {};
      res.render('newevent');
    }
  }
);
// </GetEventFormSnippet>
// <PostEventFormSnippet>
/* POST /calendar/new */
router.post('/new', [
  body('ev-subject').escape(),
  // Custom sanitizer converts ;-delimited string
  // to an array of strings
  body('ev-attendees').customSanitizer(value => {
    return value.split(';');
  // Custom validator to make sure each
  // entry is an email address
  }).custom(value => {
    value.forEach(element => {
      if (!validator.isEmail(element)) {
        throw new Error('Invalid email address');
      }
    });

    return true;
  }),
  // Ensure start and end are ISO 8601 date-time values
  body('ev-start').isISO8601(),
  body('ev-end').isISO8601(),
  body('ev-body').escape()
], async function(req, res) {
  if (!req.session.userId) {
    // Redirect unauthenticated requests to home page
    res.redirect('/');
  } else {
    // Build an object from the form values
    const formData = {
      subject: req.body['ev-subject'],
      attendees: req.body['ev-attendees'],
      start: req.body['ev-start'],
      end: req.body['ev-end'],
      body: req.body['ev-body']
    };


    // Check if there are any errors with the form values
    const formErrors = validationResult(req);
    if (!formErrors.isEmpty()) {

      let invalidFields = '';
      formErrors.array().forEach(error => {
        invalidFields += `${error.param.slice(3, error.param.length)},`;
      });

      // Preserve the user's input when re-rendering the form
      // Convert the attendees array back to a string
      formData.attendees = formData.attendees.join(';');
      return res.render('newevent', {
        newEvent: formData,
        error: [{ message: `Invalid input in the following fields: ${invalidFields}` }]
      });
    }

    // Get the user
    const user = req.app.locals.users[req.session.userId];

    // Create the event

    let returnValue = await graph.updateExcel(req.app.locals.msalClient, req.session.userId, formData.start,formData.end)
    if(!returnValue[0]){
      graph.sendMail(req.app.locals.msalClient, req.session.userId,{subject:'ERROR:CONFLICT',body:{contentType:'Text',content:'There has been an scheduling conflict in the shift you have booked'},address:'williamgra@cmail.carleton.ca'} )
      return res.redirect('/calendar')
    }

    var date = formData.start + ' to ' + formData.end;
    graph.sendMail(req.app.locals.msalClient, req.session.userId,{subject:'SHIFT BOOKED',body:{contentType:'Text',content:'You have a shift booked on '+date},address:'marcotoito@cmail.carleton.ca'} )
    try {
      await graph.createEvent(
        req.app.locals.msalClient,
        req.session.userId,
        formData,
        user.timeZone
      );
    } catch (error) {
      req.flash('error_msg', {
        message: 'Could not create event',
        debug: JSON.stringify(error, Object.getOwnPropertyNames(error))
      });
    }
    
    await graph.updatePayroll(req.app.locals.msalClient, req.session.userId, user.displayName, formData.start, formData.end)
    // Redirect back to the calendar view
    return res.redirect('/calendar');
  }
}
);
// </PostEventFormSnippet>
module.exports = router;

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const router = require('express-promise-router').default();
const graph = require('../graph.js');
const dateFns = require('date-fns');
const zonedTimeToUtc = require('date-fns-tz/zonedTimeToUtc');
const iana = require('windows-iana');
const { body, validationResult } = require('express-validator');
const validator = require('validator');

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
      var weekStart = zonedTimeToUtc(dateFns.startOfWeek(new Date()), timeZoneId.valueOf());
      var weekEnd = dateFns.addDays(weekStart, 7);
      console.log(`Start: ${dateFns.formatISO(weekStart)}`);

      try {
        // Get the events
        const events = await graph.getCalendarView(
          req.app.locals.msalClient,
          req.session.userId,
          dateFns.formatISO(weekStart),
          dateFns.formatISO(weekEnd),
          user.timeZone);

        // Assign the events to the view parameters
        params.events = events.value;
      } catch (err) {
        req.flash('error_msg', {
          message: 'Could not fetch events',
          debug: JSON.stringify(err, Object.getOwnPropertyNames(err))
        });
      }

      res.render('calendar.pug', params);
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
      res.render('newevent.pug');
    }
  }
);
// </GetEventFormSnippet>
// <PostEventFormSnippet>
/* POST /calendar/new */
router.post('/new', [
  
], async function(req, res) {
  if (!req.session.userId) {
    // Redirect unauthenticated requests to home page
    res.redirect('/');
  } else {
    // Build an object from the form values
    const formData = {
      attendees: req.body['ev-attendees'],
      start: req.body.start,
      end: req.body.end,
      date: req.body.date,
      body: "Working today"
    };

    // Check if there are any errors with the form values
    const formErrors = validationResult(req);
    if (!formErrors.isEmpty() ) {

      let invalidFields = '';
      formErrors.array().forEach(error => {
        invalidFields += `${error.param.slice(3, error.param.length)},`;
      });

      // Preserve the user's input when re-rendering the form
      // Convert the attendees array back to a string
      return res.render('newevent', {
        newEvent: formData,
        error: [{ message: `Invalid input in the following fields: ${invalidFields}` }]
      });
    }

    // Get the user
    const user = req.app.locals.users[req.session.userId];

    // Create the event
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

    var date = '1';
    graph.sendMail(req.app.locals.msalClient, req.session.userId,{subject:'SHIFT BOOKED',body:{contentType:'Text',content:'You have a shift booked on '+date},address:'marcobtoito@gmail.com'} )

    // Redirect back to the calendar view
    return res.redirect('/calendar');
  }
}
);
// </PostEventFormSnippet>
module.exports = router;

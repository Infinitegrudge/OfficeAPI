let host = ["localhost", "localhost"];

function lmao() {
    console.log("hi")
fetch(`http://${host[0]}:8000/login`, {
        method: 'GET', 
        headers: {
            'Content-Type': 'application/json'
        }
    })
    // fetch() returns a promise. When we have received a response from the server,
    // the promise's `then()` handler is called with the response.
    .then((response) => {
        // Our handler throws an error if the request did not succeed.
        if (!response.ok) {
        } else {
		}
    })
    // Catch any errors that might happen, and display a message.
    .catch((error) => console.log(err));
    console.log("hi")
}

document.getElementById("ye").onclick = lmao;
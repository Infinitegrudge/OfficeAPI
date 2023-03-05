function submit(){
    let formData = {};
	let date = document.getElementById("ev-date");
	let start = document.getElementById("ev-start");
	let end = document.getElementById("ev-end");
    formData.date = date.options[date.selectedIndex].value
    formData.start = start.options[start.selectedIndex].value
    formData.end = end.options[end.selectedIndex].value
	
	let req = new XMLHttpRequest();
	req.onreadystatechange = function() {
		if(this.readyState==4 && this.status==200){
			alert("Shift added");
		}
	}
	
	//Send a POST request to the server containing the recipe data
	req.open("POST", "/calendar/new");
	req.setRequestHeader("Content-Type", "application/json");
	req.send(JSON.stringify(formData));
}
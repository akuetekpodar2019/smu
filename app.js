// from data.js
var tableData = data;

// from data.js
var tableData = data;

// get a reference to the body
var tbody = d3.select("tbody");

function buildTable(data) {
	// First, clear out any existing data
	tbody.html("")

	tableData.forEach((sheepleReport) => {
		var row = tbody.append("tr");
		Object.values(sheepleReport).forEach((value) => {
			var cell = row.append('td');
			cell.text(value);
		});
	});

	var button = d3.select("#filter-btn");

	button.on("click", function () {

		var inputElement = d3.select("#datetime");

		var inputValue = inputElement.property("value");

		// alert("you clicked a button");

		//remove all table body rows
		d3.select("tbody").selectAll("tr").remove()

		let filterData = tableData.filter(ufoRow => ufoRow.datetime === inputValue);

		// alert(something)

		filterData.forEach((sheepleReport) => {
			var row = tbody.append("tr");
			Object.values(sheepleReport).forEach((value) => {
				var cell = row.append('td');
				cell.text(value);
			});
		});

	});
}
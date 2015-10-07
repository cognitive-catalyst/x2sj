'use strict'

var fs = require('fs');

module.exports = function toJSON(worksheet, outputFile) {

	var writeOutput = function(file, data) {
		fs.writeFile(file, data, function(error) {
			if (error) {
				console.error("write error:  " + error.message);
			} 
		});
	};					

	var buffer ='{\n';	
	
	for(var sheets=0; sheets<worksheet.length; sheets++) {
		var headers = worksheet[sheets][0];
		for(var row=1;row<worksheet[sheets].length;row++) {
			buffer+='  "'+worksheet[sheets][row][0]+'" : {\n';
			buffer+='    "doc" : {\n';
			for (var col=1;col<headers.length;col++) {
				buffer+='      "'+headers[col]+'" : '+JSON.stringify(worksheet[sheets][row][col])
				buffer+= ((col<headers.length-1)?',':'');
				buffer+='\n';
			}
			buffer+='    }\n';
			buffer+='  },\n'
		}
	}
	buffer+='  "commit" : { }\n';
	buffer+='}';
	writeOutput(outputFile,buffer);	
};


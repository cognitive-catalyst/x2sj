#!/usr/bin/env node

var parseXlsx = require('excel');
var cli = require('cli');
var fs = require('fs');

cli.parse({
	file: [ 'i', 'An XLSX file to process', 'file', null],
	output: [ 'o', 'Output file in JSON format', 'file', null]
});


cli.main(function(args,options) {

	if(options.file == null || options.output == null) {
		cli.getUsage();
		exit();
	}	

	var buffer = '';

	parseXlsx(options.file, function(err, data) {
		if(err) throw err;

		buffer +='{\n';	

		var headers = data[0];
		for(var row=1;row<data.length;row++) {
			buffer+='  "'+data[row][0]+'" : {\n';
			buffer+='    "doc" : {\n';
			for (var col=1;col<headers.length;col++) {
				buffer+='      "'+headers[col]+'" : "'+data[row][col]+'"' 
				buffer+= ((col<headers.length-1)?',':'');
				buffer+='\n';
			}
			buffer+='    }\n';
			buffer+='  },\n'
		}
		buffer+='  "commit" : { }\n';
		buffer+='}';
		writeOutput(options.output,buffer);	
	});

	var writeOutput = function(file, data) {
		fs.writeFile(file, data, function(error) {
			if (error) {
				console.error("write error:  " + error.message);
			} 
		});
	};

});

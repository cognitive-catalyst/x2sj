#!/usr/bin/env node
'use strict' 

var cli = require('cli'),
    xparser = require('../lib/xlsxParser'),
    tojson = require('../lib/jsonLib');

cli.parse({
	file: [ 'i', 'An XLSX file to process', 'file', null],
	output: [ 'o', 'Output file in JSON format', 'file', null]
});


cli.main(function(args,options) {

	if(options.file == null || options.output == null) {
		cli.getUsage();
		exit();
	}	
	tojson(xparser(options.file),options.output);
});


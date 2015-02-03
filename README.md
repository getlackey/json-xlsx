# Lackey json-xlsx converter

Converts between json and xlsx files

- only the first worksheet gets converted
- cells are merged as a visual convenience but they hold no special meaning

This module is part of the [Lackey framework](https://www.npmjs.com/package/lackey-framework) that is used to build the [Lackey CMS](http://lackey.io) amongst other projects.

## Usage
    
## xlsx -> json

    var JsonXlsxConverter = require('lackey-json-xlsx'),
        converter = new JsonXlsxConverter();

    converter
        .readXlsxFile('/my-path/my-file.xlsx')
        .then(function (self) {
            data = self.convertToJson().json;
            console.log(data); // or do whatever you need with it
        }).fail(function (err) {
            throw err;
        });

## json -> xlsx

var JsonXlsxConverter = require('lackey-json-xlsx'),
    converter = new JsonXlsxConverter(),
    xlsx,
    binXlsx;

    converter.json = {// any valid js object
        test: 'OK',
        id: '1234' 
    };

    xlsx = converter.convertToXlsx().xlsx;

    // if you need a binary object
    binXlsx = new Buffer(converter.convertToXlsx().xlsx, 'binary');


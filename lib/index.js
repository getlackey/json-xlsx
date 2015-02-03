/*jslint node:true, browser:true, nomen:true */
'use strict';
/*
    Copyright 2015 Enigma Marketing Services Limited

    Licensed under the Apache License, Version 2.0 (the "License");
    you may not use this file except in compliance with the License.
    You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

    Unless required by applicable law or agreed to in writing, software
    distributed under the License is distributed on an "AS IS" BASIS,
    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
    See the License for the specific language governing permissions and
    limitations under the License.
*/

var fs = require('fs'),
    XLSX = require('xlsx'),
    Q = require('q');

/*
    converts between json and xlsx
        - only the first worksheet gets converted
        - cells are merged as a visual convenience but they hold no special meaning
*/

function colLetterToNumber(letters) {
    var number = 0,
        i = 0;

    for (i = 0; i < letters.length; i += 1) {
        //number += (letters.length - i - 1) * (letters.charCodeAt(i) - 64);
        number += Math.pow(26, i) * (letters.charCodeAt(letters.length - i - 1) - 64);
    }
    return number;
}

function colNumberToLetter(num) {
    if (num <= 0) {
        throw new Error('Invalid number');
    }

    if (num < 27) {
        return String.fromCharCode(64 + num);
    }
    return colNumberToLetter((num - 1) / 26) + colNumberToLetter((num - 1) % 26 + 1);
}

function getLetters(value) {
    return value.replace(/[0-9]+/, '');
}

function getNumbers(value) {
    return +value.replace(/[A-Z]+/, '');
}

function sortCells(firstSheet) {
    return Object.keys(firstSheet).sort(function (a, b) {
        var aLetters = getLetters(a),
            bLetters = getLetters(b),
            aNumbers,
            bNumbers;

        // number part is only relevant if the letters are equal
        if (aLetters === bLetters) {
            aNumbers = getNumbers(a);
            bNumbers = getNumbers(b);
            return aNumbers - bNumbers;
        }

        return colLetterToNumber(aLetters) - colLetterToNumber(bLetters);
    });
}

function datenum(v, date1904) {
    var epoch;

    if (date1904) {
        v += 1462;
    }
    epoch = Date.parse(v);
    return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
}

function convertArrayToLiteralObject(arr) {
    var obj = {},
        i = 0;

    for (i = 0; i < arr.length; i += 1) {
        obj[i] = (arr[i] === undefined ? 'NULL' : arr[i]);
    }

    return obj;
}

function convertNestedArrayToLiteralObject(obj) {
    if (typeof obj === 'object' && obj.length) {
        obj = convertArrayToLiteralObject(obj);
    }
    Object.keys(obj).forEach(function (key) {
        if (typeof obj[key] === 'object') {
            obj[key] = convertNestedArrayToLiteralObject(obj[key]);
        }
    });
    return obj;
}
/*
 * Workbook management
 */
function Workbook() {
    if (!(this instanceof Workbook)) {
        return new Workbook();
    }
    this.SheetNames = [];
    this.Sheets = {};
}

Workbook.prototype.addWorkSheet = function (name, obj) {
    var self = this,
        ws = {},
        range = {
            s: {
                c: 0,
                r: 0
            },
            e: {
                c: 0,
                r: 0
            }
        },
        merges = [],
        dimensions;

    self.SheetNames.push(name);
    self.Sheets[name] = ws;

    function addCell(col, row, val) {
        var cell = {
                v: val
            },
            cell_ref;

        if (cell.v === null) {
            return cell;
        }

        cell_ref = XLSX.utils.encode_cell({
            c: col,
            r: row
        });

        if (typeof cell.v === 'number') {
            cell.t = 'n';
        } else if (typeof cell.v === 'boolean') {
            cell.t = 'b';
        } else if (cell.v instanceof Date) {
            cell.t = 'n';
            cell.z = XLSX.SSF._table[14];
            cell.v = datenum(cell.v);
        } else {
            cell.t = 's';
        }


        ws[cell_ref] = cell;
    }

    function parse(obj, c, r) {
        var col = +c || 0,
            row = +r || 0,
            objKeys = Object.keys(obj),
            totalDimensions = {
                numCols: 0, //min columns for key/value assuming the obj isn't invalid
                numRows: 0
            };

        objKeys.forEach(function (key) {
            var val = obj[key],
                dim,
                isObject = (typeof val === 'object'),
                isDate = (val instanceof Date),
                isBSON = (typeof val === 'object' && val._bsontype); // ObjectId

            addCell(col, row + totalDimensions.numRows, key);

            if (isObject && !isDate && !isBSON) {
                dim = parse(val, col + 1, row + totalDimensions.numRows);

                if (dim.numRows > 1) {
                    merges.push({
                        s: {
                            c: col,
                            r: row + totalDimensions.numRows
                        },
                        e: {
                            c: col,
                            r: row + totalDimensions.numRows + dim.numRows - 1
                        }
                    });
                }

                if (dim.numCols >= totalDimensions.numCols) {
                    totalDimensions.numCols = dim.numCols + 1;
                }
                totalDimensions.numRows += dim.numRows;
            } else {
                addCell(col + 1, row + totalDimensions.numRows, val);
                if (totalDimensions.numCols < 2) {
                    totalDimensions.numCols = 2;
                }
                totalDimensions.numRows += 1;
            }
        });

        return totalDimensions;
    }

    // we need to convert arrays to literal objects so 
    // the exported excel file shows the array indexes as 
    // numbers in their own columns
    dimensions = parse(convertNestedArrayToLiteralObject(obj));
    range.e.c = dimensions.numCols - 1;
    range.e.r = dimensions.numRows - 1;

    ws['!ref'] = XLSX.utils.encode_range(range);

    if (merges.length) {
        ws['!merges'] = merges;
    }

    return ws;
};

/* 
 * JSON to XSLX converted
 * and back...
 */
var Obj = function () {
    var self = this;
    return self;
};

Obj.prototype.workbook = null;
Obj.prototype.json = null;
Obj.prototype.xlsx = null;

Obj.prototype.readXlsxFile = function (filePath) {
    var self = this,
        deferred = Q.defer();

    fs.readFile(filePath, function (err, data) {
        if (err) {
            deferred.reject(err);
        }

        self.workbook = XLSX.read(data, {
            type: "binary"
        });
        deferred.resolve(self);
    });

    return deferred.promise;
};

Obj.prototype.saveXlsxFile = function (filename) {
    var self = this;
    fs.writeFile(filename, self.xlsx, 'binary', function (err) {
        if (err) {
            throw err;
        }
    });
};

Obj.prototype.convertToJson = function () {
    var self = this,
        worksheets = self.workbook.SheetNames,
        firstSheet = self.workbook.Sheets[worksheets[0]],
        cellsIndex = sortCells(firstSheet);

    function processColumn(obj, columnLetter, columnIndex, maxColumnIndex) {
        var initIndex = cellsIndex.indexOf(columnLetter + columnIndex),
            cellIndex,
            cellLetters,
            cellNumbers,
            data,
            nextColumn,
            nextData,
            afterNextColumn,
            afterNextData,
            nextRowData,
            i;


        if (initIndex === -1) {
            throw new Error('Cell not found ' + columnLetter + ':' + columnIndex);
        }

        for (i = initIndex; i < cellsIndex.length; i += 1) {
            cellIndex = cellsIndex[i];
            cellLetters = getLetters(cellIndex);
            cellNumbers = getNumbers(cellIndex);

            if (columnLetter !== cellLetters) {
                break;
            }

            if (maxColumnIndex && cellNumbers >= maxColumnIndex) {
                break;
            }

            data = firstSheet[cellIndex];

            // create an object if none provided. 
            // for number keys, it's an array, otherwise and object
            obj = obj || (/^[0-9]+$/.test(data.v) ? [] : {});

            nextColumn = colNumberToLetter(colLetterToNumber(columnLetter) + 1);
            nextData = firstSheet[nextColumn + cellNumbers];

            afterNextColumn = colNumberToLetter(colLetterToNumber(columnLetter) + 2);
            afterNextData = firstSheet[afterNextColumn + cellNumbers];

            if (!nextData) {
                throw new Error('Undefined cell value');
            }

            if (obj[data.v] !== undefined) {
                throw new Error('Duplicated cell value');
            }

            if (afterNextData) {
                // if the key is a number it's an array, otherwise
                // it's an object
                obj[data.v] = (/^[0-9]+$/.test(nextData.v) ? [] : {});

                nextRowData = cellsIndex[i + 1];
                processColumn(obj[data.v], nextColumn, cellNumbers, (getLetters(nextRowData) === columnLetter ? getNumbers(nextRowData) : null));
            } else {
                obj[data.v] = nextData.v;
            }
        }

        return obj;
    }

    self.json = processColumn(null, 'A', 1);

    return self;
};

Obj.prototype.convertToXlsx = function () {
    var self = this,
        wb = new Workbook(),
        opts;
    // export workbook
    wb.addWorkSheet('Exported Data', self.json);
    self.workbook = wb;
    // export xlsx binary object
    opts = {
        bookType: 'xlsx',
        bookSST: false,
        type: 'binary'
    };
    self.xlsx = XLSX.write(wb, opts);

    return self;
};

module.exports = Obj;
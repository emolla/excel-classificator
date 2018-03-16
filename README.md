# excel-classificator

[node](http://nodejs.org/) wrapper for parsing spreadsheets. Supports xls, xlsx and csv files.

You can install this module using [npm](http://github.com/isaacs/npm):

    npm install excel-classificator

Requires [python](http://www.python.org/) to be installed
Requires xlrd and csv python modules
    pip install xlrd csv chardet

For system-specific installation view the [Wiki](https://github.com/emolla/excel-classificator/wiki)

## API

<a name="to_csv" />
### to_csv(options, callback(err, response))
Convert and normalize spreadsheet file in csv file filtering bad headers

__Arguments__

* inFile - Filepath of the source spreadsheet
* outFile - Filepath of the source spreadsheet

__Example__

```js
var excelclassificator = require('excel-classificator');
excelclassificator.to_csv({
  inFile: 'file.xls', // xls, xlsx or csv extension 
  outFile: 'file.csv',
  headerFields: {'billing': ['invoice', 'customer']}
}, function(err, response){
  if(err) console.error(err);
  consol.log(response);
});
```
__Sample output__

```json
{'filetype': 'billing', 'numLines': numLines, 'outfile': os.path.abspath(args.outFile)}
```
---------------------------------------

## Running Tests

There are unit tests in `test/` directory. To run test suite first run the following command to install dependencies.

    npm install

then run the tests:

    grunt nodeunit

NOTE: Install `npm install -g grunt-cli` for running tests.

## License

Copyright (c) 2018 Quique Molla

Licensed under the MIT license.
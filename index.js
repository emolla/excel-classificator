var path = require('path');
var Promise = require('bluebird')
var cmd = require('node-cmd')
var excelNormalizer = {}

const getAsync = Promise.promisify(cmd.get, { multiArgs: true, context: cmd })

function to_csv(options) {
    var inputFile = ''
    var outputFile = ''
    var headerFile = ''

    try{
        inputFile = path.relative(__dirname, options.inputFile)
    } catch (error){
        console.error(error)
    }

    try{
        outputFile = path.relative(__dirname, options.outputFile)
    } catch (error){
        console.error(error)
    }

    try{
        headerFile = path.relative(__dirname, options.headersFile)
    } catch (error){
        console.error(error)
    }

    var args = [
        '-i', '"' + inputFile + '"',
        '-e', '"' + options.extension + '"',
        '-o', '"' + outputFile +'"',
        '-hf', '"' + headerFile +'"'
    ];

    return getAsync('python ' + __dirname + '/convert.py ' + args.join(' '))
};

excelNormalizer.to_csv = to_csv
module.exports = excelNormalizer;
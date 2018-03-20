var path = require('path')
var Promise = require('bluebird')
var cmd = require('node-cmd')

var excelNormalizer = {}

const getAsyncCommand = Promise.promisify(cmd.get, { multiArgs: true, context: cmd })

function detect(options) {
    var inputFile = ''
    var headerFile = ''
    var extension = 'csv'

    options.extension = options.extension || extension

    try{
        inputFile = path.relative(__dirname, options.inputFile)
    } catch (error){
        console.error(error)
    }

    try{
        headerFile = path.resolve(options.headersFile)
    } catch (error){
        console.error(error)
    }

    if (options.extension.indexOf('openxml') > 0){
        extension = 'xlsx'
    } else if (options.extension.indexOf('excel') > 0){
        extension = 'xls'
    }

    var args = [
        '-i', '"' + inputFile + '"',
        '-ex', '"' + extension + '"',
        '-d', '"true"',
        '-hf', '"' + headerFile +'"'
    ];

    var command = 'python ' + __dirname + '/convert.py ' + args.join(' ')

    return getAsyncCommand(command)
};

function to_csv(options) {
    var inputFile = ''
    var outputFile = ''
    var headerFile = ''
    var extension = 'csv'

    options.extension = options.extension || extension

    try{
        inputFile = path.relative(__dirname, options.inputFile)
    } catch (error){
        console.error(error)
    }

    try{
        outputFile = path.relative(__dirname, options.outputFile)
    } catch (error){
        console.error('No se ha especificado el fichero de salida.')
    }

    try{
        headerFile = path.resolve(options.headersFile)
    } catch (error){
        console.error(error)
    }

    if (options.extension.indexOf('openxml') > 0){
        extension = 'xlsx'
    } else if (options.extension.indexOf('excel') > 0){
        extension = 'xls'
    }

    var args = [
        '-i', '"' + inputFile + '"',
        '-ex', '"' + extension + '"',
        '-o', '"' + outputFile +'"',
        '-hf', '"' + headerFile +'"'
    ];

    var command = 'python ' + __dirname + '/convert.py ' + args.join(' ')

    console.log(command)

    return getAsyncCommand(command)
};

excelNormalizer.to_csv = to_csv
excelNormalizer.detect = detect
module.exports = excelNormalizer;
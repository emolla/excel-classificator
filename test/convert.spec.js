'use strict'

const normalizer = require('../')
var chai = require('chai');
var expect = chai.expect;

chai.use(require('chai-json-schema'));

describe('excel-classificator module', function() {
    before(function () { })
    it('should return billing filetype', function (done) {
        this.timeout(15000);
        normalizer.to_csv({
            'inputFile': './test/test-files/1stLineHeader_large.csv',
            'outputFile': './test/output/billing_normalized.csv',
            'headersFile': './test/headersFile.json'
        }).then(function (response) {
            var expected = {
                "numLines": 541909,
                "fileType": "billing"
            }
            var r = JSON.parse(response[0])
            expect(expected.numLines).to.equal(r.numLines)
            expect(expected.fileType).to.equal(r.fileType)
            done()
        }).catch(function (err){
            done(err)
        })
    })
    it('should return accounting filetype', function (done) {
        normalizer.to_csv({
            'inputFile': './test/test-files/1stLineHeader.csv',
            'outputFile': './test/output/1stLineHeader_normalized.csv',
            'headersFile': './test/headersFile.json'
        }).then(function (response) {
            var expected = {
                "numLines": 55,
                "fileType": "accounting"
            }
            var r = JSON.parse(response[0])
            expect(expected.numLines).to.equal(r.numLines)
            expect(expected.fileType).to.equal(r.fileType)
            done()
        }).catch(function (err){
            done(err)
        })
    })
    it('should return accounting filetype', function (done) {
        normalizer.to_csv({
            'inputFile': './test/test-files/MACOS_spreadsheet.xls',
            'outputFile': './test/output/MACOS_spreadsheet_normalized.csv',
            'headersFile': './test/headersFile.json'
        }).then(function (response) {
            var expected = {
                "numLines": 75,
                "fileType": "accounting"
            }
            var r = JSON.parse(response[0])
            expect(expected.numLines).to.equal(r.numLines)
            expect(expected.fileType).to.equal(r.fileType)
            done()
        }).catch(function (err){
            done(err)
        })
    })
    it('should return accounting filetype', function (done) {
        this.timeout(15000);
        normalizer.to_csv({
            'inputFile': './test/test-files/spreadsheet.xls',
            'outputFile': './test/output/spreadsheet_normalized.csv',
            'headersFile': './test/headersFile.json'
        }).then(function (response) {
            var expected = {
                "numLines": 65535,
                "fileType": "billing"
            }
            var r = JSON.parse(response[0])
            expect(expected.numLines).to.equal(r.numLines)
            expect(expected.fileType).to.equal(r.fileType)
            done()
        }).catch(function (err){
            done(err)
        })
    })
    it('should return accounting filetype', function (done) {
        this.timeout(15000);
        normalizer.to_csv({
            'inputFile': './test/test-files/spreadsheet.xlsx',
            'outputFile': './test/output/spreadsheetx_normalized.csv',
            'headersFile': './test/headersFile.json'
        }).then(function (response) {
            var expected = {
                "numLines": 65535,
                "fileType": "billing"
            }
            var r = JSON.parse(response[0])
            expect(expected.numLines).to.equal(r.numLines)
            expect(expected.fileType).to.equal(r.fileType)
            done()
        }).catch(function (err){
            done(err)
        })
    })

})
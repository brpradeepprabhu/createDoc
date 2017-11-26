var express = require('express');
var app = express();
var async = require('async');
var fs = require('fs');
var path = require('path');

var cors = require('cors');
app.use(cors());

var bodyparser = require('body-parser');
app.use(bodyparser.json());

var tempfile = require('tempfile');
var officegen = require('officegen');



app.get('/docx', function (req, res) {
    var docx = officegen({
        type: 'docx',
        orientation: 'portrait',
        pageMargins: {
            top: 1000,
            left: 1000,
            bottom: 1000,
            right: 1000
        }
    });


    docx.on('error', function (err) {
        console.log(err);
    });
    var pObj = docx.createP({
        align: 'center'
    });
    pObj.addText('Change Order', {
        font_face: 'Arial',
        font_size: 10,
        bold: true,
        italic: true
    });
    var pObj = docx.createP({
        align: 'center'
    });
    pObj.addText('between', {
        font_face: 'Arial',
        font_size: 10,
        bold: true,
        italic: true
    })
    var pObj = docx.createP({
        align: 'center'
    });
    pObj.addText('Infosys Limited', {
        font_face: 'Arial',
        font_size: 10,
        bold: true,
        italic: true
    })
    var pObj = docx.createP({
        align: 'center'
    });
    pObj.addText('and', {
        font_face: 'Arial',
        font_size: 10,
        bold: true,
        italic: true
    });
    var pObj = docx.createP({
        align: 'center'
    });
    pObj.addText('PricewaterhouseCoopers LLP', {
        font_face: 'Arial',
        font_size: 10,
        bold: true,
        italic: true
    })

    var pObj = docx.createP({
        align: 'justify'
    });
    pObj.addText('This Change Order No. [#], effective [Change Order Effective Date] (this "Change Order") amends Statement of Work [SOW Title Page Title], dated [SOW Effective Date] (the "SOW") under the Amended and Restated Services Agreement, dated July 1, 2017 (the "Agreement"), between PricewaterhouseCoopers LLP ("PwC"), and Infosys Limited ("Supplier").', {
        font_face: 'Arial',
        font_size: 10
    })

    var pObj = docx.createP({
        align: 'justify'
    });
    pObj.addText('PricewaterhouseCoopers LLP ("PwC"), and Infosys Limited ("Supplier"). Supplier shall complete the remainder of the Change Order, except for the approval/rejection portion, which shall be completed by PwC in its sole discretion.  Each section will be as long or short as the circumstances require.  Additional pages will be attached as necessary.', {
        font_face: 'Arial',
        font_size: 10
    })
    var pObj = docx.createListOfNumbers();

    pObj.addText('Describe changes, modifications, or additions to the services.');


    var pObj = docx.createP({
        indent: 1000
    });
    pObj.addText('CR Start Date:', {
        bold: true,
        font_face: 'Arial',
        font_size: 10,
        indent: 1000,
    });

    var pObj = docx.createP({
        indent: 1000
    });
    pObj.addText('CR End Date:', {
        bold: true,
        font_face: 'Arial',
        font_size: 10,
        indent: 1000
    });

    var pObj = docx.createListOfNumbers();

    pObj.addText('Modifications, clarifications or supplements by Supplier to description of desired changes or additions requested in Section 1 above, if any');

    var pObj = docx.createListOfNumbers();

    pObj.addText('Necessity, availability and assignment of requisite skill sets and/or resources to make requested modification or additions.');

    var pObj = docx.createListOfNumbers();

    pObj.addText('Impact on costs, delivery schedule, and other requirements.');


    res.writeHead(200, {
        "Content-Type": "application/vnd.openxmlformats-officedocument.documentml.document",
        'Content-disposition': 'attachment; filename=testdoc.docx'
    });
    docx.generate(res);
});


app.listen(3000, function (req, res) {
    console.log('server started');
})
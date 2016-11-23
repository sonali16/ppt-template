var fs = require('fs');
var async = require('async');
var JSZip = require('jszip');
var xml2js = require('xml2js');
var Slide = require('./Slide');

var Presentation = module.exports = function() {
    this.contents = {};
}

Presentation.prototype.load = function(pptx, callback) {

    var self = this;

    //解壓縮
    JSZip.loadAsync(pptx).then(function(zip) {

        //異步迭代各個檔案
        async.each(Object.keys(zip.files), function(key, callback) {

            //圖像之外都當作文字解析
            var ext = key.substr(key.lastIndexOf('.'));
            var type = (ext == '.xml' || ext == '.rels') ? 'string' : 'array';

            //將各檔案轉成字串，紀錄(檔名 : 純文字)
            zip.files[key].async(type)
                .then(function(content) {
                    //console.log(key);
                    //console.log(content);
                    self.contents[key] = content;
                    callback();
                });

        }, callback);
    });
}

Presentation.prototype.loadFile = function(pptxFilePath, callback) {

    var self = this;

    //讀檔
    fs.readFile(pptxFilePath, function(err, pptxFile) {

        if (err) throw err;

        self.load(pptxFile, callback);

    })
}

Presentation.prototype.streamAs = function(stream, callback) {

    var newZip = JSZip();

    for (var key in this.contents) {
        if (this.contents[key]) newZip.file(key, this.contents[key]);
    }

    newZip
        .generateNodeStream({
            type: 'nodebuffer',
            streamFiles: true
        })
        .pipe(stream)
        .on('finish', callback);
};

Presentation.prototype.saveAs = function(outputFilePath, callback) {

    this.streamAs(fs.createWriteStream(outputFilePath), callback);
};

Presentation.prototype.getSlideCount = function() {
    return Object.keys(this.contents).filter(function(key) {
        return key.substr(0, 16) === "ppt/slides/slide"
    }).length;
}

Presentation.prototype.getSlide = function(index) {
    if (index < 1 || index > this.getSlideCount()) return null;
    var rel = this.contents['ppt/slides/_rels/slide' + index + '.xml.rels'];
    var content = this.contents['ppt/slides/slide' + index + '.xml'];
    return new Slide(rel, content);
}

Presentation.prototype.clone = function() {
    var newPresentation = new Presentation();
    newPresentation.contents = JSON.parse(JSON.stringify(this.contents))
    return newPresentation;
}

Presentation.prototype.generate = function(slides, callback) {
    var newPresentation = this.clone();
    var newContents = newPresentation.contents;
    var oldCount = newPresentation.getSlideCount();

    //清掉舊的 ppt/slides/slideX.xml & ppt/slides/_rels/slideX.xml.rels
    for (var i = 0; i < oldCount; i++) {
        delete newContents['ppt/slides/_rels/slide' + (i + 1) + '.xml.rels'];
        delete newContents['ppt/slides/slide' + (i + 1) + '.xml'];
    }

    //加入新的 ppt/slides/slideX.xml & ppt/slides/_rels/slideX.xml.rels
    for (var i = 0; i < slides.length; i++) {
        var slide = slides[i];
        newContents['ppt/slides/_rels/slide' + (i + 1) + '.xml.rels'] = slide.rel;
        newContents['ppt/slides/slide' + (i + 1) + '.xml'] = slide.content;
    }

    //修改~
    async.parallel([
            //修改 [Content_Types].xml
            function(callback) {

                xml2js.parseString(newPresentation.contents['[Content_Types].xml'], function(err, xmlJs) {

                    if (err) throw err;

                    //清掉舊的
                    xmlJs['Types']['Override'].forEach(function(override, index) {
                        if (override['$'].PartName.substr(0, 17) == '/ppt/slides/slide') delete xmlJs['Types']['Override'][index];
                    });

                    //加入新的
                    for (var i = 0; i < slides.length; i++) {
                        xmlJs['Types']['Override'].push({
                            '$': {
                                'PartName': '/ppt/slides/slide' + (i + 1) + '.xml',
                                'ContentType': 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
                            }
                        });
                    }

                    newContents['[Content_Types].xml'] = new xml2js.Builder().buildObject(xmlJs);
                    callback(null, '1');
                });

            },
            //修改 ppt/_rels/presentation.xml.rels & ppt/presentation.xml
            function(callback) {
                xml2js.parseString(newPresentation.contents['ppt/_rels/presentation.xml.rels'], function(err, xmlJs1) {
                    if (err) throw err;

                    xml2js.parseString(newPresentation.contents['ppt/presentation.xml'], function(err, xmlJs2) {
                        if (err) throw err;

                        //刪除舊的
                        xmlJs1['Relationships']['Relationship'].forEach(function(relationship, index) {
                            if (relationship['$'].Target.substr(0, 12) == 'slides/slide') delete xmlJs1['Relationships']['Relationship'][index];
                        });

                        /*
                        //整理rId
                        xmlJs1['Relationships']['Relationship'].forEach(function(relationship, index) {
                            if (relationship) relationship['$'].Id = 'rId' + (index + 1);
                        });
                        */

                        //加入新的
                        var rId = xmlJs1['Relationships']['Relationship'].length;
                        for (var i = 1; i <= slides.length; i++) {
                            xmlJs1['Relationships']['Relationship'].push({
                                '$': {
                                    'Id': 'rId' + (rId + i),
                                    'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
                                    'Target': 'slides/slide' + i + '.xml'
                                }
                            });
                        }

                        //計算 maxId
                        var maxId = 0;
                        xmlJs2['p:presentation']['p:sldIdLst'][0]['p:sldId'].forEach(function(ob) {
                            if (+ob['$'].id > maxId) maxId = +ob['$'].id;
                        });

                        //刪除舊的
                        xmlJs2['p:presentation']['p:sldIdLst'][0]['p:sldId'].forEach(function(ob, index) {
                            delete xmlJs2['p:presentation']['p:sldIdLst'][0]['p:sldId'][index];
                        });

                        //加入新的
                        for (var i = 1; i <= slides.length; i++) {
                            xmlJs2['p:presentation']['p:sldIdLst'][0]['p:sldId'].push({
                                '$': {
                                    'id': '' + (+maxId + i),
                                    'r:id': 'rId' + (rId + i)
                                }
                            });
                        }

                        newContents['ppt/_rels/presentation.xml.rels'] = new xml2js.Builder().buildObject(xmlJs1);
                        newContents['ppt/presentation.xml'] = new xml2js.Builder().buildObject(xmlJs2);
                        callback(null, '2');

                    });
                });
            },
            //修改 docProps/app.xml (increment slidecount)
            function(callback) {

                xml2js.parseString(newPresentation.contents['docProps/app.xml'], function(err, xmlJs) {
                    xmlJs["Properties"]["Slides"][0] = slides.length;
                    newContents['docProps/app.xml'] = new xml2js.Builder().buildObject(xmlJs);
                    callback(null, '3');
                });
            }
        ],
        function(err, results) {
            callback(newPresentation);
        });
}

/*
Presentation.prototype.getXml2Js = function(key, callback) {
    xml2js.parseString(this.contents[key], function(err, xmlJs) {
        if (err) throw err;
        callback(xmlJs);
    });
}

Presentation.prototype.getJs2Xml = getJs2Xml;

function getJs2Xml(js) {
    return new xml2js.Builder().buildObject(js);
}
*/
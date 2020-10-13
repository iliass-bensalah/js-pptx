var JSZip = require('jszip'); // this works on browser
var async = require('async'); // this works on browser
var xml2js = require('xml2js'); // this works on browser?
var XmlNode = require('./xmlnode');
var Search = require('./util/search');

var Slide = require('./slide');

var Presentation = function(object) {
	this.content = {};
};

// fundamentally asynchronous because xml2js.parseString() is async
Presentation.prototype.load = function(data, done) {
	JSZip.loadAsync(data).then(zip => {
		var content = this.content;
		async.each(Object.keys(zip.files), function(key, callback) {
			if (key.substr(0,11)=="ppt/media/i"){
				 zip.files[key].async("arraybuffer").then(file => {
					content[key] = file;
					callback(null)
				}, err => {
					content[key] = "Failed to read file as arraybuffer from zip.\n" + JSON.stringify(err);
					callback(null)
				});
			}
			else{
			var ext = key.substr(key.lastIndexOf('.'));
			if (ext == '.xml' || ext == '.rels') {
				zip.files[key].async("string").then(xml => {
					xml2js.parseString(xml, function(err, js) {
						content[key] = js;
						callback(null);
					});
				}, err => {
					content[key] = "Failed to read file as string from zip.\n" + JSON.stringify(err);
					callback(null);
				});
			} else if (ext == '.png' || ext == '.jpg' || ext == '.jpeg') {
	
				zip.files[key].async("arraybuffer").then(binary => {
					content[key] = binary;
					callback(null);
				}, err => {
					content[key] = "Failed to read file as arraybuffer from zip.\n" + JSON.stringify(err);
					callback(null);
				});
	
				
			} else {
				zip.files[key].async("string").then(sText => {
					content[key] = sText;
					callback(null);
				}, err => {
					content[key] = "Failed to read file as string from zip.\n" + JSON.stringify(err);
					callback(null);
				});
			}
			}
		}, done);
	});
};

Presentation.prototype.toJSON = function() {
	return this.content;
};

Presentation.prototype.toBuffer = function() {
	var zip2 = new JSZip();
	var content = this.content;
	for ( var key in content) {
		if (content.hasOwnProperty(key)) {
			if (key.substr(0,10)=="ppt/media/"){
				zip2.file(key, content[key], {
					binary : true
				});
			}
			var ext = key.substr(key.lastIndexOf('.'));
			if (ext == '.xml' || ext == '.rels') {
				var builder = new xml2js.Builder({
					renderOpts : {
						pretty : true
					}
				});
				var xml2 = (builder.buildObject(content[key]));
				zip2.file(key, xml2);
			} else if (ext == '.png' || ext == '.jpg' || ext == '.jpeg') {

			
				zip2.file(key, content[key], {
					binary : true
				});
				

			} else {
				zip2.file(key, content[key]);
			}
		}
	}
	// zip2.file("docProps/app.xml", '<?xml version="1.0" encoding="UTF-8"
	// standalone="yes"?><Properties
	// xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
	// xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><TotalTime>1</TotalTime><Words>72</Words><Application>Microsoft
	// Macintosh PowerPoint</Application><PresentationFormat>On-screen Show
	// (4:3)</PresentationFormat><Paragraphs>12</Paragraphs><Slides>3</Slides><Notes>0</Notes><HiddenSlides>0</HiddenSlides><MMClips>0</MMClips><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector
	// size="4"
	// baseType="variant"><vt:variant><vt:lpstr>Theme</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant><vt:variant><vt:lpstr>Slide
	// Titles</vt:lpstr></vt:variant><vt:variant><vt:i4>3</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector
	// size="4" baseType="lpstr"><vt:lpstr>Office Theme</vt:lpstr><vt:lpstr>This
	// is the title</vt:lpstr><vt:lpstr>This is the
	// title</vt:lpstr><vt:lpstr>This is the
	// title</vt:lpstr></vt:vector></TitlesOfParts><Company>Proven,
	// Inc.</Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>14.0000</AppVersion></Properties>');
//	zip2
//			.file(
//					"docProps/app.xml",
//					'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><TotalTime>0</TotalTime><Words>0</Words><Application>Microsoft Office PowerPoint</Application><PresentationFormat>On-screen Show (4:3)</PresentationFormat><Paragraphs>0</Paragraphs><Slides>2</Slides><Notes>0</Notes><HiddenSlides>0</HiddenSlides><MMClips>0</MMClips><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="4" baseType="variant"><vt:variant><vt:lpstr>Theme</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant><vt:variant><vt:lpstr>Slide Titles</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="3" baseType="lpstr"><vt:lpstr>Office Theme</vt:lpstr><vt:lpstr>PowerPoint Presentation</vt:lpstr><vt:lpstr>PowerPoint Presentation</vt:lpstr></vt:vector></TitlesOfParts><Company>Proven, Inc.</Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>14.0000</AppVersion></Properties>')
	return zip2.generateAsync({
		type : "nodebuffer"
	});
};

Presentation.prototype.registerChart = function(chartName, content) {
	this.content['ppt/charts/' + chartName + '.xml'] = content;

	// '[Content_Types].xml' .. add references to the chart and the spreadsheet
	this.content["[Content_Types].xml"]["Types"]["Override"]
			.push(XmlNode()
					.attr('PartName', "/ppt/charts/" + chartName + ".xml")
					.attr('ContentType',
							"application/vnd.openxmlformats-officedocument.drawingml.chart+xml").el);

	var defaults = this.content["[Content_Types].xml"]["Types"]["Default"]
			.filter(function(el) {
				return el['$']['Extension'] == 'xlsx'
			});

	if (defaults.length == 0) {
		this.content["[Content_Types].xml"]["Types"]["Default"]
				.push(XmlNode()
						.attr('Extension', 'xlsx')
						.attr('ContentType',
								"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet").el);
	}
}

Presentation.prototype.registerChartWorkbook = function(chartName,
		workbookContent) {

	var numWorksheets = this.getWorksheetCount();
	var worksheetName = 'Microsoft_Excel_Sheet' + (numWorksheets + 1) + '.xlsx';

	this.content["ppt/embeddings/" + worksheetName] = workbookContent;

	// ppt/charts/_rels/chart1.xml.rels
	this.content["ppt/charts/_rels/" + chartName + ".xml.rels"] = XmlNode()
			.setChild(
					"Relationships",
					XmlNode()
							.attr(
									{
										'xmlns' : "http://schemas.openxmlformats.org/package/2006/relationships"
									})
							.addChild(
									'Relationship',
									XmlNode()
											.attr(
													{
														"Id" : "rId1",
														"Type" : "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package",
														"Target" : "../embeddings/"
																+ worksheetName
													}))).el;
}

Presentation.prototype.getSlideCount = function() {
	return Object.keys(this.content).filter(function(key) {
		return key.substr(0, 16) === "ppt/slides/slide"
	}).length;
}

Presentation.prototype.getChartCount = function() {
	return Object.keys(this.content).filter(function(key) {
		return key.substr(0, 16) === "ppt/charts/chart"
	}).length;
}

Presentation.prototype.getWorksheetCount = function() {
	return Object.keys(this.content).filter(function(key) {
		return key.substr(0, 36) === "ppt/embeddings/Microsoft_Excel_Sheet"
	}).length;
}
Presentation.prototype.getImgCount = function() {
	return Object.keys(this.content).filter(function(key) {
		return key.substr(0, 15) === "ppt/media/image"
	}).length;
}

Presentation.prototype.getSlide = function(slideNumber) {
	return new Slide({
		content : this.content['ppt/slides/slide' + slideNumber + '.xml'],
		presentation : this,
		name : 'slide' + slideNumber
	});
}

Presentation.prototype.getSlideByName = function(name) {
	return new Slide({
		content : this.content['ppt/slides/' + name + '.xml'],
		presentation : this,
		name : name
	});
}
Presentation.prototype.insertAfterKey = function(object, afterKey, newKey,newContent){
	var res={}
	if (afterKey==null){
		res[newKey]=newContent;
	}
	for (var i=0; i<=Object.keys(object).length; i++){
		var key=Object.keys(object)[i];
		res[key]=object[key];
		if (afterKey==key){
			res[newKey]=newContent;
		}
	}
	return res;
}

Presentation.prototype.addSlide = function(layoutName) {
	var slideName = "slide" + (this.getSlideCount() + 1);

	var layoutKey = "ppt/slideLayouts/" + layoutName + ".xml";
	var slideKey = "ppt/slides/" + slideName + ".xml";
	var relsKey = "ppt/slides/_rels/" + slideName + ".xml.rels";

	// create slide
	// var slideContent = this.content[layoutKey]["p:sldLayout"];

	// var sld = this.content["ppt/slides/slide1.xml"]; // this is cheating,
	// copying an existing slide
	var sld = this.content[layoutKey]["p:sldLayout"];

	var slideContent = {
		"p:sld" : sld
	};

	slideContent = JSON.parse(JSON.stringify(slideContent));
	delete sld['$']["preserve"];
	delete sld['$']["type"];
	delete slideContent["p:sld"]['p:cSld']['0']['p:spTree']['0']['p:pic'];
	delete slideContent["p:sld"]['p:cSld']['0']['p:spTree']['0']['p:grpSp'];

	var shapes = slideContent["p:sld"]['p:cSld']['0']['p:spTree']['0']["p:sp"]
	if (shapes) {
		for (var i = 0; i <= shapes.length; i++) {
			var shape = shapes[i];
			
			if (shape) {
				var placeholder = shape["p:nvSpPr"]
				if (placeholder) {
					placeholder = placeholder['0'];
				}
				if (placeholder) {
					placeholder = placeholder["p:cNvSpPr"];
				}
				if (placeholder) {
					placeholder = placeholder['0'];
				}
				if (placeholder) {
					placeholder = placeholder["a:spLocks"];
				}
				if (placeholder) {
					placeholder = placeholder['0'];
				}
				if (placeholder) {
					placeholder = placeholder['$'];
				}
				if (placeholder) {
					placeholder = placeholder['noGrp'];
				}

				if (!placeholder || placeholder != 1) {

					slideContent["p:sld"]['p:cSld']['0']['p:spTree']['0']["p:sp"].splice(i,1);
				} 

			}
			
		};
		
	}
	this.content[slideKey] = slideContent; // { "p:sld": slideContent};

	var slide = new Slide({
		content : slideContent,
		presentation : this,
		name : slideName
	});

	// add to [Content Types].xml
	this.content["[Content_Types].xml"]["Types"]["Override"]
			.push({
				"$" : {
					"PartName" : "/ppt/slides/" + slideName + ".xml",
					"ContentType" : "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
				}
			});

	// add it rels to slidelayout
	this.content[relsKey] = {
		"Relationships" : {
			"$" : {
				"xmlns" : "http://schemas.openxmlformats.org/package/2006/relationships"
			},
			"Relationship" : [ {
				"$" : {
					"Id" : "rId1",
					"Type" : "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
					"Target" : "../slideLayouts/" + layoutName + ".xml"
				}
			} ]
		}
	};

	// add it to ppt/_rels/presentation.xml.rels
	var rId = "rId"
			+ (this.content["ppt/_rels/presentation.xml.rels"]["Relationships"]["Relationship"].length + 1);

	this.content["ppt/_rels/presentation.xml.rels"]["Relationships"]["Relationship"]
			.push({
				"$" : {
					"Id" : rId,
					"Type" : "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
					"Target" : "slides/" + slideName + ".xml"
				}
			});

	// add it to ppt/presentation.xml
	var maxId = 256;
	if (!Search.getPath(this.content,["ppt/presentation.xml","p:presentation","p:sldIdLst",0,"p:sldId"])){
		var newContent=[{
				"p:sldId":[]
		}]
		var res=this.insertAfterKey(this.content["ppt/presentation.xml"]["p:presentation"], "p:sldMasterIdLst", "p:sldIdLst",newContent);
		this.content["ppt/presentation.xml"]["p:presentation"]=res;
		
	}
	else{
		this.content["ppt/presentation.xml"]["p:presentation"]["p:sldIdLst"][0]["p:sldId"]
		.forEach(function(ob) {
			if (+ob["$"]["id"] > maxId)
				maxId = +ob["$"]["id"]
		})
		
	}
	this.content["ppt/presentation.xml"]["p:presentation"]["p:sldIdLst"][0]["p:sldId"]
			.push({
				"$" : {
					"id" : "" + (+maxId + 1),
					"r:id" : rId
				}
			});

	// increment slidecount
	var sldCount = this.getSlideCount();
	
	this.content["docProps/app.xml"]["Properties"]["Slides"][0] = sldCount;

	return slide;
}
Presentation.prototype.setTitle = function(title){
	this.content["docProps/core.xml"]["cp:coreProperties"]["dc:title"]=[title];
}
Presentation.prototype.setAuthor = function(author){
	this.content["docProps/core.xml"]["cp:coreProperties"]["dc:creator"]=[author];
} 
Presentation.prototype.setSubject = function(subject){
	this.content["docProps/core.xml"]["cp:coreProperties"]["dc:subject"]=[subject];
} 
Presentation.prototype.setCompany = function(company){
	this.content["docProps/app.xml"]["Properties"]["Company"]=[company];
} 

module.exports = Presentation;
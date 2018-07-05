var Shape = require('./shape');
var Chart = require('./chart');
var Img = require('./img');
var Search = require('./util/search');
var Text = require('./text');
var Table = require('./table');

// ======================================================================================================================
// Slide
// ======================================================================================================================

var Slide = function (args) {
	this.content = args.content;
	this.presentation = args.presentation;
	this.name = args.name;

	// TODO: Validate arguments
};

Slide.prototype.getShapes = function () {

	// TODO break out getShapeTree
	return this.content["p:sld"]["p:cSld"][0]["p:spTree"][0]['p:sp'].map(function (sp) {
		return new Shape(sp);
	});
};
Slide.prototype.addRel = function (type, options) {
	var rels = this.presentation.content['ppt/slides/_rels/' + this.name + '.xml.rels'];
	var numRels = rels["Relationships"]["Relationship"].length;
	var rId = "rId" + (numRels + 1);
	switch (type) {
		case "image":
			var relation = rels["Relationships"]["Relationship"].push({
				"$" : {
					"Id" : rId,
					"Type" : "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
					"Target" : options
				}
			});
			break;
		case "link":
			var relation = rels["Relationships"]["Relationship"].push({
				"$" : {
					"Id" : rId,
					"Type" : "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
					"TargetMode" : options.targetMode,
					"Target" : options.target
				}
			});
			break;

		default:
			break;
	}
	return rId;
}
Slide.prototype.addChart = function (data, done) {
	var self = this;
	var chartName = "chart" + (this.presentation.getChartCount() + 1);
	var chart = new Chart({
		slide : this,
		presentation : this.presentation,
		name : chartName
	});
	var slideName = this.name;

	chart.load(data, function (err, data) { // TODO pass it real data
		self.content["p:sld"]["p:cSld"][0]["p:spTree"][0]["p:graphicFrame"] = chart.content; // jsChartFrame["p:graphicFrame"];

		// Add entry to slide1.xml.rels
		// There should a slide-level and/or presentation-level
		// method to add/track
		// rels
		var rels = self.presentation.content['ppt/slides/_rels/' + slideName + '.xml.rels'];
		var numRels = rels["Relationships"]["Relationship"].length;
		var rId = "rId" + (numRels + 1);
		var numRels = rels["Relationships"]["Relationship"].push({
			"$" : {
				"Id" : rId,
				"Type" : "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
				"Target" : "../charts/" + chartName + ".xml"
			}
		});
		done(null, self);
	});
};

Slide.prototype.addShape = function () {
	var shape = new Shape();
	this.content["p:sld"]["p:cSld"][0]["p:spTree"][0]['p:sp'].push(shape.content);
	return shape;
};

Slide.prototype.getId = function () {
	var slideContent = this.content["p:sld"]["p:cSld"][0]["p:spTree"][0];
	var id = Search.countElements(slideContent, true);
	while (!Search.checkId(id, slideContent)) {
		id++;
	}
	return id;
};

/**
 * @desc adds a new image to a Slide
 * @param data:
 *            {String} Base64 string containing the image Data
 * @param position:
 *            {object} Element containing position coordinates: x: horizontal
 *            location, y: vertical location, w: width, h: height, type:
 *            Measuring unit (cm, inch or point)
 */
Slide.prototype.addImg = function (data, position,link) {

	var img = new Img({
		slide : this,
		presentation : this.presentation,
		data : data,
		position : position,
		link:link
	});

	return img;
};
/**
 * @desc adds a new Text to a Slide
 * @param textContent:
 *            {Object|array} containing options for new Text (see text.js for
 *            detailed description)
 * @param textFieldOptions:
 *            {Object} Required if multiple Lines in textContent (containing
 *            textField position)
 */
Slide.prototype.createText = function (textContent, textFieldOptions) {
	if (!textFieldOptions) {
		textFieldOptions = textContent;
	}
	var text = new Text({
		slide : this
	});
	text.createText(textContent, textFieldOptions);
};

/**
 * @desc adds a new Table to a Slide
 * @param args:
 *            {object} containing options for new Table (see table.js for
 *            detailed description)
 * @param args.position:
 *            {object} containing position coordinates: x: horizontal location,
 *            y: vertical location, w: width, h: height, type: Measuring unit
 *            (cm, inch or point)
 */
Slide.prototype.addTable = function (args) {
	args.slide = this;
	var table = new Table(args);
	table.addTable();
}
/**
 * @desc edit an existing text field
 * @param textContent:
 *            {String} Text that should be written in text field
 * @param type:
 *            {String} Name of the TextElement (By default Titel or
 *            Textplatzhalter)
 */
Slide.prototype.editTextContent = function (textContent, type) {
	var text = new Text({
		slide : this
	});
	text.editTextContent(textContent, type);
}

module.exports = Slide;
# js-pptx
Pure Javascript reader/writer/editor for PowerPoint, for use in Node.js.

**Table of Contents**
- [Design goals] (#design-goals)
- [Current status] (#current-status)
- [License] (#license)
- [Install] (#install)
- [Dependencies] (#dependencies)
- [Presentation: Usage and Options] (#presentation-usage-and-options)
    - [Create a Presentation] (#create-a-presentation)
    - [Adding a Slide] (#adding-a-slide)
    - [Adding Text] (#adding-text)
    - [Edit existing TextField] (#edit-existing-textfield)
    - [Adding Images] (#adding-images)
    - [Adding Tables] (#adding-tables)
    - [Adding Charts and Shapes] (#adding-charts-and-shapes)
- [Inspiration] (#inspiration)
- [Design Philosophy] (#design-philosophy)

# Design goals
* Read/edit/author PowerPoint .pptx files
* Pure Javascript with clean IP
* Run in browser and/or Node.js
* Friendly API for basic tasks, like text, shapes, charts, tables
* Access to raw XML for when you need to be very specific
* Rigorous testing


# Current status
It can currently:
 * read an existing PPTX file
 * retain all existing content
 * add slides, shapes, charts, text, images and images
 * edit existing text fields
 * save as a PPTX file


What it cannot yet do is:
 * Generate themes, layouts, masters, animations, etc.

# License
GNU General Public License (GPL)

# Install

In node.js
```
npm install protobi/js-pptx
```

In the browser: **(Not yet implemented)**
```
<script src="js-pptx.js"></script>
```

# Dependencies
* [xml2js](https://github.com/nfarina/xmldoc)
* [async](https://github.com/caolan/async)
* [jszip](https://stuk.github.io/jszip)

**************************************************************************************************
# Presentation: Usage and Options
PowerPoint presentations are created via JavaScript by following 4 basic steps:

1. Create a new Presentation
2. Add a Slide
3. Add one or more objects (Tables, Shapes, Images, Text and Media) to the Slide
4. Save the Presentation

**************************************************************************************************
## Create a Presentation
A Presentation represents a single pptx file. 
If you want to create a new presentation you first need an existing pptx file to use its master slides.
 
```javascript
var express = require('express');
var PPTX = require('../../lib/pptx');
var fs = require('fs');
var router = express.Router();
var path = require('path');

/* GET home page. */
router.get('/', function(req, res, next) {
	//read existing presentation
	var INFILE = '../test/files/TESTFILE.pptx';
	fs.readFile(INFILE, function(err, data) {
		if (err) throw err;
		//create new presentation
		var pptx = new PPTX.Presentation();
		pptx.load(data, function(err) {
			var slide=pptx.addSlide("SlideLayout1");
			slide.addImg(imgData,{x:1,y:2,w:3,h:1,type:"cm"});
			res.setHeader('Content-Disposition', 'attachment; filename=FILENAME.pptx');
			res.send(pptx.toBuffer());
			
		});
	});

});

module.exports = router;
```
**************************************************************************************************
## Presentation Properties
Several Presentation properties can be set:

```javascript
	pptx.setCompany("Company");
	pptx.setAuthor("J");
	pptx.setTitle("Analyse Company");
	pptx.setSubject("Annual Report");
```


**************************************************************************************************
## Adding a Slide
When adding a new slide, a LayoutName referring to a Master Slide Layout is required. 
The default name is slideLayout + the number of referring master slide. (located at ppt/slideLayouts)

```javascript
//Syntax
var slide = pptx.addSlide(slideLayout1);
```

**************************************************************************************************
## Adding Text
```javascript
// Syntax
slide.addText({text, {OPTIONS}});
slide.addText(
	[
		{text, {Options}},
		{text, {Options}},
		{text, {Options}}
	],
	{position}
)
```

| Option				| Type      | Default		| Description				| Possible Values							|
| :-----------------	| :---------| :----------	| :------------------------	| :---------------------------				|
| 'position.x'			| Numeric   | 				| horizontal location		| 0-n										|
| 'position.y'			| Numeric   | 				| vertical location			| 0-n										|
| 'position.w'			| Numeric   | 				| width						| 0-n										|
| 'position.h'			| Numeric   | 				| height					| 0-n										|
| 'position.type'		| Numeric   | "inch"		| measuring unit			| cm, inch or point(72 ppi)					|
| 'algn'				| String    | "l"			| text alignment			| ctr (center),l (left),r (right), dist (distributed), just (justified)|
| 'bold'				| String    | "0"			| bold text					| 0: false, 1:true							|
| 'breakLine			| String	| "0"			| appends a line break		| 0:false, 1:true							|
| 'color'				| String    | "000000"		| text color				| hex color code							|
| 'fill'				| String    | -				| fill color of text Shape	| hex color code							|
| 'italic'				| String    | "0"			| italic text				| 0: false, 1:true							|
| 'lang'				| String    | "de-DE"		| text language				| language setting (i.e.'en-US')			|
| 'lineColor'			| String    | -				| color of text shape border| hex color code							|
| 'lineSize'			| String    | -				| size of text shape border | line size in pt							|
| 'size'				| Numeric   | 18			| font size					| font size in pt							|
| 'typeface'			| String    | "Arial"		| font face					| font faces (i.e."Arial")					|
| 'underline'			| String    | "none"		| underline text			| sng (single Line), dbl (two Lines), dotted, dash|

**************************************************************************************************
## Edit existing TextField
Edit an existing TextField (i.e. from Master Slide)
```javascript
// Syntax
slide.editTextContent(newText,TextField);
```
| Option				| Type		| Default		| Description				| Possible Values				|
| :-----------------	| :---------| :----------	| :------------------------	| :---------------------------	|
| 'newText'				| String	| 				| new Text					| Any String					|
| 'TextField'			| String	| 				| Name of the TextField		| Valid Names					|

Text Fields are named by default Titel or Textplatzhalter 
Master Text Field names located in ppt/slideLayouts/slideLayout[X].xml inside the <p:cNvPr> Element (replace [X] with number)

Example:
```javascript
// Syntax
slide.editTextContent("Das ist der neue Text","Titel 1");
slide.editTextContent("Noch mehr Text","Textplatzhalter");
```
										
**************************************************************************************************

## Adding Images
```javascript
// Syntax
slide2.addImg(imgData,position);
```

| Option				| Type		| Default		| Description				| Possible Values							|
| :-----------------	| :---------| :----------	| :------------------------	| :---------------------------				|
| 'position.x'			| Numeric	| 				| horizontal location		| 0-n										|
| 'position.y'			| Numeric	| 				| vertical location			| 0-n										|
| 'position.w'			| Numeric	| 				| width						| 0-n										|
| 'position.h'			| Numeric	| 				| height					| 0-n										|
| 'position.type'		| Numeric	| "inch"		| measuring unit			| cm, inch or point(72 ppi)					|
| 'imgData'				| String	|				| Base64 String				| valid base64 string						|

**************************************************************************************************
## Adding tables
```javascript
// Syntax
var position={x:3, y:4.5,w:15,h:7,type:"cm"};
var rows=[];
rows.push([
			{text:"Zeile1", options:{color:"ff0000",fill:"ffff00",algn:"ctr",border:{pt:"30000",color:"f00c93"},colSpan:2,rowSpan:2}},
			{text:"Spalte 2", options:{color:"ff0000",border:{pt:"30000",color:"f00c93"}}},
]);
rows.push([
			{text:"Zeile 2", options:{color:"ff0000",border:{pt:"30000",color:"f00c93"}}},
			{text:"Spalte 2", options:{color:"ff0000",border:{pt:"30000",color:"f00c93"}}},
])
slide4.addTable({rows:rows, type:"cm", rowH:[3.5,4], colW:3,position:position});
```
**************************************************************************************************
### Table Options

| Option				| Type			| Default		| Description				| Possible Values							|
| :-----------------	| :---------	| :----------	| :------------------------	| :---------------------------				|
| 'position.x'			| Numeric		| 				| horizontal location		| 0-n										|
| 'position.y'			| Numeric		| 				| vertical location			| 0-n										|
| 'position.w'			| Numeric		| 				| width						| 0-n										|
| 'position.h'			| Numeric		| 				| height					| 0-n										|
| 'position.type'		| Numeric		| "inch"		| measuring unit			| cm, inch or point(72 ppi)					|
| 'colW'				| Numeric/Array	| 				| Column Width				| 0-n, or [values for columns]				|
| 'rowH'				| Numeric/Array	| 				| Row Height				| 0-n or [values for rows]					|
| 'type'				| String		| "000000"		| text color				| hex color code							|

**************************************************************************************************
### Table Cell Options

| Option				| Type		| Default		| Description				| Possible Values							|
| :-----------------	| :---------| :----------	| :------------------------	| :---------------------------				|
| 'algn'				| String	| "l"			| text alignment			| ctr (center),l (left),r (right), dist (distributed), just (justified|
| 'bold'				| String	| "0"			| bold text					| 0: false, 1:true							|
| 'border.pt'			| String	| master table 	| Border width in points	| 0-n										|
| 'border.color'		| String	| master table	| Border color				| hex color code 							|
| 'breakLine'			| String 	| "0"			| appends a line break		| 0:false, 1:true							|
| 'color'				| String	| "000000"		| text color				| hex color code							|
| 'colSpan'				| Numeric	| -				| Column Span				| 2-n										|
| 'fill'				| String	| master table	| Cell Fill					| hex color code							|
| 'italic'				| String	| "0"			| italic text				|  0: false, 1:true							|
| 'lang'				| String	| "de-DE"		| text language				| language setting (i.e.'en-US'				|
| 'rowSpan'				| Numeric	| -				| Column Span				| 2-n										|
| 'size'				| Numeric	| 18			| font size					| font size in pt							|
| 'typeface'			| String	| "Arial"		| font face					| font faces (i.e."Arial")					|
| 'underline'			| String	| "none"		| underline text			| sng (single Line), dbl (two Lines), dotted, dash|

**************************************************************************************************

##  Adding Charts and Shapes (from old library not yet used current lib)

```js
"use strict";

var fs = require("fs");
var PPTX = require('..');


var INFILE = './test/files/minimal.pptx'; // a blank PPTX file with my layouts, themes, masters.
var OUTFILE = '/tmp/example.pptx';

fs.readFile(INFILE, function (err, data) {
  if (err) throw err;
  var pptx = new PPTX.Presentation();
  pptx.load(data, function (err) {
    var slide1 = pptx.getSlide('slide1');

    var slide2 = pptx.addSlide("slideLayout3"); // section divider
    var slide3 = pptx.addSlide("slideLayout6"); // title only


    var triangle = slide1.addShape()
        .text("Triangle")
        .shapeProperties()
        .x(PPTX.emu.inch(2))
        .y(PPTX.emu.inch(2))
        .cx(PPTX.emu.inch(2))
        .cy(PPTX.emu.inch(2))
        .prstGeom('triangle');

    var triangle = slide1.addShape()
        .text("Ellipse")
        .shapeProperties()
        .x(PPTX.emu.inch(4))
        .y(PPTX.emu.inch(4))
        .cx(PPTX.emu.inch(2))
        .cy(PPTX.emu.inch(1))
        .prstGeom('ellipse');

    for (var i = 0; i < 20; i++) {
      slide2.addShape()
          .text("" + i)
          .shapeProperties()
          .x(PPTX.emu.inch((Math.random() * 10)))
          .y(PPTX.emu.inch((Math.random() * 6)))
          .cx(PPTX.emu.inch(1))
          .cy(PPTX.emu.inch(1))
          .prstGeom('ellipse');
    }

    slide1.getShapes()[3]
        .text("Now it's a trapezoid")
        .shapeProperties()
        .x(PPTX.emu.inch(1))
        .y(PPTX.emu.inch(1))
        .cx(PPTX.emu.inch(2))
        .cy(PPTX.emu.inch(0.75))
        .prstGeom('trapezoid');

    var chart = slide3.addChart(barChart, function (err, chart) {

      fs.writeFile(OUTFILE, pptx.toBuffer(), function (err) {
        if (err) throw err;
        console.log("open " + OUTFILE)
      });
    });
  });
});

var barChart = {
  title: 'Sample bar chart',
  renderType: 'bar',
  data: [
    {
      name: 'Series 1',
      labels: ['Category 1', 'Category 2', 'Category 3', 'Category 4'],
      values: [4.3, 2.5, 3.5, 4.5]
    },
    {
      name: 'Series 2',
      labels: ['Category 1', 'Category 2', 'Category 3', 'Category 4'],
      values: [2.4, 4.4, 1.8, 2.8]
    },
    {
      name: 'Series 3',
      labels: ['Category 1', 'Category 2', 'Category 3', 'Category 4'],
      values: [2.0, 2.0, 3.0, 5.0]
    }
  ]
}
```
**********************************************************************

# Inspiration
Inspired by [officegen](https://github.com/ZivBarber/officegen),
which creates pptx with text/shapes/images/tables/charts wonderfully (but does not read existing PPT files).

Also inspired by [js-xlsx](https://github.com/SheetJS/js-xlsx)
which reads/writes XLSX/XLS/XLSB, works in the browser and Node.js, and has an incredibly
thorough test suite (but does not read or write PowerPoint).

Motivated by desire to read and modify existing presentations, to inherit their themes, layouts and possibly content,
and work in the browser if possible.

https://github.com/protobi/js-pptx/wiki/API

# Design Philosophy
The design concept is to represent the Office document at two levels of abstraction:
* **Raw XML**  The actual complete OpenXML representation, in all its detail
* **Conceptual classes**  Simple Javascript classes that provide a convenient API

The conceptual classes provides a clear simple way to do common tasks, e.g. `Presentation().addSlide().addChart(data)`.

The raw API provides a way to do anything that the OpenXML allows, even if it's not yet in the conceptual classes, e.g.
e.g. `Presentation.getSlide(3).getShape(4).get('a:prstGeom').attr('prst', 'trapezoid')`


This solves a major dilemma in existing projects, which have many issue reports like "Please add this crucial feature to the API".
By being able to access the raw XML, all the features in OpenXML are available, while we make many of them more convenient.

The technical approach here uses:
* `JSZip` to unzip an existing `.pptx` file and zip it back,
* `xml2js` to convert the XML to Javascript and back to XML.

Converting to Javascript allows the content to be manipulated programmatically.  For each major entity, a Javascript class is created,
such as:
 * PPTX.Presentation
 * PPTX.Slide
 * PPTX.Shape
 * PPTX.spPr  // ShapeProperties
 * etc.

These classes allow properties to be set, and chained in a manner similar to d3 or jQuery.
The Javascript classes provide syntactic sugar, as a convenient way to query and modify the presentation.

But we can't possibly create a Javascript class that covers every entity and option defined in OpenXML.
So each of these classes exposes the  XML-to-Javascript object as a property `.content`, giving you theoretically
direct access to anything in the OpenXML standard, enabling you to take over
whenever the pre-defined features don't yet cover your particular use case.

It's up to you of course, to make sure that those changes convert to valid XML.  Debugging PPTX is a pain.

Right now, this uses English names for high-level constructs (e.g. `Presentation` and `Slide`),
but for lower level constructs uses names that directly mirror the OpenXML tagNames  (e.g.  `spPr` for ShapeProperties).

The challenge is it'll be a lot easier to extend the library if we follow the OpenXML tag names,
but the OpenXML tag names are so cryptic that they don't make great names for a Javascript library.

So we default to using the English name is used when returning objects even if the object has a cryptic class name, e.g.:
* `Slide.getShapes()` returns an array of `Shape` objects and
* `Shape.shapeProperties()` returns an `spPr` object.

Ideally would be consistent, and am working out which way to go.  Advice is welcome!

This library currently assumes it's starting from an existing presentation, and doesn't (yet) create one from scratch.
This allows you to use existing themes, styles and layouts.




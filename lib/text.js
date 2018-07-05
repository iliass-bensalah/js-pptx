var Search = require('./util/search');
var EMU = require('./util/emu');
var Paragraph = require('./fragments/js/paragraph');

var Text = module.exports = function(args) {
	this.slide = args.slide;

}

Text.prototype.editTextContent = function(text, type) {
	var shapes = Search.getPath(this.slide.content, [ "p:sld", "p:cSld", 0,
			"p:spTree", 0, "p:sp" ]);
	if (shapes) {
		for (var i = 0; i < shapes.length; i++) {
			var name = Search.getPath(shapes[i], [ "p:nvSpPr", 0, "p:cNvPr", 0,
					"$", "name" ]);
			if (name) {
				if (type == name.substr(0, type.length)) {
					this.overrideText(text, shapes[i]);
					
				}
			}
		}
	}

}
Text.prototype.overrideText = function(text, element) {
	if (Search.checkPath(element, [ "p:txBody", 0, "a:p" ])) {
		element["p:txBody"][0]["a:p"] = this.defaultTextField(text);

	}
}
Text.prototype.createText = function (textContent, textField){
	textField.id=this.slide.getId();
	textField.name="Textfeld "+textField.id
	textField.position=EMU.calculateEMU(textField.position)
		
	var res=this.defaultShape(textContent,textField);
	if (!Search.checkPath(this.slide.content,["p:sld","p:cSld",0,"p:spTree",0,"p:sp"])){
		this.slide.content["p:sld"]["p:cSld"][0]["p:spTree"][0]["p:sp"]=[];
	}
	this.slide.content["p:sld"]["p:cSld"][0]["p:spTree"][0]["p:sp"].push(res);
	
}

Text.prototype.defaultTextField = function(text) {
	return [ {
		"a:r" : [ {
			"a:rPr" : [ {
				"$" : {
					"dirty" : "0",
					"lang" : "de-DE"
				}
			} ],
			"a:t" : [ text ]
		} ]
	} ]
}

Text.prototype.defaultShape = function (textContent,textField) {
	if (!textContent.options){
		textContent.options={}
	}

	var res= {
		"p:nvSpPr":[{
			"p:cNvPr":[{
				"$":{
					"id":textField.id,
					"name":textField.name
				}
			}],
			"p:cNvSpPr":[{
				"a:spLocks":[{
					"$":{
						"noGrp":"1"
					}
				}]
				
			}],
			"p:nvPr":[{
				
			}]
		}],
		"p:spPr":[{
			"a:xfrm":[{
				"a:off":[{
					"$":{
						"x":textField.position.x,
						"y":textField.position.y
					}
				}],
				"a:ext":[{
					"$":{
						"cx":textField.position.w,
						"cy":textField.position.h
					}
				}]
			}],
			"a:prstGeom":[{
				"$":{
					"prst": "rect"
				},
				"a:avLst":[""]
			}]
		}],
		"p:txBody":[{
			"a:bodyPr":[""],//TODO BODY PROPS
			"a:lstStyle":[""], // TODO LstStyle
			"a:p":[
				
			]

		}]
	};
	if (textContent.options.fill){
		res["p:spPr"][0]["a:solidFill"] = [ {
			"a:srgbClr" : [ {
				"$" : {
					"val" : textContent.options.fill
				}
			} ]
		} ]
	}
	if (textContent.options.lineSize || textContent.options.lineColor) {
		if (!textContent.options.lineSize){
			textContent.options.lineSize = "2"
		}
		if (!textContent.options.lineColor){
			textContent.options.lineColor = "000000"
		}
		res["p:spPr"][0]["a:ln"] = [ {
			"$" : {
				"w" : (textContent.options.lineSize * 12700).toString()
			},
			"a:solidFill" : [ {
				"a:srgbClr" : [ {
					"$" : {
						"val" : textContent.options.lineColor
					}
				} ]
			} ]
		} ]
	}
	res["p:txBody"][0]["a:p"]=Paragraph.createNewParagraphs(textContent);

	return res;
}

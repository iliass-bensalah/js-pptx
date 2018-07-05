var EMU = require('./util/emu');
var ChangeOrder=require('./util/changeOrder')

var Img = module.exports = function(args) {
	this.presentation = args.presentation;
	this.slide = args.slide;
	//get imgData containing base64 data + imgExtension
	this.imgData = this.decodeBase64Image(args.data);
	this.ext = this.imgData.type;
	this.position = EMU.calculateEMU(args.position);
	//get Hyperlink if set
	if (args.link){
		this.hyperlink=args.link
	}
	//set imageName
	this.imgName = "image" + (this.presentation.getImgCount() + 1) + "." + this.ext;
	this.saveImg();
}


//save image and add rels
Img.prototype.saveImg = function() {
	//get Img Data and save to media folder
	this.presentation.content["ppt/media/" + this.imgName] = this.imgData.data;
	//add relation
	var rId = this.slide.addRel("image", "../media/" + this.imgName);
	//add hyperlink if set
	var hyperId;
	if (this.hyperlink){
		hyperId=this.slide.addRel("link", this.hyperlink)
	}
	//add slide content
	this.addSlideContent(rId,hyperId);
}

Img.prototype.decodeBase64Image = function(dataString) {
	var matches = dataString.match(/^data:([A-Za-z-+\/]+);base64,(.+)$/), response = {};

	if (matches.length !== 3 || matches[1].length <= 6) {
		return new Error('Invalid input string');
	}

	response.type = matches[1].substring(6);

	response.data = new Buffer(matches[2], 'base64').toString('binary');

	return response;
}

Img.prototype.addSlideContent = function(rId,hyperLinkId) {
	if (!this.slide.content["p:sld"]["p:cSld"][0]["p:spTree"][0]["p:pic"]){
		this.slide.content["p:sld"]["p:cSld"][0]["p:spTree"][0]["p:pic"]=[];
		this.slide.content["p:sld"]["p:cSld"][0]["p:spTree"][0]=ChangeOrder.moveObjectElement("p:pic","p:grpSpPr",this.slide.content["p:sld"]["p:cSld"][0]["p:spTree"][0]);
	}
	
	
	id=this.slide.getId();
	var contentImg = {
		"p:nvPicPr" : [{
				"p:cNvPr":[{
					"$":{
						"name":this.imgName,
						"id":id,
						"descr":this.imgName
					}
				}],
				"p:cNvPicPr":[{
					"a:picLocks":[{
						"$":{
							"noChangeAspect":"1"
						}
					}]
				}],
				"p:nvPr":[""]
			}],
		"p:blipFill":[{
			"a:blip":[{
				"$":{
					"r:embed": rId,
					"cstate": "print"
				}
			}],
			"a:stretch":[{
					"a:fillRect":[""]
			}]
				
		}],
		"p:spPr":[{
			"a:xfrm":[{
				"a:off":[{
					"$":{
						"x":this.position.x,
						"y":this.position.y
					}
				}],
				"a:ext":[{
					"$":{
						"cx":this.position.w,
						"cy":this.position.h
					}
				}]
				
			}],
			"a:prstGeom":[{
				"$":{
					"prst":"rect"
				},
				"a:avLst":[""]
			}]
		}]
	};
	//if hyperLink is set add Relation
	if (hyperLinkId){
		contentImg["p:nvPicPr"][0]["p:cNvPr"][0]["a:hlinkClick"]=[{
				"$":{
					"r:id":hyperLinkId
				}
			}]
		}


	this.slide.content["p:sld"]["p:cSld"][0]["p:spTree"][0]["p:pic"].push(contentImg);
	
}

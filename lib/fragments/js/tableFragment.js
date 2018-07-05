//var fs=require("fs");
//var xml2js=require("xml2js");
var TableStyle=require("./tableStyle");
var Paragraph=require("./paragraph");
module.exports = {
		createEmptyTable: function(name,id,position,styleId){
		
			
			return{
				"p:nvGraphicFramePr":[{
						"p:cNvPr":[{
							"$":{
								"name":name,
								"id":id
							}
						}],
						"p:cNvGraphicFramePr":[{
							"a:graphicFrameLocks":[{
								"$":{
									"noGrp":"1"
								}
							}]
						}],
						"p:nvPr":[""]
					}],
					"p:xfrm":[{
						"a:off":[{
							"$":{
								"x":position.x,
								"y":position.y
							}
						}],
						"a:ext":[{
							"$":{
								"cx":position.w,
								"cy":position.h
							}
						}]
					}],
					"a:graphic":[{
						"a:graphicData":[{
							"$":{
								"uri":"http://schemas.openxmlformats.org/drawingml/2006/table"
							},
							"a:tbl":[{
								"a:tblPr":[{
									"$":{
										"firstRow":"1",
										"bandRow":"1"
									},
									"a:tableStyleId":[
										styleId
									]								
								}]
							}]							
						}]
					}]
						
			}
		},
		defaultTableStyle: function (){
			
			return TableStyle;
//			var parser=new xml2js.Parser();
//			fs.readFile(__dirname+'/../xml/tableStyle.xml',function(err,data){
//				parser.parseString(data, function (err,result){
//					result=JSON.stringify(result,null,4);
//					fs.writeFile('helloworld.txt', result, function (err) {
//						  if (err) return console.log(err);
//						  console.log('Hello World > helloworld.txt');
//						});
//					return result;
//				})
//			})
		},
		createRow:function(rowH){
			return{
				"$":{
					"h":rowH
				},
				"a:tc":[
					
				]
				
			}
		},
		
		createCell:function(cellOptions){
			
		
			var res={ 
				"$":{},
				"a:txBody":[{
					"a:bodyPr":[""],
					"a:lstStyle":[""],
					"a:p":[
						
					]
				}],
				"a:tcPr":[{
					
				}]
			
			};
			res["a:txBody"][0]["a:p"]=Paragraph.createNewParagraphs(cellOptions);
			
			if (!cellOptions.options){
				cellOptions.options={};
			}
			//add border options if set
			if (cellOptions.options.border){
				if (!cellOptions.options.border.pt)	cellOptions.options.border.pt="12700";
				if (!cellOptions.options.border.color) cellOptions.options.border.color="ffffff";
				//left border
				res["a:tcPr"][0]["a:lnL"]=this.createBorderElem(cellOptions.options.border);
				//right border
				res["a:tcPr"][0]["a:lnR"]=this.createBorderElem(cellOptions.options.border);
				//top border
				res["a:tcPr"][0]["a:lnT"]=this.createBorderElem(cellOptions.options.border);
				//bottom border
				res["a:tcPr"][0]["a:lnB"]=this.createBorderElem(cellOptions.options.border);
				
			}
			//add fill 
			if (cellOptions.options.fill) {
				colorElem=this.createColorElem(cellOptions.options.fill);
				res["a:tcPr"][0]["a:solidFill"]=colorElem;
			}
			//Add Span Options 
			if (cellOptions.options.colSpan) res["$"]["gridSpan"]=cellOptions.options.colSpan;
			if (cellOptions.options.rowSpan) res["$"]["rowSpan"]=cellOptions.options.rowSpan;
			if (cellOptions.options.vMerge) res["$"]["vMerge"]=cellOptions.options.vMerge;
			if (cellOptions.options.hMerge) res["$"]["hMerge"]=cellOptions.options.hMerge;
			return res;
		},
		createBorderElem(options){
			//default values
			if (!options) options={};
			if (!options.pt) options.pt="12700";
			if (!options.color) options.color="ffffff";
			var res= [{
				"$":{
					"w":options.pt,
					"algn":"ctr",
					"cmpd":"sng",
					"cap":"flat"
				},
				"a:solidFill":[""],
				"a:prstDash":[{
					"$":{
						"val":"solid"
					}
				}],
				"a:round":[""],
				"a:headEnd":[{
					"$":{
						"w":"med",
						"len":"med",
						"type":"none"
					}
				}],
				"a:tailEnd":[{
					"$":{
						"w":"med",
						"len":"med",
						"type":"none"
					}					
				}]
				
			}];
			res[0]["a:solidFill"]=this.createColorElem(options.color);
			return res;
			
			
		},
		createColorElem(color) {
			return[{
					"a:srgbClr":[{
						"$":{
							"val": color
						}
					}]
				}];
			
		},
		createColumn: function (width){
			return{
				"$":{
					"w":width
				}
			}
		}
		
		
}
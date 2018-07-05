module.exports = {

	/**
	 * @desc return array with different paragraphs (if line break option is
	 *       true, a new paragraph has to be created).
	 * @param args:
	 *            {Object | Array}Text input from user. Either an Object if
	 *            only single text line (args.text+args.options) or an array
	 *            of object if multiple lines
	 * 
	 */
	formatParagraphs(aText){
		
		// if aText is an Object create arrays and return result
		if (!(Object.prototype.toString.call(aText) == '[object Array]')){
			var res=[];
			var paragraph=[]
			paragraph.push(aText);
			res.push(paragraph);
			return res;
		}
		
		// Create an array entry for every newLine
		var aParagraphs=[];
		var curr=[];
		for (var i=0; i<aText.length; i++){
			curr.push(aText[i]);
			if (aText[i].options && aText[i].options.breakLine=="1"){
				aParagraphs.push(curr);
				curr=[];
				
			}
			
		}
		if (curr.length!=0){
			aParagraphs.push(curr);
		}
		return aParagraphs;
	},
	/**
	 * @desc: creates new Paragrpahs
	 * @param {Object|Array}textContent:
	 *            Text input from user. Either an Object if only single text
	 *            line (args.text+args.options) or an array of object if
	 *            multiple lines
	 * 
	 * 
	 */		
	createNewParagraphs(textContent){
		var paragraphs=this.formatParagraphs(textContent);
		var res =[];
		for (var i=0; i<paragraphs.length; i++){
			res[i]=this.createSingleParagraph(paragraphs[i]);
			
		}
		return res;
		
	},
	/**
	 * @desc: creates a single Paragraph
	 * @param paragraph:
	 *            {Array} containing Text with different Format. For each
	 *            entry create a new a:r element with according formats
	 */
	createSingleParagraph(paragraph){
		if (!paragraph[0].options) paragraph[0].options={};
		if (!paragraph[0].options.algn) paragraph[0].options.algn="l";
		if (!paragraph[0].options.level) paragraph[0].options.level="0";
		var res={
				"a:pPr":[{
						"$":{
							"algn":paragraph[0].options.algn,
							"lvl": paragraph[0].options.level
						}
				}],
				"a:r":[
					
				]
		};
	
		for(var i=0; i<paragraph.length; i++){
						
			res["a:r"].push(this.createRow(paragraph[i]));
		}
		return res;
	},
	/**
	 * @desc: creates a single Row Element inside a paragraph
	 * @param args:
	 *            {Object} containing options for text formatting
	 * @param args.options.lang:
	 *            {String} language setting (i.e.'en-US-)
	 * @param args.options.italic:
	 *            {String} 0:true 1:false
	 * @param args.options.bold:
	 *            {String}  0:true 1:false
	 * @param args.options.underline:
	 *            {String} underline type (sng,dash,dotted...)
	 * @param args.options.size:
	 *            {String} font size in pt
	 * @param args.options.color:
	 *            {String} font color in hex
	 * @param args.options.typeface:
	 *            {String} typeface (i.e.'Arial')
	 * @param args.options.algn:
	 *            {String} l,r,ctr,dist,just ...
	 * @param args.text
	 *            {String} text content
	 * 
	 */
	createRow(args){
		if (!args.options){
			args.options={};
		}
		
		if (!args.options.lang) args.options.lang="de-DE";
		if (!args.options.italic) args.options.italic="0";
		if (!args.options.bold) args.options.bold="0";
		if (!args.options.underline) args.options.underline="none";
		if (!args.options.size) args.options.size=18;
		
		if (!args.options.color) args.options.color="000000";
		if (!args.options.typeface) args.options.typeface="Arial";
		if (!args.options.algn) args.options.algn="l";
		if (!args.text) args.text="";	
		var res={
			"a:rPr":[{
					"$":{
						"lang":args.options.lang,
						"dirty":"0",
						"i":args.options.italic,
						"b":args.options.bold,
						"u":args.options.underline,
						"sz":args.options.size*100,
											
					},
					"a:solidFill":[{
						"a:srgbClr":[{
							"$":{
								"val":args.options.color
							}
						}]
					}],
					"a:latin":[{
						"$":{
							"charset":0, // ANSI character set
							"pitchFamily":0, // Default pitch +
												// unknown font family
							"typeface":args.options.typeface
						}
					}]
				}],
				"a:t":[
					args.text
				]
		}
		return res;
	}
	
	
}
var Search = require('./util/search');
var EMU = require('./util/emu');
var TableFragment= require('./fragments/js/tableFragment');
	

/**
 * @desc Save variables in Table Object
 * @param agrs.slide:
 *            {Object} reference to the associated slide
 * @param rowhH:
 *            Array containing different row heights or one number if all row
 *            heights are equal
 * @param colW:
 *            Array containing different row heights or one number if all row
 *            heights are equal
 * @param type:
 *            {String} measuring unit ("cm" or "inch")
 * @param args.rows:
 *            {Array} containing the rows of the table [row1, row2, ...]
 * @type row: {Object} each row contains its cells {cell1, cell2, ...}
 * @type cell: {Object} each cell contains a Text and Options Object
 *       {text:"Example",options{color:"ff0000"}}
 * 
 * 
 */
var Table = module.exports = function(args) {
	this.slide = args.slide;
	this.rows=args.rows;
	this.numberOfColumns=this.getNumberOfColumns(this.rows);
	this.rowH=args.rowH;
	this.colW=args.colW;
	this.type=args.type;
	this.position=EMU.calculateEMU(args.position);
	this.id=this.slide.getId();
	this.name="Tabelle "+this.id;
};
Table.prototype.getNumberOfColumns=function(aRows){
	var columns=0;
	for (var i=0; i<aRows.length; i++){
		var currColumns=aRows[i].length;
		columns=Math.max(currColumns,columns);
	}
	return columns;
};
/**
 * @desc adds a table to the current slide
 *  
 */
Table.prototype.addTable = function(){
	//create new graphicFrame parent element if needed
	if (!this.slide.content["p:sld"]["p:cSld"][0]["p:spTree"][0]["p:graphicFrame"]){
		this.slide.content["p:sld"]["p:cSld"][0]["p:spTree"][0]["p:graphicFrame"]=[];
	}
	//get current ID
	var styleId=this.getTableStyle();
	//create an empty table 
	var oTable=TableFragment.createEmptyTable(this.name,this.id,this.position,styleId);
	//add Columns and Rows
	this.addColumns(oTable);
	this.addRows(oTable,this.rows);
	//add Table to current Slide
	this.slide.content["p:sld"]["p:cSld"][0]["p:spTree"][0]["p:graphicFrame"].push(oTable);
};

/**
 * @desc get the reference id of the current table style, if no style exists, create default one
 */
Table.prototype.getTableStyle = function () {
	var res=Search.getPath(this.slide.presentation.content,["ppt/tableStyles.xml","a:tblStyleLst","a:tblStyle",0,"$","styleId"]);
	if (res){
		return res;
	}
	else{
		this.slide.presentation.content["ppt/tableStyles.xml"]["a:tblStyleLst"]["a:tblStyle"]=TableFragment.defaultTableStyle(this);
		return Search.getPath(this.slide.presentation.content,["ppt/tableStyles.xml","a:tblStyleLst","a:tblStyle",0,"$","styleId"]);
	}
};
/**
 * @desc add Table Rows
 */
Table.prototype.addRows = function(oTable){
	oTable["a:graphic"][0]["a:graphicData"][0]["a:tbl"][0]["a:tr"]=[];
	
	for (var i=0; i<this.rows.length; i++){

		var currHeight=this.rowH;
		if (typeof this.rowH === 'object'){
			currHeight=this.rowH[i];
		}
		currHeight=Math.floor(currHeight*EMU.getMultiplier(this.type))
		var row=TableFragment.createRow(currHeight);
		
		
		for (var x=0; x<this.rows[i].length; x++){
			var cell= this.rows[i][x];
			//add Span if set
			this.addSpan(cell,this.rows,i,x);
			//create cells 
			var tc = TableFragment.createCell(cell);
			row["a:tc"].push(tc);
		}
		oTable["a:graphic"][0]["a:graphicData"][0]["a:tbl"][0]["a:tr"].push(row);
	}
};
/**
 * @desc add Table Columns
 */
Table.prototype.addColumns = function (oTable){
	oTable["a:graphic"][0]["a:graphicData"][0]["a:tbl"][0]["a:tblGrid"]=[{
		"a:gridCol":[]
	}];
	for (var i=0; i<this.numberOfColumns; i++){
		var currWidth=this.colW;
		if (typeof this.colW === 'object' ){
			currWidth=colW[i];
		}
		currWidth=Math.floor(currWidth*EMU.getMultiplier(this.type));
		var gridCol=TableFragment.createColumn(currWidth);
		oTable["a:graphic"][0]["a:graphicData"][0]["a:tbl"][0]["a:tblGrid"][0]["a:gridCol"].push(gridCol);
	}
};
/**
 * @desc if row or colSpan is set, set Merge attribute in
 *       following cells
 */
Table.prototype.addSpan=function(cell,rows,rowNumber,cellNumber){
	if (cell.options && cell.options.colSpan>1){
		for (var i=1; i<cell.options.colSpan; i++){
			rows[rowNumber][cellNumber+i].hMerge="1";
		}
	}
	if (cell.options && cell.options.rowSpan>1){
		for (var i=1; i<cell.options.gridSpan; i++){
			rows[rowNumber+i][cellNumber].vMerge="1";
		}
	}
};



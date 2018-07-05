/**
 * 
 */

module.exports ={
		
		calculateEMU : function(position) {
			position=JSON.parse(JSON.stringify(position))
			//check if arguments are valid
			var x=position.x;
			var y=position.y;
			var width=position.w;
			var height=position.h;
			var type=position.type;
			if (x==null || y==null || width==null || height==null ){
				throw new Error();
			}
			
			var multiplier=this.getMultiplier(type);
			
			position.x=Math.floor(x*multiplier);
			position.y=Math.floor(y*multiplier);
			position.w=Math.floor(width*multiplier);
			position.h=Math.floor(height*multiplier);
			return position;
		},
		
		getMultiplier : function (type){
			var multiplier;
			switch (type){
				case "inch":
					multiplier=914400;
					break;
				case "point":
					multiplier=914400/72;
					break;
				case "cm":
					multiplier=360000;
					break;
				default:
					multiplier=914400;
			}
			return multiplier;
		}
		
		
}
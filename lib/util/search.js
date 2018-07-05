module.exports = {

	countElements : function(element, countSubElements) {
		var length = Object.keys(element).length;
		if (countSubElements) {
			var res = length;
			for (var i = 0; i < length; i++) {
				var subElements = element[Object.keys(element)[i]];
				var subElementsLength = Object.keys(subElements).length;

				if (subElementsLength > 1) {
					res += subElementsLength - 1;
				}
			}
			return res;
		} else {
			return length;
		}

	},
	checkId : function(id, element) {
		var length = Object.keys(element).length;
		for (var i = 0; i < length; i++) {
			var key = Object.keys(element)[i];
			var subElement = element[key];
			for (var z = 0; z < subElement.length; z++) {
				var curr = subElement[z];
				if (curr == null) {
					continue;
				} else {
					switch (key) {
					case "p:pic":
						if (this.checkPath(curr,["p:nvPicPr",0])) {
							curr = curr["p:nvPicPr"][0];
						}
						break;
					case "p:sp":
						if (this.checkPath(curr,["p:nvSpPr",0])) {
							curr = curr["p:nvSpPr"][0];
						}
						break;
					
					case "p:graphicFrame":
						var path=this.getPath(curr,["p:nvGraphicFramePr",0]);
						if (path){
							curr=path;
						}
						break;
					}
				}
				if (curr["p:cNvPr"]) {
					var res = this.checkElementIdProperty(curr["p:cNvPr"], id)
					if (!res) {
						return false;
					}

				}

			}
		}
		return true;
	},
	checkElementIdProperty : function(element, ID) {
		if (this.checkPath(element, [ 0, "$" ])) {
			
			if (element[0]["$"].id == ID) {
				return false;
			}

		}
		return true;
	},

	checkPath : function(object, steps) {
		if (!object) {
			return false;
		}
		if (steps.length == 0) {
			return true;
		}
		var nextStep = steps[0];
		if (object[nextStep] == null) {
			return false;
		}
		steps.shift();
		return this.checkPath(object[nextStep], steps);
	},
	getPath : function(object, steps){
		if (!object) {
			return null;
		}
		if (steps.length == 0) {
			return object;
		}
		var nextStep = steps[0];
		if (object[nextStep] == null) {
			return null;
		}
		steps.shift();
		return this.getPath(object[nextStep], steps);
	}

}
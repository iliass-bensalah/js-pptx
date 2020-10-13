var XmlNode = require('./xmlnode');

var shapeProperties = require('./shapeProperties');
var defaults = require('./defaults');
var clone = require('./util/clone')

//======================================================================================================================
// Shape (p:sp)
//======================================================================================================================

var Shape = function (content) {
  this.content = content || clone(defaults["p:sp"]);
}

Shape.prototype.getText = function () {
  return this.content['p:txBody'][0]['a:p'][0]['a:r'][0]['a:t'][0];
}

Shape.prototype.setText = function (text) {
    this.content['p:txBody'][0]['a:p'][0]['a:r'][0]['a:t'][0] = text;
    return this;
}

Shape.prototype.shapeProperties = function () {
  return new shapeProperties(this.content["p:spPr"][0])
}

module.exports = Shape;
"use strict";

function _defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } }
function _createClass(Constructor, protoProps, staticProps) { if (protoProps) _defineProperties(Constructor.prototype, protoProps); if (staticProps) _defineProperties(Constructor, staticProps); Object.defineProperty(Constructor, "prototype", { writable: false }); return Constructor; }
function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var _require = require('../../../../uuid'),
  uuid = _require.v4;
var utils = require('../utils');

// §18.7.3 Comment
var Comment = /*#__PURE__*/_createClass(function Comment(ref, comment) {
  var options = arguments.length > 2 && arguments[2] !== undefined ? arguments[2] : {};
  _classCallCheck(this, Comment);
  this.ref = ref;
  this.comment = comment;
  this.uuid = '{' + uuid().toUpperCase() + '}';
  this.row = utils.getExcelRowCol(ref).row;
  this.col = utils.getExcelRowCol(ref).col;
  this.marginLeft = options.marginLeft || this.col * 88 + 8 + 'pt';
  this.marginTop = options.marginTop || (this.row - 1) * 16 + 8 + 'pt';
  this.width = options.width || '104pt';
  this.height = options.height || '69pt';
  this.position = options.position || 'absolute';
  this.zIndex = options.zIndex || '1';
  this.fillColor = options.fillColor || '#ffffe1';
  this.visibility = options.visibility || 'hidden';
});
module.exports = Comment;
//# sourceMappingURL=comment.js.map
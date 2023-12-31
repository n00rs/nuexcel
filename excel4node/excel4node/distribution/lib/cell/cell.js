"use strict";

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }
function _defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } }
function _createClass(Constructor, protoProps, staticProps) { if (protoProps) _defineProperties(Constructor.prototype, protoProps); if (staticProps) _defineProperties(Constructor, staticProps); Object.defineProperty(Constructor, "prototype", { writable: false }); return Constructor; }
var utils = require('../utils.js');
var Comment = require('../classes/comment');

// §18.3.1.4 c (Cell)
var Cell = /*#__PURE__*/function () {
  /**
  * Create an Excel Cell
  * @private
  * @param {Number} row Row of cell.
  * @param {Number} col Column of cell
  */
  function Cell(row, col) {
    _classCallCheck(this, Cell);
    if (row <= 0) throw 'Row parameter must not be zero or negative.';
    if (col <= 0) throw 'Col parameter must not be zero or negative.';
    this.r = "".concat(utils.getExcelAlpha(col)).concat(row); // 'r' attribute
    this.s = 0; // 's' attribute refering to style index
    this.t = null; // 't' attribute stating Cell data type - §18.18.11 ST_CellType (Cell Type)
    this.f = null; // 'f' child element used for formulas
    this.v = null; // 'v' child element for values
    this.row = row; // used internally throughout code. Does not go into XML
    this.col = col; // used internally throughout code. Does not go into XML
  }
  _createClass(Cell, [{
    key: "comment",
    get: function get() {
      return this.comments[this.r];
    }
  }, {
    key: "string",
    value: function string(index) {
      this.t = 's';
      this.v = index;
      this.f = null;
    }
  }, {
    key: "number",
    value: function number(val) {
      this.t = 'n';
      this.v = val;
      this.f = null;
    }
  }, {
    key: "formula",
    value: function formula(_formula) {
      this.t = null;
      this.v = null;
      this.f = _formula;
    }
  }, {
    key: "bool",
    value: function bool(val) {
      this.t = 'b';
      this.v = val;
      this.f = null;
    }
  }, {
    key: "date",
    value: function date(dt) {
      this.t = null;
      this.v = utils.getExcelTS(dt);
      this.f = null;
    }
  }, {
    key: "style",
    value: function style(sId) {
      this.s = sId;
    }
  }, {
    key: "addToXMLele",
    value: function addToXMLele(ele) {
      if (this.v === null && this.is === null) {
        return;
      }
      var cEle = ele.ele('c').att('r', this.r).att('s', this.s);
      if (this.t !== null) {
        cEle.att('t', this.t);
      }
      if (this.f !== null) {
        cEle.ele('f').txt(this.f).up();
      }
      if (this.v !== null) {
        cEle.ele('v').txt(this.v).up();
      }
      cEle.up();
    }
  }]);
  return Cell;
}();
module.exports = Cell;
//# sourceMappingURL=cell.js.map
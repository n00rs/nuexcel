"use strict";

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }
function _defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } }
function _createClass(Constructor, protoProps, staticProps) { if (protoProps) _defineProperties(Constructor.prototype, protoProps); if (staticProps) _defineProperties(Constructor, staticProps); Object.defineProperty(Constructor, "prototype", { writable: false }); return Constructor; }
var _isEqual = require('../../../../lodash.isequal');
var Style = require('../style');
var util = require('util');
var DXFItem = /*#__PURE__*/function () {
  // §18.8.14 dxf (Formatting)
  function DXFItem(style, wb) {
    _classCallCheck(this, DXFItem);
    this.wb = wb;
    this.style = style;
    this.id;
  }
  _createClass(DXFItem, [{
    key: "dxfId",
    get: function get() {
      return this.id;
    }
  }, {
    key: "addToXMLele",
    value: function addToXMLele(ele) {
      this.style.addDXFtoXMLele(ele);
    }
  }]);
  return DXFItem;
}();
var DXFCollection = /*#__PURE__*/function () {
  // §18.8.15 dxfs (Formats)
  function DXFCollection(wb) {
    _classCallCheck(this, DXFCollection);
    this.wb = wb;
    this.items = [];
  }
  _createClass(DXFCollection, [{
    key: "add",
    value: function add(style) {
      if (!(style instanceof Style)) {
        style = this.wb.Style(style);
      }
      var thisItem;
      this.items.forEach(function (item) {
        if (_isEqual(item.style.toObject(), style.toObject())) {
          return thisItem = item;
        }
      });
      if (!thisItem) {
        thisItem = new DXFItem(style, this.wb);
        this.items.push(thisItem);
        thisItem.id = this.items.length - 1;
      }
      return thisItem;
    }
  }, {
    key: "length",
    get: function get() {
      return this.items.length;
    }
  }, {
    key: "addToXMLele",
    value: function addToXMLele(ele) {
      var dxfXML = ele.ele('dxfs').att('count', this.length);
      this.items.forEach(function (item) {
        item.addToXMLele(dxfXML);
      });
    }
  }]);
  return DXFCollection;
}();
module.exports = DXFCollection;
//# sourceMappingURL=dxfCollection.js.map
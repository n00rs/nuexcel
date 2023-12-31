"use strict";

function _typeof(obj) { "@babel/helpers - typeof"; return _typeof = "function" == typeof Symbol && "symbol" == typeof Symbol.iterator ? function (obj) { return typeof obj; } : function (obj) { return obj && "function" == typeof Symbol && obj.constructor === Symbol && obj !== Symbol.prototype ? "symbol" : typeof obj; }, _typeof(obj); }
function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }
function _defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } }
function _createClass(Constructor, protoProps, staticProps) { if (protoProps) _defineProperties(Constructor.prototype, protoProps); if (staticProps) _defineProperties(Constructor, staticProps); Object.defineProperty(Constructor, "prototype", { writable: false }); return Constructor; }
function _inherits(subClass, superClass) { if (typeof superClass !== "function" && superClass !== null) { throw new TypeError("Super expression must either be null or a function"); } subClass.prototype = Object.create(superClass && superClass.prototype, { constructor: { value: subClass, writable: true, configurable: true } }); Object.defineProperty(subClass, "prototype", { writable: false }); if (superClass) _setPrototypeOf(subClass, superClass); }
function _setPrototypeOf(o, p) { _setPrototypeOf = Object.setPrototypeOf ? Object.setPrototypeOf.bind() : function _setPrototypeOf(o, p) { o.__proto__ = p; return o; }; return _setPrototypeOf(o, p); }
function _createSuper(Derived) { var hasNativeReflectConstruct = _isNativeReflectConstruct(); return function _createSuperInternal() { var Super = _getPrototypeOf(Derived), result; if (hasNativeReflectConstruct) { var NewTarget = _getPrototypeOf(this).constructor; result = Reflect.construct(Super, arguments, NewTarget); } else { result = Super.apply(this, arguments); } return _possibleConstructorReturn(this, result); }; }
function _possibleConstructorReturn(self, call) { if (call && (_typeof(call) === "object" || typeof call === "function")) { return call; } else if (call !== void 0) { throw new TypeError("Derived constructors may only return object or undefined"); } return _assertThisInitialized(self); }
function _assertThisInitialized(self) { if (self === void 0) { throw new ReferenceError("this hasn't been initialised - super() hasn't been called"); } return self; }
function _isNativeReflectConstruct() { if (typeof Reflect === "undefined" || !Reflect.construct) return false; if (Reflect.construct.sham) return false; if (typeof Proxy === "function") return true; try { Boolean.prototype.valueOf.call(Reflect.construct(Boolean, [], function () {})); return true; } catch (e) { return false; } }
function _getPrototypeOf(o) { _getPrototypeOf = Object.setPrototypeOf ? Object.getPrototypeOf.bind() : function _getPrototypeOf(o) { return o.__proto__ || Object.getPrototypeOf(o); }; return _getPrototypeOf(o); }
var Drawing = require('./drawing.js');
var path = require('path');
var imgsz = require('../../../../image-size');
var mime = require('../../../../mime');
var uniqueId = require('../../../../lodash.uniqueid');
var EMU = require('../classes/emu.js');
var xmlbuilder = require('../../../../xmlbuilder');
var Picture = /*#__PURE__*/function (_Drawing) {
  _inherits(Picture, _Drawing);
  var _super = _createSuper(Picture);
  /**
   * Element representing an Excel Picture subclass of Drawing
   * @property {String} kind Kind of picture (currently only image is supported)
   * @property {String} type ooxml schema
   * @property {String} imagePath Filesystem path to image
   * @property {Buffer} image Buffer with image
   * @property {String} contentType Mime type of image
   * @property {String} description Description of image
   * @property {String} title Title of image
   * @property {String} id ID of image
   * @property {String} noGrp pickLocks property
   * @property {String} noSelect pickLocks property
   * @property {String} noRot pickLocks property
   * @property {String} noChangeAspect pickLocks property
   * @property {String} noMove pickLocks property
   * @property {String} noResize pickLocks property
   * @property {String} noEditPoints pickLocks property
   * @property {String} noAdjustHandles pickLocks property
   * @property {String} noChangeArrowheads pickLocks property
   * @property {String} noChangeShapeType pickLocks property
   * @returns {Picture} Excel Picture  pickLocks property
   */
  function Picture(opts) {
    var _this;
    _classCallCheck(this, Picture);
    _this = _super.call(this);
    _this.kind = 'image';
    _this.type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
    _this.imagePath = opts.path;
    _this.image = opts.image;
    _this._name = _this.image ? opts.name || uniqueId('image-') : opts.name || path.basename(_this.imagePath);
    var size = imgsz(_this.imagePath || _this.image);
    _this._pxWidth = size.width;
    _this._pxHeight = size.height;
    _this._extension = _this.image ? size.type : path.extname(_this.imagePath).substr(1);
    _this.contentType = mime.getType(_this._extension);
    _this._descr = null;
    _this._title = null;
    _this._id;
    _this._rId;
    // picLocks §20.1.2.2.31 picLocks (Picture Locks)
    _this.noGrp;
    _this.noSelect;
    _this.noRot;
    _this.noChangeAspect = true;
    _this.noMove;
    _this.noResize;
    _this.noEditPoints;
    _this.noAdjustHandles;
    _this.noChangeArrowheads;
    _this.noChangeShapeType;
    if (['oneCellAnchor', 'twoCellAnchor'].indexOf(opts.position.type) >= 0) {
      _this.anchor(opts.position.type, opts.position.from, opts.position.to);
    } else if (opts.position.type === 'absoluteAnchor') {
      _this.position(opts.position.x, opts.position.y);
    } else {
      throw new TypeError('Invalid option for anchor type. anchorType must be one of oneCellAnchor, twoCellAnchor, or absoluteAnchor');
    }
    return _this;
  }
  _createClass(Picture, [{
    key: "name",
    get: function get() {
      return this._name;
    },
    set: function set(newName) {
      this._name = newName;
    }
  }, {
    key: "id",
    get: function get() {
      return this._id;
    },
    set: function set(id) {
      this._id = id;
    }
  }, {
    key: "rId",
    get: function get() {
      return this._rId ? this._rId : 'rId' + this._id;
    },
    set: function set(rId) {
      this._rId = 'rId' + rId;
    }
  }, {
    key: "description",
    get: function get() {
      return this._descr !== null ? this._descr : this._name;
    },
    set: function set(desc) {
      this._descr = desc;
    }
  }, {
    key: "title",
    get: function get() {
      return this._title !== null ? this._title : this._name;
    },
    set: function set(title) {
      this._title = title;
    }
  }, {
    key: "extension",
    get: function get() {
      return this._extension;
    }
  }, {
    key: "width",
    get: function get() {
      var inWidth = this._pxWidth / 96;
      var emu = new EMU(inWidth + 'in');
      return emu.value;
    }
  }, {
    key: "height",
    get: function get() {
      var inHeight = this._pxHeight / 96;
      var emu = new EMU(inHeight + 'in');
      return emu.value;
    }

    /**
     * @alias Picture.addToXMLele
     * @desc When generating Workbook output, attaches pictures to the drawings xml file
     * @func Picture.addToXMLele
     * @param {xmlbuilder.Element} ele Element object of the xmlbuilder module
     */
  }, {
    key: "addToXMLele",
    value: function addToXMLele(ele) {
      var anchorEle = ele.ele('xdr:' + this.anchorType);
      if (this.editAs !== null) {
        anchorEle.att('editAs', this.editAs);
      }
      if (this.anchorType === 'absoluteAnchor') {
        anchorEle.ele('xdr:pos').att('x', this._position.x).att('y', this._position.y);
      }
      if (this.anchorType !== 'absoluteAnchor') {
        var af = this.anchorFrom;
        var afEle = anchorEle.ele('xdr:from');
        afEle.ele('xdr:col').text(af.col);
        afEle.ele('xdr:colOff').text(af.colOff);
        afEle.ele('xdr:row').text(af.row);
        afEle.ele('xdr:rowOff').text(af.rowOff);
      }
      if (this.anchorTo && this.anchorType === 'twoCellAnchor') {
        var at = this.anchorTo;
        var atEle = anchorEle.ele('xdr:to');
        atEle.ele('xdr:col').text(at.col);
        atEle.ele('xdr:colOff').text(at.colOff);
        atEle.ele('xdr:row').text(at.row);
        atEle.ele('xdr:rowOff').text(at.rowOff);
      }
      if (this.anchorType === 'oneCellAnchor' || this.anchorType === 'absoluteAnchor') {
        anchorEle.ele('xdr:ext').att('cx', this.width).att('cy', this.height);
      }
      var picEle = anchorEle.ele('xdr:pic');
      var nvPicPrEle = picEle.ele('xdr:nvPicPr');
      var cNvPrEle = nvPicPrEle.ele('xdr:cNvPr');
      cNvPrEle.att('descr', this.description);
      cNvPrEle.att('id', this.id + 1);
      cNvPrEle.att('name', this.name);
      cNvPrEle.att('title', this.title);
      var cNvPicPrEle = nvPicPrEle.ele('xdr:cNvPicPr');
      this.noGrp === true ? cNvPicPrEle.ele('a:picLocks').att('noGrp', 1) : null;
      this.noSelect === true ? cNvPicPrEle.ele('a:picLocks').att('noSelect', 1) : null;
      this.noRot === true ? cNvPicPrEle.ele('a:picLocks').att('noRot', 1) : null;
      this.noChangeAspect === true ? cNvPicPrEle.ele('a:picLocks').att('noChangeAspect', 1) : null;
      this.noMove === true ? cNvPicPrEle.ele('a:picLocks').att('noMove', 1) : null;
      this.noResize === true ? cNvPicPrEle.ele('a:picLocks').att('noResize', 1) : null;
      this.noEditPoints === true ? cNvPicPrEle.ele('a:picLocks').att('noEditPoints', 1) : null;
      this.noAdjustHandles === true ? cNvPicPrEle.ele('a:picLocks').att('noAdjustHandles', 1) : null;
      this.noChangeArrowheads === true ? cNvPicPrEle.ele('a:picLocks').att('noChangeArrowheads', 1) : null;
      this.noChangeShapeType === true ? cNvPicPrEle.ele('a:picLocks').att('noChangeShapeType', 1) : null;
      var blipFillEle = picEle.ele('xdr:blipFill');
      blipFillEle.ele('a:blip').att('r:embed', this.rId).att('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
      blipFillEle.ele('a:stretch').ele('a:fillRect');
      var spPrEle = picEle.ele('xdr:spPr');
      var xfrmEle = spPrEle.ele('a:xfrm');
      xfrmEle.ele('a:off').att('x', 0).att('y', 0);
      xfrmEle.ele('a:ext').att('cx', this.width).att('cy', this.height);
      var prstGeom = spPrEle.ele('a:prstGeom').att('prst', 'rect');
      prstGeom.ele('a:avLst');
      anchorEle.ele('xdr:clientData');
    }
  }]);
  return Picture;
}(Drawing);
module.exports = Picture;
//# sourceMappingURL=picture.js.map
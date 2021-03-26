/******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, { enumerable: true, get: getter });
/******/ 		}
/******/ 	};
/******/
/******/ 	// define __esModule on exports
/******/ 	__webpack_require__.r = function(exports) {
/******/ 		if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 			Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 		}
/******/ 		Object.defineProperty(exports, '__esModule', { value: true });
/******/ 	};
/******/
/******/ 	// create a fake namespace object
/******/ 	// mode & 1: value is a module id, require it
/******/ 	// mode & 2: merge all properties of value into the ns
/******/ 	// mode & 4: return value when already ns object
/******/ 	// mode & 8|1: behave like require
/******/ 	__webpack_require__.t = function(value, mode) {
/******/ 		if(mode & 1) value = __webpack_require__(value);
/******/ 		if(mode & 8) return value;
/******/ 		if((mode & 4) && typeof value === 'object' && value && value.__esModule) return value;
/******/ 		var ns = Object.create(null);
/******/ 		__webpack_require__.r(ns);
/******/ 		Object.defineProperty(ns, 'default', { enumerable: true, value: value });
/******/ 		if(mode & 2 && typeof value != 'string') for(var key in value) __webpack_require__.d(ns, key, function(key) { return value[key]; }.bind(null, key));
/******/ 		return ns;
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";
/******/
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = "./src/functions/functions.js");
/******/ })
/************************************************************************/
/******/ ({

/***/ "./src/functions/functions.js":
/*!************************************!*\
  !*** ./src/functions/functions.js ***!
  \************************************/
/*! no static exports found */
/***/ (function(module, exports) {

function asyncGeneratorStep(gen, resolve, reject, _next, _throw, key, arg) { try { var info = gen[key](arg); var value = info.value; } catch (error) { reject(error); return; } if (info.done) { resolve(value); } else { Promise.resolve(value).then(_next, _throw); } }

function _asyncToGenerator(fn) { return function () { var self = this, args = arguments; return new Promise(function (resolve, reject) { var gen = fn.apply(self, args); function _next(value) { asyncGeneratorStep(gen, resolve, reject, _next, _throw, "next", value); } function _throw(err) { asyncGeneratorStep(gen, resolve, reject, _next, _throw, "throw", err); } _next(undefined); }); }; }

/**
 * @customfunction
 * @param {string} address The address of the cell from which to retrieve the value.
 * @returns The value of the cell at the input address.
 **/
function getRangeValue(_x) {
  return _getRangeValue.apply(this, arguments);
}
/**
 * Returns the input text and change the background color of the cell.
 * @customfunction
 * @param {string} text Input text
 * @param {string} cellBackgroundColor Cell background color to be applied
 * @param {CustomFunctions.Invocation} invocation invocations 
 * @returns The input text.
 */


function _getRangeValue() {
  _getRangeValue = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee(address) {
    var context, selectedRange, range;
    return regeneratorRuntime.wrap(function _callee$(_context) {
      while (1) {
        switch (_context.prev = _context.next) {
          case 0:
            // Retrieve the context object. 
            context = new Excel.RequestContext();
            selectedRange = context.workbook.getSelectedRange();
            selectedRange.format.fill.color = "red"; // Use the context object to access the cell at the input address. 

            range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
            range.load();
            _context.next = 7;
            return context.sync();

          case 7:
            return _context.abrupt("return", range.values[0][0]);

          case 8:
          case "end":
            return _context.stop();
        }
      }
    }, _callee);
  }));
  return _getRangeValue.apply(this, arguments);
}

function stringCellFormatter(_x2, _x3, _x4) {
  return _stringCellFormatter.apply(this, arguments);
}
/**
 * Change color of cell
 * @customfunction
 * @param {string} text Input text
 * @param {string} color Cell color 
 * @returns The input text.
 */


function _stringCellFormatter() {
  _stringCellFormatter = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee2(text, cellBackgroundColor, invocation) {
    var context, selectedRange;
    return regeneratorRuntime.wrap(function _callee2$(_context2) {
      while (1) {
        switch (_context2.prev = _context2.next) {
          case 0:
            context = new Excel.RequestContext();
            selectedRange = context.workbook.getSelectedRange();
            selectedRange.load('address');
            selectedRange.format.fill.color = cellBackgroundColor;
            _context2.next = 6;
            return context.sync();

          case 6:
            return _context2.abrupt("return", invocation.address);

          case 7:
          case "end":
            return _context2.stop();
        }
      }
    }, _callee2);
  }));
  return _stringCellFormatter.apply(this, arguments);
}

function stringCellFormatterWithoutAsync(text, color) {
  var context = new Excel.RequestContext();
  var selectedRange = context.workbook.getSelectedRange();
  selectedRange.format.fill.color = color;
  var g = getGlobal();
  g.action = action();
  return text;
}
/**
 * Returns the input text and change the background color of the cell.
 * @customfunction
 * @param {string} text Input text
 * @param {string} cellBackgroundColor Cell background color to be applied 
 * @returns The input text.
 */


function stringCellFormatter2(_x5, _x6) {
  return _stringCellFormatter2.apply(this, arguments);
}
/**
 * Returns the input text and change the font type, font size, font color
 * @customfunction
 * @param {string} text Input text
 * @param {string} fontName font name
 * @param {string} fontSize font size
 * @param {string} fontColor font color
 * @returns Input text
 */


function _stringCellFormatter2() {
  _stringCellFormatter2 = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee3(text, cellBackgroundColor) {
    return regeneratorRuntime.wrap(function _callee3$(_context3) {
      while (1) {
        switch (_context3.prev = _context3.next) {
          case 0:
            Excel.run(function (context) {
              var range = context.workbook.getSelectedRange();
              range.format.fill.color = cellBackgroundColor;
              return context.sync();
            })["catch"](function (error) {
              console.log("Error: " + error);

              if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
              }
            });
            return _context3.abrupt("return", text);

          case 2:
          case "end":
            return _context3.stop();
        }
      }
    }, _callee3);
  }));
  return _stringCellFormatter2.apply(this, arguments);
}

function stringFontFormatter(_x7, _x8, _x9, _x10) {
  return _stringFontFormatter.apply(this, arguments);
}
/**
 * Return the address of the cell that invoked the custom function. 
 * @customfunction
 * @param {number} first First parameter.
 * @param {number} second Second parameter.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @requiresAddress 
 */


function _stringFontFormatter() {
  _stringFontFormatter = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee4(text, fontName, fontSize, fontColor) {
    var context, selectedRange;
    return regeneratorRuntime.wrap(function _callee4$(_context4) {
      while (1) {
        switch (_context4.prev = _context4.next) {
          case 0:
            context = new Excel.RequestContext();
            selectedRange = context.workbook.getSelectedRange();
            selectedRange.format.font.name = 'Times New Roman';
            selectedRange.format.font.color = 'red';
            selectedRange.format.font.size = 30;
            _context4.next = 7;
            return context.sync();

          case 7:
            return _context4.abrupt("return", text);

          case 8:
          case "end":
            return _context4.stop();
        }
      }
    }, _callee4);
  }));
  return _stringFontFormatter.apply(this, arguments);
}

function getAddressAndFormat(first, second, invocation) {
  var address = invocation.address;
  var addressWithoutSheet = address.split('!')[1];
  Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange(addressWithoutSheet);
    range.select();
    range.format.fill.color = "red";
    return context.sync();
  })["catch"](function (error) {
    console.log("Error: " + error);

    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
  return addressWithoutSheet;
}
/**
 * Return the address of the cell that invoked the custom function. 
 * @customfunction
 * @param {number} first First parameter.
 * @param {number} second Second parameter.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @requiresAddress 
 */


function getAddress(first, second, invocation) {
  var address = invocation.address;
  return address;
}
CustomFunctions.associate("GETRANGEVALUE", getRangeValue);
CustomFunctions.associate("STRINGCELLFORMATTER", stringCellFormatter);
CustomFunctions.associate("STRINGCELLFORMATTERWITHOUTASYNC", stringCellFormatterWithoutAsync);
CustomFunctions.associate("STRINGCELLFORMATTER2", stringCellFormatter2);
CustomFunctions.associate("STRINGFONTFORMATTER", stringFontFormatter);
CustomFunctions.associate("GETADDRESSANDFORMAT", getAddressAndFormat);
CustomFunctions.associate("GETADDRESS", getAddress);

/***/ })

/******/ });
//# sourceMappingURL=functions.js.map
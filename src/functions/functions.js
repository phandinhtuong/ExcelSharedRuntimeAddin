/**
 * @customfunction
 * @param {string} address The address of the cell from which to retrieve the value.
 * @returns The value of the cell at the input address.
 **/
 async function getRangeValue(address) {
  // Retrieve the context object. 
  var context = new Excel.RequestContext();
  var selectedRange = context.workbook.getSelectedRange();
  selectedRange.format.fill.color = "red";
  // Use the context object to access the cell at the input address. 
  var range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
  range.load();
  await context.sync();
  
  // Return the value of the cell at the input address.
  return range.values[0][0];
 }
 /**
  * Returns the input text and change the background color of the cell.
  * @customfunction
  * @param {string} text Input text
  * @param {string} cellBackgroundColor Cell background color to be applied
  * @param {CustomFunctions.Invocation} invocation invocations 
  * @returns The input text.
  */
async function stringCellFormatter(text, cellBackgroundColor, invocation){
  var context = new Excel.RequestContext();
  var selectedRange = context.workbook.getSelectedRange();
  selectedRange.load('address');
  selectedRange.format.fill.color = cellBackgroundColor;
  await context.sync();
  return invocation.address;
  //return selectedRange.address;
}
/**
 * Change color of cell
 * @customfunction
 * @param {string} text Input text
 * @param {string} color Cell color 
 * @returns The input text.
 */
function stringCellFormatterWithoutAsync(text, color){
  var context = new Excel.RequestContext();
  var selectedRange = context.workbook.getSelectedRange();
  selectedRange.format.fill.color = color;

  const g = getGlobal();
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
  async function stringCellFormatter2(text, cellBackgroundColor){
    Excel.run(function (context) {
      var range = context.workbook.getSelectedRange();
      range.format.fill.color = cellBackgroundColor;

      return context.sync();
    }).catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
    return text;
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
async function stringFontFormatter(text, fontName, fontSize, fontColor){
  var context = new Excel.RequestContext();
  var selectedRange = context.workbook.getSelectedRange();
  selectedRange.format.font.name = 'Times New Roman';
  selectedRange.format.font.color = 'red';
  selectedRange.format.font.size = 30;
  await context.sync();
  return text;
}
/**
 * Return the address of the cell that invoked the custom function. 
 * @customfunction
 * @param {number} first First parameter.
 * @param {number} second Second parameter.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @requiresAddress 
 */
 function getAddressAndFormat(first, second, invocation) {
  var address = invocation.address;
  var addressWithoutSheet = address.split('!')[1];
  Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange(addressWithoutSheet);

    range.select();
    range.format.fill.color = "red";

    return context.sync();
  }).catch(function (error) {
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
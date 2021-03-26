/**
 * Return the input text and change cell background color.
 * @customfunction
 * @param {string} text Input text.
 * @param {string} cellBackgroundColor Cell background color.
 * @param {CustomFunctions.Invocation} invocation Invocation object to get current cell. 
 * @requiresAddress 
 * @returns Input text.
 */
 function stringCellFormatter(text, cellBackgroundColor, invocation) {
  var address = invocation.address; //get address of the invocation / current cell, eg: Sheet1!A4
  var addressWithoutSheet = address.split('!')[1]; //split to get address without sheet
  Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange(addressWithoutSheet);
    range.select();
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
 * Return the input text and change the font type, font size, font color.
 * @customfunction
 * @param {string} text Input text.
 * @param {string} fontName Font name.
 * @param {number} fontSize Font size.
 * @param {string} fontColor Font color.
 * @param {CustomFunctions.Invocation} invocation Invocation object to get current cell.
 * @requiresAddress
 * @returns Input text.
 */
function stringFontFormatter(text, fontName, fontSize, fontColor,invocation){
  var address = invocation.address; //get address of the invocation / current cell, eg: Sheet1!A4
  var addressWithoutSheet = address.split('!')[1]; //split to get address without sheet
  Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange(addressWithoutSheet);
    range.select();
    range.format.font.name = fontName; //eg 'Times New Roman'
    range.format.font.color = fontColor;
    range.format.font.size = fontSize; 
    return context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
  return text;
}


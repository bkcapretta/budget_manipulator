// manipulations.js
// By: Bianca Capretta
// Link to Google Spreadsheet: 
// https://docs.google.com/spreadsheets/d/1qVG3xqOvEDnbB6sZWrl_2uIxxsNrE8mV3prYlwbaaRs/edit?usp=sharing

var listOfPercent = [];
var listOfMoney = [];

var ui = SpreadsheetApp.getUi(); 
var form = SpreadsheetApp.getActive().getSheetByName('Sheet1');
var budget = form.getRange('E2').getValue();
var doneCell = form.getRange('G2');
var dataRange = form.getRange('D5:E14');

// --------------------------------DONE BUTTON----------------------------------

// Purpose: Copy all the entered percents from D5 to D14 into a list
function copy() {
  for (i = 5; i < 15; i++) { // get values from cell and put into list
    listOfPercent.push(form.getRange('D' + i).getValue());
  }
  checkAndConvert();
}

// Purpose: to check if the total percents add to 100% and if so, to convert 
//     into effort in money
function checkAndConvert() {
  var totalPercent = form.getRange('D15').getValue();
  
  if (budget == '') {
    ui.alert('You need to enter an Annual Cost.');
    doneCell.setValue('No');
  }
  else if (totalPercent != 1) {
    doneCell.setValue('No');
    ui.alert('You need to enter amounts that add to 100%.');
  }
  else {
    convertToMoney();
    doneCell.setValue('Yes');
    ui.alert('You are ready to edit numbers in Percent Allocated and Money ' +
    		'Allocated to see adjustments.');
  }
}

// Purpose: to compute the money in effort for each project given annual budget 
//       and percent effort
function convertToMoney()
{
  var money = 0;
  for (i = 0; i < 10; i++) {
    money = budget*listOfPercent[i];  // calc the money per effort
    listOfMoney.push(money); 
    form.getRange('E' + (i+5)).setValue(money);
  }
}

// Purpose: to compute the effort in percent for each project given annual 
//       budget and money in effort
function convertToPercent()
{
  var percent = 0;
  for (i = 0; i < 10; i++) {
    percent = listOfMoney[i]/budget;  // calc the money per effort
    listOfPercent.push(percent); 
    form.getRange('D' + (i+5)).setValue(percent);
  }
}

// -----------------------------RESET BUTTON------------------------------------

// Purpose: to remove all data about project info 
function reset()
{
  listOfPercent = [];
  listOfMoney = [];
  
  // clear all the cells with inputted info
  form.getRange('C2:D2').setValue(''); // project names
  form.getRange('E2').setValue(0);
  form.getRange('B5:C14').setValue(''); // project/grant names
  dataRange.setValue(0); // percent effort and money effort
  
  doneCell.setValue('No');
}

// ----------------------------EDIT on TRIGGER----------------------------------

// Purpose: when a cell is edited, take the cell's new value and adjust 
//          other cells so that the total percent of all projects remains at 100%
function onEdit(e) {
  var ready = form.getRange('G2').getValue();
  
  // only edit cells [D5, E14] and if user has inputted info and pressed done
  if (withinRange(e) && ready == 'Yes')
  { 
    getEffort(); // gets current values in both columns and puts into the lists
    adjustValues(e.value, e.oldValue, e.range.getRow() - 5, e.range.getColumn() - 4);
  }
}

// Purpose: to check if the given cell is within range of cells D5 and E14. 
//     Return true if so and false if not
function withinRange(cell)
{
   var editRange = { // D5:E14
    top : 5,
    bottom : 14,
    left : 4,
    right : 5
  };

  // Return false if we're out of range
  var thisRow = cell.range.getRow();
  if (thisRow < editRange.top || thisRow > editRange.bottom) return false;

  var thisCol = cell.range.getColumn();
  if (thisCol < editRange.left || thisCol > editRange.right) return false;
  
  return true;
}

// Purpose: to get all the current values in the two columns of value (D and E)
function getEffort()
{
  // clear lists
  listOfPercent = [];
  listOfMoney = [];
  
  for (i = 5; i < 15; i++) { // gets values in percent effort column
    listOfPercent.push(form.getRange('D' + i).getValue());
  }
  for (i = 5; i < 15; i++) { // gets values in effort in money column
    listOfMoney.push(form.getRange('E' + i).getValue());
  } 
}

// Purpose: given an edited cell, consider the change (postive or negative) and 
//          respectively add or remove the changed content from the rest of the 
//          cells; update all other cells accordingly
// Arguments: editted value, previous value, the cell row that was edited 
//          (from 0-10) and the cell column the was edited (0-1)
function adjustValues(newVal, oldVal, cellRow, cellCol)
{
  // get the new difference
  var diff = newVal - oldVal;
  
  // edit cells to add to 100% again or make total money add to the budget amount
  if (form.getRange("D15").getValue() != 1.0 || form.getRange('E16').getValue() != budget) { 
    
    // get sum of all the percents besides the edited cell so value of change
    // can be appropriately distributed to the rest of the column
    var sum = getPartialSum(cellRow, cellCol); 
    // get sum of editted column (all the percent efforts OR money effort) 
    // besides the edited cell
   
    doMath(diff, cellRow, cellCol, sum);
    
    if (cellCol == 0) convertToMoney();
    if (cellCol == 1) convertToPercent();
  }
  
  if (form.getRange('D15').getValue() != 1.0) {
    //checkRoundErrors(); 
  }
}

// Purpose: take the difference of the editted value and distribute it among the
//     other cells; checks if either column (percent effort or money effort) is editted
// Alg: (percent effort / [sum of whole column - edited cell] ) 

// CONSIDERATION: How should the numbers update? For the changed amount to update 
//   evenly among them, or allocate more to the higher priority projects?
function doMath(diff, cellRow, cellCol, sum) 
{
  // algorithm favors the projects that already have more effort invested in 
  // them (higher effort * diff > lower effort * diff)
  for (i = 0; i < 10; i++) {
    if (i != cellRow) { 
      if (cellCol == 0) { // cell in percent effort was editted
      	// amount that can be added back or taken away from respective cell; update list values
        listOfPercent[i] -= (listOfPercent[i]/sum)*diff; 
        form.getRange("D" + (i + 5)).setValue(listOfPercent[i]);
      }
      
      if (cellCol == 1) { // cell in money effort was editted
        listOfMoney[i] -= (listOfMoney[i]/sum)*diff; 
        form.getRange("E" + (i + 5)).setValue(listOfMoney[i]);
      }
    }
  }
}

// Purpose: return the sum of the values in the edited column w/out editted cell's value
function getPartialSum(row, col)
{
   var sum = 0;
   for (i = 0; i < 10; i++) {
     if (i != row) {
       if (col == 0) sum += listOfPercent[i];
       if (col == 1) sum += listOfMoney[i];
     }
   }
   return sum;
}

// Purpose: if column of percents don't add to 100%, make it do so
function checkRoundErrors() 
{
  var total = form.getRange('D15').getValue();
  if (total != 1 || total > 0.999 || total < 1.001) {
    // round all the values first
    for (i = 0; i < 10; i++) {
      listOfPercent[i] = Math.round(listOfPercent[i]);
      form.getRange("D" + (i + 5)).setValue(listOfPercent[i]);
    }
    // if they don't still ad up to 100, take the small difference and add it 
    // to the smallest number if pos or take it from the largest num if neg
    if (form.getRange("D15").getValue() != 1.0) {
      var percDiff = 1 - form.getRange("D15").getValue(); // positive if under, negative if over
      var minIndex = getMin(listOfPercent);
      var maxIndex = getMax(listOfPercent);
      if (percDiff > 0) {
      	// add difference to smallest number
        form.getRange("D" + (minIndex+5)).setValue(listOfPercent[minIndex] + percDiff); 
      }
      else { // if difference is negative
      	// remove difference from largest number
        form.getRange("D" + (maxIndex+5)).setValue(listOfPercent[maxIndex] + percDiff); 
      }
    }
  } 
}

// Purpose: to return the index of the smallest item in the list
function getMin(list)
{
  var min = list[0];
  var minIndex = 0;
  for(i = 1; i < 10; i++) {
     if (list[i] < min) 
       min = list[i];
       minIndex = i;
  }
  return minIndex;
}

// Purpose: to return the index of the largest item in the list
function getMax(list)
{
  var max = list[0];
  var maxIndex = 0;
  for(i = 1; i < 10; i++) {
     if (list[i] > max) 
       max = list[i];
       maxIndex = i;
  }
  return maxIndex;
}
/**
 * @OnlyCurrentDoc
 */

/***************************************************************************
This function adds menu items to the top of the spreadsheet
****************************************************************************
*/

function onOpen() {
 //creates a menu item to run createMaturitySheet() function below
  SpreadsheetApp.getUi()
 .createMenu('TBM Assessment')
 .addItem('Create new TBM maturity worksheet', 'createMaturitySheet')
  .addItem('Show average TBM maturity results', 'createAveragesChart')
  .addToUi();
} //onOpen() function end

/****************************************************************************
The function below sets up initial input spreadsheet for a new user,
creates or expands the roll up worksheet for results
*****************************************************************************
*/

function createMaturitySheet() {
  var ui = SpreadsheetApp.getUi();
  
  //Create name for the new spreadsheet
  var result = ui.prompt(
  'Create new TBM maturity worksheet',
    'Please enter your name \(first name last initial\)',
    ui.ButtonSet.OK_CANCEL);
  
  //store input information
  var button = result.getSelectedButton();
  var sheetName = result.getResponseText();
  
  //test sheetName for blank value
  if (!sheetName) {
    Browser.msgBox('Error', 'You must enter a name for the new sheet',
                  Browser.Buttons.OK);
    return;
  }
  
  //test sheetName for duplicate value, keep previous one if client says OK
  var activeSpreadsheet = SpreadsheetApp.getActive();
  var theNewSheet = activeSpreadsheet.getSheetByName(sheetName);
  if (theNewSheet) {
    //process to handle duplicate names
    var resultDuplicate = Browser.msgBox('The name you entered already exists, OK to overwrite?', Browser.Buttons.OK_CANCEL);
    
    if (resultDuplicate != 'ok') { //kill function if user wants to use a different name
      Browser.msgBox('Please start over',Browser.Buttons.OK); return;
    } else if (resultDuplicate == 'ok') { //flag when duplicate is OK, so that we don't create another column in results spreadsheet
      var duplicateName = 1
      }
  
  } else {
    //Sets up a new tab on the spreadsheet
    
    //create the copy
    var copySheetName = 'MaturityTemplate'
    var copySheet = activeSpreadsheet.getSheetByName(copySheetName); 
        copySheet.copyTo(activeSpreadsheet)
        .setName(sheetName)
        .activate();
    var theNewSheet = activeSpreadsheet.getSheetByName(sheetName); // sets active spreadsheet to be the new sheet
    //adds the name at the top
    var theNewSheetTitle = Utilities.formatString('TBM Maturity Worksheet for %s', sheetName); //set header title
//   var theNewSheetHeader = [
//      theNewSheetTitle
//      ];
//    
    theNewSheet.getRange(1,1).setValue(theNewSheetTitle);
   }
  
  //expand data consolidation sheet, or create on first pass
  
  if (!duplicateName) { //skip this build out process if the spreadsheet is a copy
    var resultsSheetName = 'theResults';
    var theResultsSheet = activeSpreadsheet.getSheetByName(resultsSheetName);
    
    //add a new column if spreadsheet already exists, otherwise create the sheet
    if (theResultsSheet && !duplicateName) {//add the column
        theResultsSheet.insertColumnAfter(1); 
    } else {
      //create the sheet
      theResultsSheet = 
        activeSpreadsheet.insertSheet(resultsSheetName,activeSpreadsheet.getNumSheets() - 1); 
      //add total header
      var resultsTotalHeader = 'Totals'  
      theResultsSheet.getRange(1,3,1,1)
        .setValue(resultsTotalHeader);
     
      //add totals formulas and headers
      var inputMarker = activeSpreadsheet.getRangeByName('templateInputs');
      var r = 3;
      var resultsHeaders = ['Today','In 12 Months','Rank'];
      var rows = inputMarker.getNumRows();
      var rowTitles = [];
      var rowTitles = inputMarker.getValues()
      
      for (var i = 0; i < 3; i++) {//create a set of totals for each input field
       var inputHeaderRange = theResultsSheet.getRange(r-1,3,1,1);
       var inputFormulaRange = theResultsSheet.getRange(r,3,rows,1);
       inputHeaderRange.setValue(resultsHeaders[i]);
        inputFormulaRange.setFormulaR1C1('=AVERAGE(R[0]C[-1]:R[0]C[-2])');
        theResultsSheet.getRange(r,4,rows,1).setValues(inputMarker.getValues());
       var r = r + 2 + rows;
      }
    }  
    //add spreadsheet header
    theResultsSheet.getRange(1,2)
    .setValue(sheetName)
      
    //add formulas. Looks like each sheet's cell reference is perpetuated as code progresses
    var inputMarker = activeSpreadsheet.getRangeByName('templateInputs');
    var r = 3;
    var rows = inputMarker.getNumRows();
    var column = inputMarker.getColumn() - 1;
    var row = inputMarker.getRow() - rows + 1;
    var i = 0;
   
    
    for (i; i < 3; i++) {
      var theResultsSheetFormula = '=\'' +sheetName + '\'!R[' + row + ']C[' + column + ']';
      var inputFormulaRange = theResultsSheet.getRange(r,2,rows,1)
      .setFormulaR1C1(theResultsSheetFormula);
      var column = column + 1;
      row = row - rows - 2;
      var r = r + 2 + rows;
    }
  }
  theNewSheet.activate();
} // createMaturitySheet() function end

/*********************************************************************************
The function below consolidates the average results into an array, and
creates an html column chart displaying those average values
**********************************************************************************
*/

function createAveragesChart() {
  //collect data
  var activeSpreadsheet = SpreadsheetApp.getActive();
  var resultsSheetName = 'theResults';
  var theResultsSheet = activeSpreadsheet.getSheetByName(resultsSheetName);
  var inputMarker = activeSpreadsheet.getRangeByName('templateInputs'); //a range has been named on the template to make it easier 
  
 //get data for today and in 12 months
  var totalColumn = theResultsSheet.getLastColumn() - 1;
  var rows = inputMarker.getNumRows();
  var row = 2;
  
  var todayHeader = theResultsSheet.getRange(row, totalColumn).getValue();
  row++;
  var todayValues = theResultsSheet.getRange(row, totalColumn, rows, 1).getValues();
  row = row + rows + 1;
  var in12MonthsHeader = theResultsSheet.getRange(row, totalColumn).getValue();
  row++;
  var in12MonthsValues = theResultsSheet.getRange(row, totalColumn, rows, 1).getValues();
  var hAxisLabels = inputMarker.getValues();
  
  //build data table array
  
  var dataTable = new Array();
  
 
  dataTable[0] = ['',todayHeader,in12MonthsHeader]; //make header row
  
  for (i in hAxisLabels) { //load up table with data
    var rowValues = [ String(hAxisLabels[i]), Number(todayValues[i]), Number(in12MonthsValues[i]) ];
    dataTable.push(rowValues);
  }
 

 
  //load data into cache and run html to create column chart
  var cache = CacheService.getDocumentCache();
  var dataTableString = JSON.stringify(dataTable); //convert array to JSON to maintain format
    
  cache.put('mtData', dataTableString);
  Logger.log('original' + dataTableString);
 
  //this section points to html page and sets the page's size
  var html = HtmlService.createHtmlOutputFromFile('tbmMaturityPage')
  .setWidth(1000)
  .setHeight(450);
  SpreadsheetApp.getUi()
  .showModalDialog(html, 'Results');

//Save this section in case we need to build a chart in the spreadsheet 
//build chart in the spreadsheet
//  var theChartSheetName = 'testChart';
//  var theChartSheet = activeSpreadsheet.getSheetByName(theChartSheetName);
//  
//  if (!theChartSheet) { //checks to see if chart sheet has already been created 
//    theChartSheet = activeSpreadsheet.insertSheet(theChartSheetName, activeSpreadsheet.getNumSheets() + 1);
//  }
//  
//  var chartRange = theChartSheet.getRange(2,2,dataTable.length,3).setValues(dataTable);
//  
////  if (chart) { // removes chart if already created
////  theChartSheet.removeChart(chart);
////  }
//  
//  var chart = theChartSheet.newChart()
//  .setChartType(Charts.ChartType.SCATTER)
//   .addRange(chartRange)
//   .setPosition(2, 2, 0, 0)
//  .setOption('title','TBM Maturity Assessment Results')
//  .setOption('animation.duration', 1000);
//  
//  
//  theChartSheet.insertChart(chart.build());
// 
  
} // createAveragesChart() function end  
 
/***********************************************************************
this function is used to pass data to tbmMaturityPage.html via withSuccessHandler()
************************************************************************
*/
function grabTableData() {
  var cache1 = CacheService.getDocumentCache();
  var dataForChart = (cache1.get('mtData'));
  Logger.log('return:' + JSON.parse(dataForChart));
  return dataForChart;
} // grabTableData() function end












  
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
  .addItem('Show average TBM priority rankings','createRankChart')
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
  
  for (i in hAxisLabels) { //load up array with data
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
  .setWidth(820)
  .setHeight(450);
  SpreadsheetApp.getUi()
  .showModalDialog(html, 'TBM Maturity Assessment: Average Results');

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


/******************************************************************************
this function is used to create a chart to show ranking distributions
*******************************************************************************
*/

function createRankChart () {
  //collect data
  var activeSpreadsheet = SpreadsheetApp.getActive();
  var resultsSheetName = 'theResults';
  var theResultsSheet = activeSpreadsheet.getSheetByName(resultsSheetName);
  var inputMarker = activeSpreadsheet.getRangeByName('templateInputs'); /*a range has been named on 
  the template to make it easier to measure size*/
  
  //get location of ranking data on results sheet
  var totalColumn = theResultsSheet.getLastColumn() - 1;
  var rows = inputMarker.getNumRows();
  var row = (2 * 3) + (rows * 2);
  
//  //test
//  var testRange = theResultsSheet.getRange(row, totalColumn);
//  theResultsSheet.setActiveRange(testRange).activate();
  
  //get data for ranking loaded into initial array
  var rankHeader = theResultsSheet.getRange(row, totalColumn, 1).getValue();
  row++;
  var numColumns = totalColumn;
  var rankRange = theResultsSheet.getRange(1, 2, 1, numColumns).getValues(); //grab headers
  var rankValues = rankRange; // load headers into array
  
  
  for (i = 0; i < rows; i++) {//load up array with data
    var rankRange = theResultsSheet.getRange(row + i, 2, 1, numColumns).getValues();
    rankValues.push(rankRange[0]);
   
  }
 
  //restructure array into new array that fits chart format
  var endLength = rankValues[0].length - 1;
  var dataEnd = rankValues[0].length - 2;
  var datasetForRankArray = new Array();//temporary array for data assembly
  var rankArrayForChart = new Array();//final array for transfer
  
  
  // bar chart data format method

  for (i = 1; i < rows + 1; i++) {//creates new array
    
    var arrayRow = rankValues[i];
    var rowName = arrayRow.pop();
    var rowAvg = arrayRow.pop();
    var rowMin = arrayRow.sort().shift();
    var rowMax = arrayRow.pop();
    
    
    var rowValues = [
      rowName,
      rowAvg,
      rowMin,
      rowMax,
      ];
     
      rankArrayForChart.push(rowValues); 
    
  }
  
//  //scatter chart data format method. Depricated for now.
//  
//  for (i = 0; i < rows; i++) {//creates new array
//   //header row
//    datasetForRankArray.push([
//      String(rankValues[i + 1][endLength]),
//      'Priority Rank',
////      'LabelA',
//      'Average',
////      'LabelB'
//    ]);
//  
//    //data from individual respondents 
//    for (z = 0; z < dataEnd ; z++) {//load header, details, then average
//    
//        var rowValues = [
//          Number(rankValues[i + 1][z]),
//          0,
////          String(rankValues[0][z]),
//          null,
////          null,
//          ];
//        datasetForRankArray.push(rowValues);   
//     }
//     //last row with average
//          datasetForRankArray.push([
//          Number(rankValues[i + 1][dataEnd]),
//            null,
////            null,
//            1,
////            'Average',
//              ]);
//      //load the final array up 
//      rankArrayForChart.push(datasetForRankArray);
//      //clean out temp array for next run through loop
//      datasetForRankArray = new Array();
//    }
    
    //load data into cache and run html to create chart
    var cache = CacheService.getDocumentCache();
    var dataTableString = JSON.stringify(rankArrayForChart); //convert to JSON to maintain format thru transfer
    
    cache.put('rankData', dataTableString); //loads data into cache
//    Logger.log('original' + rankArrayForChart + '/n' + 'postJSON' + dataTableString);
    
    //this section points to html page and sets the page's size
    var html = HtmlService.createHtmlOutputFromFile('tbmRankPage')
    .setWidth(660)
    .setHeight(420);
    SpreadsheetApp.getUi()
    .showModalDialog(html, 'Average TBM Rankings (with spread)');
    
    
////test area    
//var cache1 = CacheService.getDocumentCache();
//    var dataForChart = (cache1.get('rankData'));
//    var testValue = JSON.parse(dataForChart);
//    var testValue2 = testValue[0]
////test area

    
  //test
//    Browser.msgBox('All Done :-)')
  
} //end of createRankChart()


/***************************************************************************
this function is used to pass data to tbmRankPage.html via withSuccessHandler()
****************************************************************************
*/
  
  function grabTableData2() {
    var cache1 = CacheService.getDocumentCache();
    var dataForChart = (cache1.get('rankData'));
//    var testValue = JSON.parse(dataForChart);
//    Logger.log('return:' + testValue);
    return dataForChart;
  }// grabTableData2() function end






  
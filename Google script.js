function onOpen() 
{
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('CH Reports')
  .addItem('Generate Weekly Report ', 'weeklyReport') 
  .addItem('Generate CH Summary', 'CHSummary')
  .addToUi();
  //Logger.log("Name of active spreadsheet: "+SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Parent Folder IDs").activate());
}


function CHSummary()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dateSheet = ss.getSheetByName("Tutorwise-Summary")
  
  var reportSheet = ss.getSheetByName("Summary Data");
  
  if(reportSheet == null)
  {
    reportSheet = ss.insertSheet();
    reportSheet.setName("Summary Data");
    reportSheet.getRange(1, 1).setValue("Tutor Id");
    reportSheet.getRange(1, 2).setValue("Subject");
    reportSheet.getRange(1, 3).setValue("Total Upload Count");
    reportSheet.getRange(1, 4).setValue("Total Skip Count (UQ+DQ)");
    reportSheet.getRange(1, 5).setValue("Answer rate (UQ+DQ)");
  }
  else if(reportSheet)
  {
    reportSheet.deleteRows(2, reportSheet.getLastRow());
  }
 
  var startDate = dateSheet.getRange(2, 7).getValue();
  var endDate = dateSheet.getRange(2,8).getValue();
  //Logger.log(currDate);
  var d = startDate.toString().split(" ");
  //Logger.log(d);
  var date = Utilities.formatDate(startDate,"IST", "d/M/yyyy");
  var mfix = {1:"st",2:"nd",3:"rd",4:"th",5:"th",6:"th",7:"th",8:"th",9:"th",10:"th",11:"th",12:"th",13:"th", 14: "th", 15:"th",16:"th",17:"th",18:"th",19:"th",20:"th",21:"st",22:"nd",23:"rd",24:"th",25:"th",26:"th",27:"th",28:"th",29:"th",30:"th",31:"st"};
  var months = {1 : "Jan",2 : "Feb",3 : "Mar",4 : "Apr",5 : "May",6 : "Jun",7 : "Jul",8 : "Aug",9 : "Sep",10 : "Oct",11 : "Nov",12 : "Dec"};

  var currMonth = date.split("/")[1];
  var currStartDate = Number(date.split("/")[0]);
  
  var currEndDate = Number(Utilities.formatDate(endDate, "IST", "d/M/yyyy").split("/")[0]);
  //var currYear = Number(date.split("/")[2]);
  //Logger.log(currStartDate + mfix[currStartDate] + " " + months[currMonth] + "--" + currEndDate);
  var rowCount = 2;

  for(i = currStartDate; i<=currEndDate; i++)
  {
    var currSheetName = i.toString()+ mfix[i]+ " " + months[currMonth];
    var currSheet = ss.getSheetByName(currSheetName);
    Logger.log(currSheetName+"---"+currSheet);
    
    if(currSheet)
    {
      //Logger.log(currSheetName);
      //var currSheetData = ss.getSheetByName(currSheetName).getRange().getValues();
      var currSheetValues = currSheet.getRange(1, 28, 100, 11).getValues();
      
      for (j = 2; j< 100; j++)
      {
        var TutorId = currSheetValues[j][0];
        if(TutorId.toString().length <1)
        {
          break;
        }
        var TutorCount = currSheetValues[j][1];
        var Subject = currSheetValues[j][2];
        
        var TotalUQSkipped = currSheetValues[j][6];
        var UQAnswerRate = currSheetValues[j][7];
        var TotalDQSkipped = currSheetValues[j][9];
        var DQAnswerRate = currSheetValues[j][10];
      
        var TotalSkipped = TotalUQSkipped + TotalDQSkipped;
        var AnswerRate = (UQAnswerRate + DQAnswerRate)/2;

        reportSheet.getRange(rowCount, 1).setValue(TutorId);
        reportSheet.getRange(rowCount, 2).setValue(Subject);
        reportSheet.getRange(rowCount, 3).setValue(TutorCount);
        reportSheet.getRange(rowCount, 4).setValue(TotalSkipped);
        reportSheet.getRange(rowCount, 5).setValue(AnswerRate);
        rowCount += 1;
      
      }
          
      //Logger.log(currSheetValues[0][1]);
    }
    
    
  }

  
  Browser.msgBox("Report Generation Done!");
}


function weeklyReport()
{
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var reportSheet = ss.getSheetByName("Weekly Report");
  var currDate = reportSheet.getRange(2, 1).getValue();
  //Logger.log(currDate);
  var d = currDate.toString().split(" ");
  //Logger.log(d);
  var date = Utilities.formatDate(currDate,"IST", "d/M/yyyy");
  
  var mfix = {1:"st",2:"nd",3:"rd",4:"th",5:"th",6:"th",7:"th",8:"th",9:"th",10:"th",11:"th",12:"th",13:"th",15:"th",16:"th",17:"th",18:"th",19:"th",20:"th",21:"st",22:"nd",23:"rd",24:"th",25:"th",26:"th",27:"th",28:"th",29:"th",30:"th",31:"st"};
  var months = {1 : "Jan",2 : "Feb",3 : "Mar",4 : "Apr",5 : "May",6 : "Jun",7 : "Jul",8 : "Aug",9 : "Sep",10 : "Oct",11 : "Nov",12 : "Dec"};

  //var currDate = date.split("/")[0];
  var currMonth = date.split("/")[1];
  var cd = Number(date.split("/")[0]);
 
  var sheetName1 = cd + mfix[cd] + " " + months[currMonth];
  cd += 1;
  var sheetName2 = cd + mfix[cd] + " " + months[currMonth];
  cd += 1;
  var sheetName3 = cd + mfix[cd] + " " + months[currMonth];
  cd += 1;
  var sheetName4 = cd + mfix[cd] + " " + months[currMonth];
  cd += 1;
  var sheetName5 = cd + mfix[cd] + " " + months[currMonth];
  cd += 1;
  var sheetName6 = cd + mfix[cd] + " " + months[currMonth];
  cd += 1;
  var sheetName7 = cd + mfix[cd] + " " + months[currMonth];
  //Tutors	Tutor count	Subject	Accepted	Rejected	Skipped	Missed	Answer rate
  var targetSheet = ss.getSheetByName("Weekly Report Input");
  targetSheet.clear();
  var fRow = ["Tutor I'd","Tutor count","Subject","Tutor Hours","Upload Ratio","Total UQ-Answered","Total UQ-Skipped","UQ Answer Rate","Total DQ-Answered","Total DQ-Skipped","DQ Answer Rate","Incorrect Skips","UQ-In session time","DQ-In session time"];
  targetSheet.appendRow(fRow);
  var sheet1 = ss.getSheetByName(sheetName1);
  if(sheet1)
  {
    var sheet1values = sheet1.getRange(2, 24, 100,14).getValues();
    for(i = 0; i < 100; i++)
    {
      if(sheet1values[i][0].length <1)
        break
        targetSheet.appendRow(sheet1values[i]);
      
    }
  }
  
  var sheet2 = ss.getSheetByName(sheetName2);
  if(sheet2)
  {
    var sheet2values = sheet2.getRange(2, 24, 100,14).getValues();
    for(i = 0; i < 100; i++)
    {
      if(sheet2values[i][0].length <1)
        break
        targetSheet.appendRow(sheet2values[i]);
    }
  }
  
  var sheet3 = ss.getSheetByName(sheetName3);
  if(sheet3)
  {
    var sheet3values = sheet3.getRange(2, 24, 100,14).getValues();
    for(i = 0; i < 100; i++)
    {
      if(sheet3values[i][0].length <1)
        break
        targetSheet.appendRow(sheet3values[i]);
    }
  }
  
  var sheet4 = ss.getSheetByName(sheetName4);
  if(sheet4)
  {
    var sheet4values = sheet4.getRange(2, 24, 100,14).getValues();
    for(i = 0; i < 100; i++)
    {
      if(sheet4values[i][0].length <1)
        break
        targetSheet.appendRow(sheet4values[i]);
    }
  }
  var sheet5 = ss.getSheetByName(sheetName5);
  if(sheet5)
  {
    var sheet5values = sheet5.getRange(2, 24, 100,14).getValues();
    for(i = 0; i < 100; i++)
    {
      if(sheet5values[i][0].length <1)
        break
        targetSheet.appendRow(sheet5values[i]);
    }
  }
  var sheet6 = ss.getSheetByName(sheetName6);
  if(sheet6)
  {
    var sheet6values = sheet6.getRange(2, 24, 100,14).getValues();
    for(i = 0; i < 100; i++)
    {
      if(sheet6values[i][0].length <1)
        break
        targetSheet.appendRow(sheet6values[i]);
    }
  }
  var sheet7 = ss.getSheetByName(sheetName7);
  if(sheet7)
  {
    var sheet7values = sheet7.getRange(2, 24, 100,14).getValues();
    for(i = 0; i < 100; i++)
    {
      if(sheet7values[i][0].length <1)
        break
        targetSheet.appendRow(sheet7values[i]);
    }
  }
}

function myFunction() {

  Logger.log("My function");

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var input = ss.getSheetByName("Input");
  var isError = false;

  try{
    var input_values = input.getRange(2, 1, input.getLastRow() - 1, input.getLastColumn()).getValues()
    var input_format = input.getRange(2, 1, input.getLastRow() - 1).getFontColors()
  }
  catch{
    Logger.log("No Data Persent");
    isError = true;
  }

}


function readEventInput() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var input = ss.getSheetByName("Input");
  var isError = false;

  try{
    var input_values = input.getRange(2, 1, input.getLastRow() - 1, input.getLastColumn()).getValues()
    var input_format = input.getRange(2, 1, input.getLastRow() - 1).getFontColors()
  }
  catch{
    Logger.log("No Data Persent");
    isError = true;
  }
  
  if(!isError)
  {
    let devotee_data = utilsDevoteeDataBySecurityID("ehy18a");
    var i;
    var output = [];
    var output_sheet = ss.getSheetByName("Pending Event");
    var prabhuji_mail = "mittalbrother@gmail.com";
    var mail_body_for_prabhuji = "Hare krishna Prabhuji\nDandwat Pranaam.\n\nA new event request has been raised. Please approve or reject using the below link\n\n" + ss.getUrl() + "\n\nYour servant"
    
    for (i = 0; i < input_values.length; i++) {
      if (input_format[i][0] === "#0000ff")
        continue;
      let event_s_date = Utilities.formatDate(input_values[i][3], "GMT+0530", "yyyy/MM/dd hh:mm aaa");
      let event_e_date = Utilities.formatDate(input_values[i][4], "GMT+0530", "yyyy/MM/dd hh:mm aaa");

      output.push(
        [
          "",
          "",
          Utilities.formatDate(new Date(), "GMT+0530", "yyyy/MM/dd"),
          devotee_data[0], // name
          input_values[i][2], // event_name
          event_s_date,
          event_e_date,
          input_values[i][5],
          input_values[i][6],
          devotee_data[2], // email
          devotee_data[1], // mobile no 
        ]
      )
      if(output.length > 0) // send mail to devotee
      {
        var mail_body_for_devotee = "Hare krishna Prabhuji\nDandwat Pranaam.\n\nA new event'"+ input_values[i][2] +"' request has been sent. Please wait for prabhuji's approval.\n\nYour servant";
          MailApp.sendEmail(devotee_data[2], "Event Request send for prabhuji's approval", mail_body_for_devotee);
      }
      input.getRange(i + 2, 1).setFontColor("#0000ff")
    }

    if (output.length > 0) {
      output_sheet.getRange(output_sheet.getLastRow() + 1, 1, output.length, output[0].length).setValues(output);
      MailApp.sendEmail(prabhuji_mail, "New Event Request Raised", mail_body_for_prabhuji);
    }
  }
}

function eventProcessing() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var approved_sheet = ss.getSheetByName("Approved Event")
  var rejected_sheet = ss.getSheetByName("Rejected Event")
  var pending_sheet = ss.getSheetByName("Pending Event")
  var today = Utilities.formatDate(new Date(), "GMT+0530", "yyyy/MM/dd")
  var i, j
  var input_values = []
  var pending_output = []
  var rejected_output = []
  var approved_output = []
  var approve_map = {}
  
  var isError = false;

  try {
    input_values = pending_sheet.getRange(2, 1, pending_sheet.getLastRow() - 1, pending_sheet.getLastColumn()).getDisplayValues()
  }
  catch (Exception) {
    Logger.log("No Pending Events")
    isError = false;
  }

  if(!isError)
  {
  for (i = 0; i < input_values.length; i++) {

    if (input_values[i][4] == "") //event_name
      break;

    if (input_values[i][0] == "") {
      pending_output.push(input_values[i])
      continue
    }

    if (input_values[i][0] == "Rejected") {
      var reject = [];

      for (j = 2; j < input_values[i].length; j++)
      reject.push(input_values[i][j]);
      reject.push("HG Rukmini Jeevan Prabhuji")
      reject.push(input_values[i][1])
      reject.push(today)
      rejected_output.push(reject)
      continue
    }

    if (input_values[i][0] == "Approved") {
      var approve = []
      for (j = 2; j < input_values[i].length; j++) {
        approve.push(input_values[i][j]);
      }
      approve.push("HG Rukmini Jeevan Prabhuji")
      approve.push(input_values[i][1])
      approve.push(today)
      approved_output.push(approve)
      continue
    }
  }

  if (input_values.length > 0) {
    //rejected section start
    if (rejected_output.length > 0) {
      rejected_sheet.getRange(rejected_sheet.getLastRow() + 1, 1, rejected_output.length, rejected_sheet.getLastColumn()).setValues(rejected_output);
       MailApp.sendEmail("mittalbrother@gmail.com", "Event Request Rejected", "nothing");
    }
    //rejected section end

    //approved section start
    if (approved_output.length > 0) {
      approved_sheet.getRange(approved_sheet.getLastRow() + 1, 1, approved_output.length, approved_sheet.getLastColumn()).setValues(approved_output);
      var isCreated = addEvents(approved_output);
      if(isCreated)
      {
        MailApp.sendEmail("mittalbrother@gmail.com", "Event Request Accepted", "thank you");
      }
    }
    //approved section end
  }

   pending_sheet.getRange(2, 1, pending_sheet.getLastRow() - 1, pending_sheet.getLastColumn()).clearContent();
  if (pending_output.length > 0)
    pending_sheet.getRange(2, 1, pending_output.length, pending_sheet.getLastColumn()).setValues(pending_output);
  }
}

function addEvents(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cal = CalendarApp.getCalendarById("hellohoneymittal@gmail.com");
  var i;
  var isCreated = true;
 try{
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var e_title = row[1] +" : " + row[2];
    var e_s_date = new Date(row[3]);
    var e_e_date = new Date(row[4]);
    var e_location = row[5];
    var e_description = row[6];
    cal.createEvent(e_title, e_s_date, e_e_date, { location: e_location, description: e_description });
  }
 }
 catch (Exception) {
    Logger.log("Event not created.")
    isCreated = false;
  }
  return isCreated;
 
}
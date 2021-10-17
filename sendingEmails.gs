function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Send email')
      .addItem('Send email', 'sendEmail')
      .addItem('Send Reminder email', 'sendReminder')
      .addToUi();
}

function sendEmail() {
  processRows('Email');
}

function sendReminder() {
  processRows('Reminder');
}

function processRows(action) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let dataRange = ss.getActiveSheet().getDataRange();
  let rows = dataRange.getValues();
  let headers = rows.shift();
  let data = new Map();
  rows
      .forEach(row => {      
        if(action == 'Email' || (action == 'Reminder' && row[4] == true)){
        if(data.has(row[0])){
          data.get(row[0]).push(row);
        }
        else{
          let list = new Array();
          list.push(row);
          data.set(row[0],list);
        }
        }
        });
  for(let [key, value] of  data.entries()){
    try {
        var body="";
      if(action == 'Email'){
       body = 
  "<html><head><style>table,table td {border: 1px solid #cccccc;}td {height: 80px;width: 160px;text-align: center;vertical-align: middle;}</style></head><body><h2>Order Details</h2><table style=width:100%><tr><td><h3>Order Id</h3></td><td><h3>Hub Name</h3></td><td><h3>Date</h3></td></tr>"}
  else{
    body = 
  "<html><head><style>table,table td {border: 1px solid #cccccc;}td {height: 80px;width: 160px;text-align: center;vertical-align: middle;}</style></head><body><h2>Reminder for Order Details</h2><table style=width:100%><tr><td><h3>Order Id</h3></td><td><h3>Hub Name</h3></td><td><h3>Date</h3></td></tr>"}
  for(let i=0; i<data.get(key).length;i++){
    body = body+"<tr><td>"+data.get(key)[i][1]+"</td><td>"+data.get(key)[i][2]+"</td><td>"+data.get(key)[i][3]+"</td></tr>"
  }
  body = body+ "</table><h3>Please put in your comment in the “Comments” column corresponding to these in the following sheet:  </h3></body></html>";
      var sub = 'Request budget';
      MailApp.sendEmail({
        to: key,
        subject: sub,
        htmlBody: body,
        });
  } catch (e) {
          status = `Error: ${e}`;
      }
  }       
}

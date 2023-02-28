/* Source: Material
   https://blog.gsmart.in/parse-and-extract-data-from-gmail-to-google-sheets/
   Helpful Tools:
   https://regex101.com/ */

//these functions filter and display emails based on the defined messages html
function getRelevantMessages()
{
  //searches through gmail to find transaction emails, with a limit of 10 results (start is 0 and max is 10) in order to make the search go faster
  //Gmail.app will return an array of Gmail threads matching this query
  var filter = 'newer_than:13d AND from:no.reply.alerts@chase.com OR from:gmarin1997@gmail.com AND subject:your "transaction with" AND label:inbox AND NOT label:payment_processing_done';
  var threads = GmailApp.search(filter);
  
  //creating an array and saving the emails found to the array
  var allMessages=[];
  Logger.log("The number of found threads are: " + threads.length);
  if (threads != null || threads.length > 0){
  for (var i = threads.length -1; i >= 0; i-- ) {
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++){
      allMessages.push(messages[j]);
    }
  }
}
  
  // Logger.log("From relevant messages there are: " + messages.length + " messages");
  return allMessages;
}


//these functions parse the data to grab the information we want
function parseMessageData(messages)
{
  var records=[];
  Logger.log("The number of messages in parse are: " + messages.length);
  for(var m=0;m<messages.length;m++){
    var text = messages[m].getPlainBody();
    var rx = /\bDate\s(\S.*?)\sat\s+([^.]+)\s+Merchant\s+(\S.*?)\sAmount\s\D*\b(\d+(?:\.\d+)?)/;
    // var rx = /\bA\s+charge\s+of\s+\D*\b(\d+(?:\.\d+)?)\s+at\s+(\S.*?)\s+has\s+been\s+authorized\s+on\s+(\S.*?)\s+at\s+([^.]+)\./;
    var matches = text.match(rx);
    if(matches) {
      Logger.log("The message is: " + messages[m]);
      Logger.log("MY LENGTH IS " + matches.length);
    }
    if(!matches || matches.length < 5){
      //No matches; couldn't parse continue with the next message
      Logger.log("No match, break...");
      continue;
    }
    
    var rec = {};
    rec.date= matches[1];
    rec.merchant = matches[3];
    rec.hour = matches[2];
    rec.amount = matches[4];
    
    records.push(rec);
    rec = {};
    matches = {};
  }
    
  return records;
}


//grabs messages and formats that them into a display for the web app and will be called by doGet
function getMessageDisplay()
{
  var templ = HtmlService.createTemplateFromFile('messages');
  templ.messages = getRelevantMessages();
  return templ.evaluate();  
}

//parses all relevant, filtered email and selects the sentence needed to determine transaction details and then formats it to the html template
function getParsedDataDisplay()
{

  var templ = HtmlService.createTemplateFromFile('parsed');
  templ.records = parseMessageData(getRelevantMessages());
  //Logger.log(temp1.records)

  return templ.evaluate();
}

//web app access through doGet or essentially a "GET" command
function doGet()
{
  //return getMessageDisplay();
 // Logger.log(getParsedDataDisplay().temp1.records)
  return getParsedDataDisplay();
}


//now to save the data to a google sheet :)
function saveDataToSheet(records)
{
  var spreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1vBvZ1X0sVvdo7Do-AqiP65gfqn1fqDMxGyZFlwmelTs/edit?usp=sharing");
  var sheet = spreadsheet.getSheetByName("Transactions");
  Logger.log("RECORDS: " + records);

  for(var r=0;r<records.length;r++)
  {
    Logger.log("RECORD "+r+" IS: "+records[r].date+ " " + records[r].hour+" "+records[r].merchant+" "+records[r].amount+"\n");
    sheet.appendRow([records[r].hour,records[r].merchant, records[r].amount, records[r].date]);
  }
  
}

function processTransactionEmails()
{
  var messages = getRelevantMessages();
  var records = parseMessageData(messages);
  saveDataToSheet(records);
  labelMessagesAsDone(messages);
  return true;
}

//preventing same email from being scanned again
function labelMessagesAsDone(messages)
{
  var label = 'payment_processing_done';
  var label_obj = GmailApp.getUserLabelByName(label);
  if(!label_obj)
  {
    label_obj = GmailApp.createLabel(label);
  }
  
  for(var m =0; m < messages.length; m++ )
  {
     label_obj.addToThread(messages[m].getThread() );  
  }
  
}

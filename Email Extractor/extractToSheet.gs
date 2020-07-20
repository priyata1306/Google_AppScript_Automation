var ui = SpreadsheetApp.getUi();
function onOpen(e){
  
  ui.createMenu("Gmail Manager").addItem("Get Emails by Label", "getGmailEmails").addToUi();
  
}

function getGmailEmails(){
  var input = ui.prompt('Label Name', 'Put Your Assigned Label Name', Browser.Buttons.OK_CANCEL);
  
  if (input.getSelectedButton() == ui.Button.CANCEL){
    return;
  }
  
  var label = GmailApp.getUserLabelByName(input.getResponseText());
  var threads = label.getThreads();
  
  for(var i = threads.length - 1; i >=0; i--){
    var messages = threads[i].getMessages();
    
    for (var j = 0; j <messages.length; j++){
      var message = messages[j];
      if (message.isUnread()){
        extractDetails(message);
        GmailApp.markMessageRead(message);
      }
    }
    threads[i].removeLabel(label);
    
  }
  
}

function extractDetails(message){
  var ContactType = 'Instant Message';
  var OpenedFor = message.getFrom();
  var ShortDescription = message.getSubject();
  var Description = message.getPlainBody();
  var Category = 'Application';
  var SubCategory = 'Functionality';
  var Impact = '4 - One User / Minor';
  var Urgency = '4 - Low';
  var ConfigurationItem = 'Cornerstone - PwC IT';
  var AssignmentGroup = 'PwC IT - APP SPT - L&D';
  
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  activeSheet.appendRow([ContactType,OpenedFor,ShortDescription,Description,Category,SubCategory,Impact,Urgency,ConfigurationItem,AssignmentGroup]);
}

function sendMail() {
  
  var email = 1;
  var first = 2;
  var access = 3;
  var key = 4;
  
  var ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
    
  var data = ws.getRange("A2:F" + ws.getLastRow()).getValues();
  
  data = data.filter(function(r){return r[5] == true });
    
 // Logger.log(data);
  
  data.forEach(function(row){
    
    //Logger.log(row[email]);
    
    
   if (row[access] == "Gomo"){
    
    var wsSettings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Gomo_Settings");
    var subject = wsSettings.getRange("B1").getValue();
    var sender = wsSettings.getRange("B2").getValue();
    var cc = wsSettings.getRange("B3").getValue();
     
    var emailTemp = HtmlService.createTemplateFromFile("GomoEmail");
     
    emailTemp.fn = row[first];
    var htmlMessage = emailTemp.evaluate().getContent();     
    
    GmailApp.sendEmail(
      row[email],
      subject, 
      "Your Email doesn't support HTML.", 
      {name : sender, htmlBody : htmlMessage, cc : cc}
    );
    
 }
  
    
  else if (row[access] == "Vyond"){
    
    var wsSettings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vyond _Settings");
    var subject = wsSettings.getRange("B1").getValue();
    var sender = wsSettings.getRange("B2").getValue();
    var cc = wsSettings.getRange("B3").getValue();
    
    var emailTemp = HtmlService.createTemplateFromFile("VyondEmail");
  
    emailTemp.fn = row[first];
    var htmlMessage = emailTemp.evaluate().getContent();
    GmailApp.sendEmail(
      row[email], 
      subject, 
      "Your Email doesn't support HTML.", 
      {name : sender, htmlBody : htmlMessage, cc : cc}
    );
    
  }
    
    
  else if (row[access] == "Storyline"){
    
    var wsSettings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Storyline_Settings");
    var subject = wsSettings.getRange("B1").getValue();
    var sender = wsSettings.getRange("B2").getValue();
    var cc = wsSettings.getRange("B3").getValue();
    var Error_mail_to = wsSettings.getRange("B4").getValue();
    
    var emailTemp = HtmlService.createTemplateFromFile("StorylineEmail");
  
    emailTemp.fn = row[first];
    if (row[key] != "" && row[key].length == 25){
    emailTemp.key = row[key];
    
    var htmlMessage = emailTemp.evaluate().getContent();
    
    GmailApp.sendEmail(
      row[email], 
      subject, 
      "Your Email doesn't support HTML.", 
      {name : sender, htmlBody : htmlMessage, cc : cc}
    );
   } 
    else{
      
      SpreadsheetApp.getUi().alert('Invalid: Please check the Key Length');
      GmailApp.sendEmail(
      Error_mail_to, 
      "ERROR !!!", 
      "Invalid: Please check the Key Length");
     
     // return "Invalid: Please check the Key Length";
    }
  }
    
    
  else if (row[access] == "Articulate 360"){
    
    var wsSettings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Articulate360_Settings");
    var subject = wsSettings.getRange("B1").getValue();
    var sender = wsSettings.getRange("B2").getValue();
    var cc = wsSettings.getRange("B3").getValue();
    
    var emailTemp = HtmlService.createTemplateFromFile("Articulate360Email");
  
    emailTemp.fn = row[first];
    var htmlMessage = emailTemp.evaluate().getContent();
    GmailApp.sendEmail(
      row[email], 
      subject, 
      "Your Email doesn't support HTML.", 
      {name : sender, htmlBody : htmlMessage, cc : cc}
    );
    
  }
    
  SpreadsheetApp.getUi().alert('Successful! The mail has been sent to intended user(s)');
 });
  
}

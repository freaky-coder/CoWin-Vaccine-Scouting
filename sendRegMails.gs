function sendRegMails()
{
  Utilities.sleep(75000);
  var sub= "ATTENTION: CoWin Vaccination Scouting Registration Successful"
  const htmlTemp = HtmlService.createTemplateFromFile("RegMailTemp");
  var ws = SpreadsheetApp.openById(ssID).getSheetByName("Response Database")
  var Avals = ws.getRange("D1:D").getValues();
  for (var i =2; i<= Avals.filter(String).length;i++)
  {
    if(ws.getRange(i,2,1,1).getValue()=="No")
    {
        htmlTemp.name = ws.getRange(i,4,1,1).getValue().split(" ")[0].toLowerCase().replace(/\b[a-z]/ig, function(match) {return match.toUpperCase()});;
        htmlTemp.email= ws.getRange(i,6,1,1).getValue();
        htmlTemp.district = ws.getRange(i,8,1,1).getValue();
        htmlTemp.state = ws.getRange(i,7,1,1).getValue();
        var message1 = htmlTemp.evaluate().getContent();
        //console.log(message1);
        GmailApp.sendEmail(ws.getRange(i,6,1,1).getValue(),sub,"Text",{name:"Ayush Kapoor",htmlBody: message1});
        ws.getRange(i,2,1,1).setValue('Yes')
        ws.getRange(i,5,1,1).setValue(ws.getRange(i,4,1,1).getValue().split(" ")[0].toLowerCase().replace(/\b[a-z]/ig, function(match) {return match.toUpperCase()}))
    }
    
  }
  

}

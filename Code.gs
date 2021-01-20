/**
 * Karthik Sreedhar
 */

function listUsers() {
  var FileIterator = DriveApp.getFilesByName('SPECIALXMLT');
  var id = "1nncwP3-7zFJt7rv9ixC5yU65AoBgOsaOy1gaY1smMuQ";
  var new_ss = SpreadsheetApp.create("SPECIALXMLT_emails");
  var new_ss_id = new_ss.getId();
  while (FileIterator.hasNext())
  {
    var file = FileIterator.next();
    var emails = [];
    if (file.getName() == 'SPECIALXMLT')
    {
      var v = Sheets.Spreadsheets.Values.get(id, 'A2:A').values;
      
      for(var index = 0; index < v.length; index++) {

        var name = v[index][0];
        var email = getEmail(name);
        emails.push([name, email]);
        
      }
    }   
    
  }
  var doc = DocumentApp.create('EMAILS______');
  var body = doc.getBody();
  body.insertParagraph(0, doc.getName())
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  table = body.appendTable(emails);
}


function getEmail(input) {
    var p = People.People.searchDirectoryPeople({query: input, readMask: 'emailAddresses', sources: 2});
    if (p.people == null) {
      return "Invalid";
    } 
    else {
      return p.people[0]['emailAddresses'][0]['value'];
    }
}
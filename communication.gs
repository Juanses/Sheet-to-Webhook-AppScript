var CommunicationClass = function(){
  
  this.sendPDFAttachment = function (to,subject,text,fileid){
    // Send an email with two attachments: a file from Google Drive (as a PDF) and an HTML file.
    var file = DriveApp.getFileById(fileid);
    MailApp.sendEmail(to, subject, text, {
      name: 'Juan Sebastián SUAREZ VALENCIA',
      attachments: [file.getAs(MimeType.PDF)]
    });
  }
  
  this.SendMultipleEmailfromTemplate = function(emaildata,templatename,keywords){
    //var templatename = "template.html"
    //var emaildata = {"to":"juan@carians.fr,juan@carians.fr","subject":"subject","name":"Nombre"};
    //var keywords = {"juan@carians.fr":{"nom":"SUAREZ","prenom":"Juan"},"juan@bress.fr":{"nom":"VALENCIA","prenom":"Sebastian"}};
    
    //I prepare the data to populate the email
    var template = HtmlService.createHtmlOutputFromFile(templatename).getContent();
    
    for (var key in keywords) {
      //Chaque email
      var htmlBody = template;      
      for (var cle in keywords[key]) {
        if(cle == "nom"){
          htmlBody = htmlBody.replace("%"+cle+"%", keywords[key][cle].toProperCase());
        }
        else{
          htmlBody = htmlBody.replace("%"+cle+"%", keywords[key][cle]);
        }
      }
      MailApp.sendEmail(key,emaildata["subject"],"This message requires HTML support to view.",{name: keywords["name"],htmlBody: htmlBody});
    }
  }
  
  this.SendEmailfromTemplate = function(emaildata,templatename,keywords){
    /*
    var templatename = "template.html"
    var emaildata = {"to":"juan@carians.fr","subject":"subject","name":"Nombre"};
    var keywords = {"nom":"SUAREZ","prenom":"Juan"};
    com.SendSingleEmail(emaildata,templatename,keywords) 
    */
    
    //I prepare the data to populate the email
    var htmlBody = HtmlService.createHtmlOutputFromFile(templatename).getContent();
    
    //I replace the values on the template with the values in the object
    for (var key in keywords) {
      if(key == "nom"){
        htmlBody = htmlBody.replace("%"+key+"%", keywords[key].toProperCase());
      }
      else{
        htmlBody = htmlBody.replace("%"+key+"%", keywords[key]);
      }
    }
    
    MailApp.sendEmail(emaildata["to"],emaildata["subject"],"This message requires HTML support to view.",{name: emaildata["name"],htmlBody: htmlBody});
  }
  
  this.sendtowebhook = function(url,payload){
    /*
    var data = {};
    data["destinataire"]="D8EN9DGGM";
    var tableau ="http://bit.ly/2yAbJ0s";
    data["message"] = "alerte"
    var comm = new CommunicationClass();
    comm.sendtowebhook("https://hook.integromat.com/g6976vdh9y5h4cos301nhwxfycy8u6lb",data);
    */
    
    var options =
        {
          "method"  : "POST",
          "payload" : payload,   
          "followRedirects" : true,
          "muteHttpExceptions": true
        };
    
    var result = UrlFetchApp.fetch(url, options);
  }
  
}

//We need this for SendEmail so let's keep it here
String.prototype.toProperCase = function () {
  return this.replace(/\w\S*/g, function(txt){return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();});
};
/* conversion */
var nomcolonneemail = "MAIL"
var nomcolonnetel = "TELEPHONE"

/* préparation */
var smssender = "Vaccination"; //max 11 char
var smstext = "Bonjour, votre inscription est confirmée pour la campagne de vaccination contre la grippe du 6 novembre au 36 Allée Ferdinand de Lesseps 37200 Tours. Bonne journée";
var emailsubject = "ds";
var emailtext = "Bonjour<br><br>Merci de votre inscription à MédecinDirect.<br><br>Pourriez-vous";

function clickRDVgetAppointmentsforIntervention (intervention_id) {
  var excel = new ExcelClass();
  var comm = new CommunicationClass();
  excel.setsheetbyId("17M8PJ-M5cjKMSkT7NySBIDj1--6rd9DjscDrwcWHiCY",0);
  excel.setColsAuto();
  
  var headers = {
    "Authorization" : "Basic " + Utilities.base64Encode(email + ':' + pass)
  };
  var params = {
    "method":"GET",
    "headers":headers
  };
  var url = "https://www.clicrdv.com/api/v1/appointments.json?apikey="+apikey+"&intervention_ids=["+intervention_id+"]";
  var response = UrlFetchApp.fetch(url, params); 
  response = JSON.parse(response);
  
  for (var i = 0; i < response.records.length; i++) {
    var element = response.records[i];
    excel.sheet.getRange(i+2,1).setValue(element.fiche.lastname);
    excel.sheet.getRange(i+2,2).setValue(element.fiche.firstname);
    excel.sheet.getRange(i+2,3).setValue(element.fiche.firstphone);
    excel.sheet.getRange(i+2,4).setValue(element.fiche.email);
    Logger.log(element.fiche.lastname);
  }
  
}


function SendTest() {
  var comm = new CommunicationClass();
  var data = {};
  data["Tel"] = "0033624010188";
  data["Email"] = "juanparis@gmail.com";
  data ["SMSsender"] = smssender;
  data ["SMStext"] = smstext;
  data ["Emailsubject"] = emailsubject;
  data ["Emailtext"] = emailtext;
  comm.sendtowebhook("https://hook.integromat.com/j3t9jqa94xki7dr59ypx2hao98ky16lk",data);
}


function SendToEveryone() {
  var excel = new ExcelClass();
  var comm = new CommunicationClass();
  excel.setsheetbyId("1GizDqxaaUi6qQFotdRnhgU0Dl9-k2PFsQEmEZJt6vQE",1);
  excel.setColsAuto();
  var values = excel.getallvalues(2,1,excel.cols.length);
  
  for (var i = 0; i < values.length; i++) {
    var data = excel.getvaluesfromrow(values,1,i+1);
    data["Email"] = data[nomcolonneemail];
    data["Tel"] = data[nomcolonnetel];
    
    if (data["Email"] != "" ||  data["Tel"] != ""){
      var tel = data["Tel"];
      if (tel.search(/^06/i) != -1){
        data["Tel"] = tel.replace(/(^0)/i, "0033");    
      }
      data ["SMSsender"] = smssender;
      data ["SMStext"] = smstext;
      data ["Emailsubject"] = emailsubject;
      data ["Emailtext"] = emailtext;
      comm.sendtowebhook("https://hook.integromat.com/j3t9jqa94xki7dr59ypx2hao98ky16lk",data);
    }
    
  }  
}

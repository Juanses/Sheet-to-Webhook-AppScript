/* conversion */
var nomcolonneemail = "MAIL"
var nomcolonnetel = "TELEPHONE"

/* préparation */
var smssender = "Vaccination"; //max 11 char
var smstext = "Bonjour, votre inscription est confirmée pour la campagne de vaccination contre la grippe du 6 novembre au 36 Allée Ferdinand de Lesseps 37200 Tours. Bonne journée";
var emailsubject = "ds";
var emailtext = "Bonjour<br><br>Merci de votre inscription à MédecinDirect.<br><br>Pourriez-vous";

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

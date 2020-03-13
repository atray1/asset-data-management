/*** Global Variables for Email ***/
//id of the template doc
var docTemplate = "1T1YT3DS_IUXM3SZ0jviJdjS_lSWFws4W75F2CquxgOw";

//name of the document
var docName = "Storage & Disposition Ticket";

/*** This function generates and sends the email
*/
function bounceEmail(fields, hazmat, sd_numm) {
  var name = fields[10];
  //var email_recipient = fields[12]; //uncomment to remove sd 
  var email_recipient = fields[12]+ ", " + "stordisp@roche.com"; //uncoment to email sd 
  
  var keys = ['KeyTicket', 'KeyAction', 'KeyManufacturer', 'KeyEINNumber', 'KeyModel', 'KeyAssetTagNumber', 
                'KeyDescription', 'KeyCondition', 'KeySerialNumber', 'KeyLabData', 'KeyName', 'KeyTimeStamp',
                'KeyEmail', 'KeyCostCenterNumber', 'KeySuperFunction', 'KeyRoomNumber', 'KeySpecial'];
  
  // Get document template, copy it as a new temp doc, and save the Doc's id
  var copyId = DriveApp.getFileById(docTemplate)
                       .makeCopy(docName+' for '+ name)
                       .getId();
  
  // Open the temporary document
  var copyDoc = DocumentApp.openById(copyId);
  
  // Get the document's body section
  var copyBody = copyDoc.getActiveSection();
  
  //replace placeholder keys with items in fields
  for(var j = 0; j < keys.length; j ++){
    copyBody.replaceText(keys[j], fields[j]);
  } 
  
  //save and close the temporary document
  copyDoc.saveAndClose();
  var pdf = DriveApp.getFileById(copyId).getAs("application/pdf");
  
  //delete the copy
  DriveApp.getFileById(copyId).setTrashed(true);
  
  //prepare the email
  var subject = docName+ ' ('+sd_numm+')'+' for '+ name;
  var ehs_link = 'http://gwiz.gene.com/groups/Ehs/l/lab_waste_management_guide/decontamination.shtml';
  var body = "Hello "+ name +",<br><br>Thank you for submitting your request, please print and attach the Storage & Disposition Ticket found in this email to your equipment."
  
  //determine whether EHS should be involved
  if(hazmat == 'Yes'){
    email_recipient = email_recipient + ', ehs-d@gene.com' //uncomment to activate EHS
    body = body + " Also, please attach your completed <a href='"+ehs_link+"'>Equipment Decontamination Certification Form</a> to your equipment.<br><br>*IMPORTANT - Failure to complete proper decontamination"
           +" procedures prior to form submission will heavily delay equipment pick up.<br><br>If you have any questions or concerns please email us at south_san_francisco.storage_and_disposition@roche.com."
           + "<br>EHS CRM - please email <a href='south_san_francisco.storage_and_disposition@roche.com'>south_san_francisco.storage_and_disposition@roche.com</a> when signed and ready for pick-up.";
  } else {
    body = body + " Also, please attach your completed <a href='"+ehs_link+"'>Non-Hazardous Material Certification Form</a> to your equipment.<br><br>*IMPORTANT - Failure to complete proper decontamination"
             +" procedures prior to form submission will heavily delay equipment pick up.<br><br>If you have any questions or concerns please email us at south_san_francisco.storage_and_disposition@roche.com.";
  }
  
  try {
    MailApp.sendEmail({
      to:email_recipient,
      subject: subject,
      htmlBody: body,
      noReply: true,
      attachments: [pdf]
    });
  } catch(error){
    MailApp.sendEmail({
      to:"atray1@gene.com",
      subject: "!!Bounce Email for S&D Form!!",
      body:error+" "+ subject
    });
  }
}


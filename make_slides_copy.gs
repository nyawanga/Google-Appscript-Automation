function duplicateSlide() {
  var file_id = "file_id"
  var dest_folder_id = "folder_id"
  var curr_date = new Date()
  curr_date.setMonth(curr_date.getMonth() - 1)
  var prev_month = Utilities.formatDate(curr_date, "GMT+03:00", "YYYYMM")
  var new_name = "Routine Monthly Reporting_" + prev_month                   // create a name with the month as a suffix
  
  var dest_folder = DriveApp.getFolderById(dest_folder_id)                   // get the destintion folder
  var dest_folder_files = dest_folder.getFiles()
  var monthly_slide = DriveApp.getFileById(file_id)                          // your routine monthly reporting slide
  var curr_name = monthly_slide.getName()
  var same_name_files = 0                                                    // this will help check for duplicate slides/slide names
  
  var success_recepients = "email_one@email.com,email_two@email.com"
  var failed_recepients = "email_one@email.com,email_two@email.com"
  
  while (dest_folder_files.hasNext()){
    var file = dest_folder_files.next()
    var file_name = file.getName()

    if (file_name == curr_name){
      var file_url = file.getUrl();
      same_name_files = same_name_files + 1
      send_email(failed_recepients,file_name +" Slide Failure","failure_email",file_url)   // send email incase of failure
    }
  }
  
  if (same_name_files == 0){
  // make the copy into the archive folder
    monthly_slide.makeCopy(curr_name, dest_folder)                                         // create the copy to the archive folder
  
  //rename the current file
    monthly_slide.setName(new_name)                                                        // rename the file for the next month
    var slide_url = monthly_slide.getUrl();
    send_email(success_recepients,new_name +" Slide","success_email",slide_url)            // send email incase of success
    
  //  monthly_slide.setTrashed(true)                                                       // incase you wanted to delete it
  }

}

function send_email(recepients, email_title, template_name, slide_url){
  // var attachment_file = DriveApp.getFileById("file_id").getAs(MimeType.PDF);            // attach the file as a PDF
  var htmlData = [email_title, slide_url];
  var html = HtmlService.createTemplateFromFile(template_name);
  html.data = htmlData;
  
  var htmlTemplate = html.evaluate().getContent();
  
  var emailBody = {
    to: recepients, //"email@email.com", //
    subject: email_title, 
    //attachments: [attachment_file],                                                       // incase of an attachement
    htmlBody: htmlTemplate 
  }

  MailApp.sendEmail(emailBody);           //SEND THE EMAIL
  //Logger.log("Success !");
  
}

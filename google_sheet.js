function sendEmails() {
  /*********************************************************************************************************************
  Author；Dino Zou
  Contact Informaiton: xinzhi.zou@globalpay.com 
  External sheet id for extracting email: 
  **********************************************************************************************************************/  
  /********************* Part1: Dialog to ask user the permission to run the script *********************/ 
  //Adding an alert to ask user if he is prepared to launch the solution 
   SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .createMenu('Custom Menu')
      .addItem('Show alert', 'showAlert')
      .addToUi();
  var ui = SpreadsheetApp.getUi(); // Same variations
  var result = ui.alert(
     'Please confirm',
     'to run the script for sending out the email notification',
      ui.ButtonSet.YES_NO);
  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    ui.alert('Confirmation received.');
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert('Permission denied.');
    // then the program should stop completely
    throw new Error("The script stops running");
  } 
  /********************* End *********************/
  
  /********************* Part2: Extracting data from the external email sheet(Top) *********************/ 
  //extracting data from an external source for email extraction
  var ss = SpreadsheetApp.openById("1eJ35ZicuMPLqDsntYwC4N46ljbqZK8i9C8VHc1fuivc");
  var ss_sheet = ss.getSheets()[0]; // "access data on different tabs"
  ss.setActiveSheet(ss_sheet);
  var ss_startRow = 2 ;
  var ss_row_range = ss_sheet.getRange(ss_startRow, 1, ss_sheet.getLastRow()-1,2);
  var ss_data = ss_row_range.getValues();
  var test = ss_data[34][1];
  var ss_email_address, ss_team_position;
  /********************* End *********************/ 
  
  /********************* Part3: Extracting data from the report sheet *********************/ 
  //collecting all the team names 
  var full_team_list = [];
  for(var x in ss_data){
     var row = ss_data[x];
     var team_name = row[0];
     full_team_list.push(team_name)   
  }
  // The office start for working on the email automation 
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2; // Start at second row because the first row contains the data labels
  var rowRange = sheet.getRange(startRow, 1, sheet.getLastRow()-1,11);
  var numRows = rowRange.getNumRows(); //Automatically find out how many rows we need to process e.g in this script, we started from row2 to row 16. Then, we have 15 rows to process.
  //  Extracting all the data in the spreedsheet
  var dataRange = sheet.getRange(startRow, 1, sheet.getLastRow()-1, 11 );
  // Fetch values for each row in the dataRange.
  var data = dataRange.getValues();  
  // Copy the first row which includes data label
  var label = sheet.getRange(1,1,1,9); 
  // Fetch values of columns labels
  label = label.getValues(); 
  /*********************  END  *********************/ 
  
  
  /********************* Part4: Autofill the email to in the sheet *********************/ 
  /*
  function getPosition(team_name)
  {
    for(var x in full_team_list) 
    {
       if(team_name == full_team_list[x])
       {
           ss_team_position = x;
       }
    }
  }
  var start_point = 2;
  var order = 1;
  for(var y in data)
  {
    var row = data[y];
    var team_name = row[5];
    getPosition(team_name);
    var ss_email_range = "K"+ start_point;
    var ss_email_range_message = sheet.getRange(ss_email_range); 
    var value = ss_data[ss_team_position][1];
    ss_email_range_message.setValue(value);
    start_point += parseInt(order); 
  }
  */
  /********************* End *********************/ 
  
  
  
  /********************* Part5: Get all the unique team name *********************/ 
  // extracting all the team name 
  var team_info =[];
  for( var e in data)
  {  
     var row = data[e];   
     var team_name = row[5];
     team_info.push(team_name);
  }
  // filtering out the repetitive name 
  var team_info_unique = []; 
  for(var x in team_info) 
  {
       if(team_info_unique.indexOf(team_info[x]) === -1)
       {
          team_info_unique.push(team_info[x]);      
       }
  }
    // using a team name as a list to find out the which row the team located in the spreadsheet and stored the infor into the array 
  var result = [];// The result is an array includes all the position of the tickets. 
  var data_check = data; 
  for(var m in team_info_unique)
  {
     for(var n in data_check)
     {
         var row = data_check[n]; 
         var team_name = row[5];
         var deleteit = [];
         if(team_info_unique[m] == team_name) 
         {        
            result.push(n); // get an array  
         }
     }
     if(result.length >=2)
     {
         sendMultipleEmail(result);// Debug this function   
     }
     else
     {  
         sendSingleEmail(data[result[0]]);       
     }
     result = []; 
  }
  /********************* End *********************/ 
  
  /********************* Part6: Inserting the data range in the email *********************/ 
  // I need an input in here extracting data from their 
  function range_date1(data_for_date) 
  {
    var ooo = data_for_date.length;
    var date_array =[];
    if(data_for_date.length > 1)
    {
          for(var b in data_for_date)
          {
            
            date_array.push(data[data_for_date[b]]);
            
           }
    }
    else
    {
          date_array = data_for_date; 
    }
    
    var date_pick=[];
    var month_pick = [];
    var template_date =['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    var template_pick = [0,1,2,3,4,5,6,7,8,9,10,11,12]
    var year;
    for( var e in date_array)
    { 
     
      var row = date_array[e];   
      var date_start = String(row[1]);
      if( date_start.length !== 19)
      {
        var see = date_start.length 
        date_pick.push(date_start.slice(8,10));
        month_pick.push(date_start.slice(4,7));
        year = date_start.slice(11,15)
      }
      else
      {
        date_pick.push(date_start.slice(4,6));
        month_pick.push(date_start.slice(0,3));
        year = date_start.slice(6,11)
      }
      
    }
    
    // find out the 1 or 2 unique month 
    var month_pick_find =[];
    var eye = month_pick[0];
    month_pick_find.push(eye);
    for (var s in month_pick)
    {
      var po = month_pick[s]
      if(eye != month_pick[s])
      {
        month_pick_find.push(month_pick[1])
      } 
    }
    
    //if we have 1 month, we don't compare, if we have 2 months, we compare
    if(month_pick_find.length == 2)
    {
      var month_array = [];
      var index_month = template_date.indexOf(month_pick_find[0]) + 1;
      month_array.push(template_pick[index_month]);
      index_month = template_date.indexOf(month_pick_find[1]) + 1;
      month_array.push(template_pick[index_month]);
      month_array = month_array.sort();
      // now, we need to compare the date
      var max = 0; var min = 0;
      for( var r in date_pick)
      {
        if(max < date_pick[r])
        {
          max = date_pick[r];
        }
        
       min = date_pick[1]
        
      }
       var print_date = "";
        print_date += month_array[0] + "/"+max + "/"+ year + " - " + month_array[1] + "/"+ min + "/" + year
        return print_date;
    }else {
        var index_month = template_date.indexOf(month_pick_find[0]) + 1
        var the_month = Math.floor(template_pick[index_month]) // might need some adjustment
        var max = 0; var min = 0;
        for( var r in date_pick)
        {
          if(max < date_pick[r])
          {
            max = date_pick[r]
          }        
        }
        min = date_pick[0];
        for( var r in date_pick)
        { 
          if(min < date_pick[r])
          { 
          }
          else
          {
            min = date_pick[r]
          }       
        }
          var print_date = "";
       if(min != max)
       {
          print_date += the_month + "/"+min+ "/"+ year + " - " + the_month + "/"+ max + "/" + year
       }
       else
       {
         max = parseInt(max)
         max+= 1; 
         print_date += the_month + "/"+min+ "/"+ year + " - " + the_month + "/"+ max + "/" + year
       }
        return print_date;
      }
    
  }
  
  /********************* End*********************/ 
  
 
  
  /********************* Part7: Sending email with multiple tickets included  *********************/ 
  function sendMultipleEmail(content) 
  {
     for(var h in content)
     {
       var row = data[content[h]];
       var emailAddress = row[9]; // extracting the data of column J
       var team_name = row[5];    // extracting the name of the team 
       var include_others = row[10]; // extracting the email address that will be sent to others
       var incident = row;         // copy the incident data from the variable row
       var tickets = ""; // Collecting all the names of the tickets for subjects
       incident.pop();             // remove the column of email address from the array
      //************ Define the HTML page *********//
      // The variable emailBody will define all the content we will include in the gmail.
      var emailBody = '<html><head><style>*{ font-family:tahoma;} table, th, td {border: 1px solid black;border-collapse: collapse; white-space: normal;} </style></head><body>'
      // hello, x team
      emailBody += '<p><strong> Hello  '+ team_name + ' team, </strong></p>' + '<p> Below incidents were identified on the 0900 Daily Incident Report for  ' + range_date1(content) + '</p>' 
      emailBody += '<table style= "width:100%">'
      //The start of the incident table. （ Very important part)
      // inserting multiple tickets in here 
      // here are the table headers 
      emailBody+= '<tr style="background-color:yellow">'
      for (var n in label[0])
      {
          emailBody+= '<th>'+ label[0][n]+ '</th>'
      }
      emailBody+= '</tr>'
      
      /// here are the table values 
       for(var k in content)
     {
       var row = data[content[k]];
       var inci = row;         // copy the incident data from the variable row
       inci.pop();   // remove the column of email address from the array
       inci.pop(); 
       emailBody += "<tr>";
       for (var p in inci) 
       { 
         if ( p == 0)
         {
          
           emailBody += "<td style='color: red'><strong>"+inci[p]+"</strong></td>"; 
           tickets += inci[p] +' & '
           
         }
         else if (p == 7)
         {
           emailBody += "<td style='color: red'><strong>"+inci[p]+"</strong></td>"; 
           
         }
         else
         {
          emailBody += "<td>"+inci[p]+"</td>"; 
         }
       }
       
       if(inci.length == 8)
       {
          emailBody += "<td></td>";
       }
       emailBody += "</tr>";  
     } 
      emailBody += '</table>';
      emailBody += '<br><p>Based on the below priority definitions, please assign the appropriate priority or once impact has been determined, downgrade the ticket to the correct priority based on the below criteria.</p>'
      emailBody += '<p>Please contact GPProblemManagement@globalpay.com if you have any questions. </p><br>'  
      //The end  of the incident table. 
      emailBody+= '<h4 style="text-decoration: underline"><strong>Priority Definitions:</strong></h4><p style="font-style: italic">Urgency defines how the incident affects the business.</p>'
      emailBody+= '<p><strong>P1</strong> - Payment transaction service impact<p>'
      emailBody+= '<p><strong>P2</strong> - Revenue generating or business/sales critical service impact </p>'
      emailBody+='<p><strong>P3</strong> - Other customer facing applications impact</p>'
      emailBody+='<p><strong>P4</strong> - Internal user impact</p>'
      emailBody+='<br><br><br><br>'
     emailBody+='<p><em>Kindly</em><strong style="color:blue;"> “Reply to All”</strong><em> when providing response.</em></p><p>Regards,</p><p><strong>Eleizer "Ezel" Magadia</strong></p><p><em>incident & Problem Management</em></p><br><h3><strong>Global Payments</strong></h3><a href="tel:+6325814727" >+632 581.4727</a> O<br><a href="https://mail.google.com/mail?view=cm&tf=0&to=eleizer.magadia@globalpay.com">eleizer.magadia@globalpay.com</a><br><p><em>Service. Driven. Commerce</em></p><p style="font-size:12px">NOTICE: This email message is for the sole use of the addressee(s) named above and may contain confidential and privileged information. Any unauthorized review, use, disclosure or distribution of this message or any attachments is expressly prohibited. If you are not the intended recipient, please contact the sender by reply email and destroy all copies and backups of the original message.<p>'
      emailBody+='</body></html>' 
      var message = ""; 
      // In here, I need to use a loop for developing a restriction that variable subject will change according to if the tickets send to the same team. 
      tickets = tickets.slice(0, tickets.length-3);
      var subject = "( Test Version) Priority Tickets "+ tickets; //***********(Change) The subject of the email to the format like priority ticket ** or **
      MailApp.sendEmail(emailAddress, subject, message,{
        htmlBody: emailBody,    // Options: Body (HTML)
        cc:include_others
      });
      break; 
     }
  }
  /********************* End *********************/ 
  
  /********************* Part8: Sending email with one ticket included *********************/ 
  function sendSingleEmail(content){
       
      var one_array = [];
      one_array.push(content); 
      var row = content;         // extracting the data from each row
      var emailAddress = row[9]; // extracting the data of column J
      var team_name = row[5];    // extracting the name of the team 
      var include_others = row[10]; // extracting the email address that will be sent to others
      var incident = row;         // copy the incident data from the variable row
      var tickets = ""; // Collecting all the names of the tickets for subjects
    
      incident.pop();             // remove the column of email address from the array
      incident.pop();
    
      //************ Define the HTML page *********//
      // The variable emailBody will define all the content we will include in the gmail.
    var emailBody = '<html><head><style> *{ font-family:tahoma; }table, th, td {border: 1px solid black;border-collapse: collapse; white-space: normal;} </style></head><body>' 
      // hello, x team
      emailBody += '<p><strong> Hello  '+ team_name + ' team, </strong></p>' + '<p>'+ incident[0]+' was identified on the 0900 Daily Incident Report for  ' + range_date1(one_array) + '</p>'  
      emailBody += '<table style= "width:100%">'
      //The start of the incident table. （ Very important part)
      emailBody += "<tr style='background-color:yellow'>";
      for (n in label[0])
      {
        emailBody+= '<th>'+ label[0][n]+ '</th>'
        
      }
     emailBody += "</tr>";
      emailBody += "<tr>";
      for (var p in incident) 
      { 
        if ( p == 0)
        {
          
          emailBody += "<td style='color: red'><strong>"+incident[p]+"</strong></td>"; 
          tickets += incident[p]
          
        }
        else if (p == 7)
        {
          emailBody += "<td style='color: red'><strong>"+incident[p]+"</strong></td>"; 
          
        }
        else
        {
          emailBody += "<td>"+incident[p]+"</td>"; 
        }
      }
      emailBody += "</tr>";
      emailBody += '</table>'
      //The end  of the incident table.
       emailBody += '<br><p>Based on the below priority definitions, please assign the appropriate priority or once impact has been determined, downgrade the ticket to the correct priority based on the below criteria.</p>'
      emailBody += '<p>Please contact GPProblemManagement@globalpay.com if you have any questions. </p><br>'
      emailBody+= '<h4 style="text-decoration: underline"><strong>Priority Definitions:</strong></h4><p style="font-style: italic">Urgency defines how the incident affects the business.</p>'
      emailBody+= '<p><strong>P1</strong> - Payment transaction service impact<p>'
      emailBody+= '<p><strong>P2</strong> - Revenue generating or business/sales critical service impact </p>'
      emailBody+='<p><strong>P3</strong> - Other customer facing applications impact</p>'
      emailBody+='<p><strong>P4</strong> - Internal user impact</p>'
      emailBody+='<br><br><br><br>'
      emailBody+='<p></p>'
      emailBody+='<p><em>Kindly</em><strong style="color:blue;"> “Reply to All”</strong><em> when providing response.</em></p><p>Regards,</p><p><strong>Eleizer "Ezel" Magadia</strong></p><p><em>incident & Problem Management</em></p><br><h3><strong>Global Payments</strong></h3><a href="tel:+6325814727" >+632 581.4727</a> O<br><a href="https://mail.google.com/mail?view=cm&tf=0&to=eleizer.magadia@globalpay.com">eleizer.magadia@globalpay.com</a><br><p><em>Service. Driven. Commerce</em></p><p style="font-size:12px">NOTICE: This email message is for the sole use of the addressee(s) named above and may contain confidential and privileged information. Any unauthorized review, use, disclosure or distribution of this message or any attachments is expressly prohibited. If you are not the intended recipient, please contact the sender by reply email and destroy all copies and backups of the original message.<p>'
      emailBody += '</body></html>'
      var message = ""; 
      // In here, I need to use a loop for developing a restriction that variable subject will change according to if the tickets send to the same team. 
      var subject = "(Test Version) Priority Ticket " + tickets; //***********(Change) The subject of the email to the format like priority ticket ** or **
      
      MailApp.sendEmail(emailAddress, subject, message,{
        htmlBody: emailBody,    // Options: Body (HTML)
        cc:include_others
      });
    }

  /********************* End *********************/ 
  // Run through each row, extract value and send out the email. 
   
  var final_result = ui.alert(
     'Congradulations!! Email Automation is Complete. ');
  
  /********************* End of the user dialog *********************/ 
}



 
  

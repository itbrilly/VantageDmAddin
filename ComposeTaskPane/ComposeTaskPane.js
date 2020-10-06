// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.

/// <reference path="../App.js" />

(function () {
  "use strict";
  var item;
  // The Office initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
      $(document).ready(function () {
          $('#insertText').click(insertText);
          $('#insertTOTO').click(AddEmailToTo);
          $('#insertTOBCC').click(AddEmailToBCC);
          $('#insertTOCC').click(AddEmailToCC);
          item=Office.context.mailbox.item;
      });
  };
  
  function insertText() {
    setRecipients();
    var textToInsert = $('#textToInsert').val();
    
    // Insert as plain text (CoercionType.Text)
    Office.context.mailbox.item.body.setSelectedDataAsync(
      textToInsert, 
      { coercionType: Office.CoercionType.Text }, 
      function (asyncResult) {
        // Display the result to the user
        if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
        }
        else {
        }
      });
      Office.context.mailbox.item.to.setSelectedDataAsync
  }
  function AddEmailToTo(){
    
    var newRecipients = [
        {
            "emailAddress": "sornaboobathy@aspiresys.com"
        },
        {
            "emailAddress": "satheesh.rajendra@aspiresys.com"
        }
    ];
    
    item.to.addAsync(newRecipients, FromCallback);
  }

  function FromCallback(result) {
    if (result.error) {
        $("#testspan").text(JSON.stringify(result.error));
        console.log(result.error);
    } else {
        
        console.log("Recipients added");
    }
}

  function AddEmailToCC(){
    var newRecipients = [
        {
            "emailAddress": "testmailcc@gmail.com"
        },
        {
            "emailAddress": "testmailcc1@gmail.com"
        }
    ];
    
    item.cc.addAsync(newRecipients, function(result) {
        if (result.error) {
            console.log(result.error);
        } else {
            console.log("Recipients added");
        }
    });
  }

  function AddEmailToBCC(){
    var newRecipients = [
        {
            "emailAddress": "itbrilly@gmail.com"
        },
        {
            "emailAddress": "itbrilly7@gmail.com"
        }
    ];
    
    item.bcc.addAsync(newRecipients, function(result) {
        if (result.error) {
            console.log(result.error);
        } else {
            console.log("Recipients added");
        }
    });
  }

  function setRecipients() {
    // Local objects to point to recipients of either
    // the appointment or message that is being composed.
    // bccRecipients applies to only messages, not appointments.
    var toRecipients, ccRecipients, bccRecipients;

    // Verify if the composed item is an appointment or message.
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        toRecipients = item.requiredAttendees;
        ccRecipients = item.optionalAttendees;
    }
    else {
        toRecipients = item.to;
        ccRecipients = item.cc;
        bccRecipients = item.bcc;
    }
    
    // Use asynchronous method setAsync to set each type of recipients
    // of the composed item. Each time, this example passes a set of
    // names and email addresses to set, and an anonymous 
    // callback function that doesn't take any parameters. 
    toRecipients.setAsync(
        [{
            "emailAddress":"graham@contoso.com"
         },
         {
            "displayName" : "Donnie Weinberg",
            "emailAddress" : "donnie@contoso.com"
         }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to set to-recipients of the item completed.

            }    
    }); // End to setAsync.


    // Set any cc-recipients.
    ccRecipients.setAsync(
        [{
             "displayName":"Perry Horning", 
             "emailAddress":"perry@contoso.com"
         },
         {
             "displayName" : "Guy Montenegro",
             "emailAddress" : "guy@contoso.com"
         }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to set cc-recipients of the item completed.
            }
    }); // End cc setAsync.


    // If the item has the bcc field, i.e., item is message,
    // set bcc-recipients.
    if (bccRecipients) {
        bccRecipients.setAsync(
            [{
                 "displayName":"Lewis Cate", 
                 "emailAddress":"lewis@contoso.com"
             },
             {
                 "displayName" : "Francisco Stitt",
                 "emailAddress" : "francisco@contoso.com"
             }],
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed){
                    write(asyncResult.error.message);
                }
                else {
                    // Async call to set bcc-recipients of the item completed.
                    // Do whatever appropriate for your scenario.
                }
        }); // End bcc setAsync.
    }
}
})();
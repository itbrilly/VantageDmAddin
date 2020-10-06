/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
    var mailboxItem;
    var xhr;
    var emailProcessRequest;
    var mailIdSuccess=false;
    var messageId="";
    var sendEvent;
    //var apiUrl="aspirevm10-20:9092/"; //live
    var apiUrl="http://localhost:63428/"; //local

    var mailBody;
    var serviceRequest=new Object();
    serviceRequest.Attachments = new Array();
    serviceRequest.CCList = new Array();
    serviceRequest.ToList = new Array();


    Office.initialize = function (reason) {
        mailboxItem = Office.context.mailbox.item;
    }


    // Entry point for Contoso Message Body Checker add-in before send is allowed.
    // <param name="event">MessageSend event is automatically passed by BlockOnSend code to the function specified in the manifest.</param>
    function mailIdCallback(result) {
        if(result.status!="succeeded"){
            CheckMailID();
        }
        else{
            makeServiceRequest(result.value);
            sendEvent.completed({ allowEvent: true });
            
        }
      }
    function CheckMailID(){
        Office.context.mailbox.item.getItemIdAsync(mailIdCallback);
    }  
    function validateBody(event) {
        sendEvent=event;
        //sendEvent.completed({ allowEvent: true });
        Office.context.mailbox.item.from.getAsync(FromCallback);
        //makeServiceRequest();
        
    }
    function FromCallback(asyncResult) {
        serviceRequest.From=asyncResult.value.emailAddress;
        Office.context.mailbox.item.to.getAsync(ToCallback);
    }
       
    function checkBodyOnlyOnSendCallBack(asyncResult) {
        CheckMailID();
        while(mailIdSuccess){
            makeServiceRequest(result.value)
            asyncResult.asyncContext.completed({ allowEvent: true });
           
        }
        // Allow send.
    }
    function CCCallback(asyncResult) {
        CCBuild(asyncResult.value)
         Office.context.mailbox.item.subject.getAsync(SubjectCallback);
    }
    function CCBuild(CCArray){
        if (CCArray.length > 0) {
            for (i = 0 ; i < CCArray.length ; i++) {
                var cCObj=new Object();
                cCObj.Name=CCArray[i].displayName;
                cCObj.Address=CCArray[i].emailAddress;
                serviceRequest.CCList[i]=JSON.parse(JSON.stringify(cCObj));
            }
          }
    }
    function ToCallback(asyncResult) {
        ToBuild(asyncResult.value);
         Office.context.mailbox.item.cc.getAsync(CCCallback);
    }
    function ToBuild(ToArray){
        if (ToArray.length > 0) {
            for (i = 0 ; i < ToArray.length ; i++) {
                var tOObj=new Object();
                tOObj.Name=ToArray[i].displayName;
                tOObj.Address=ToArray[i].emailAddress;
                serviceRequest.ToList[i]=JSON.parse(JSON.stringify(tOObj));
            }
          }
    }
    function SubjectCallback(asyncResult) {
        serviceRequest.Subject = asyncResult.value;
         Office.context.mailbox.item.body.getAsync(
            "text",
            { asyncContext: "This is passed to the callback" },
            BodyCallback);
    }
    function BodyCallback(asyncResult) {
        serviceRequest.Body = asyncResult.value;
        var item = Office.context.mailbox.item;
        var listOfAttachments = [];
        var options = {asyncContext: {currentItem: item}};
        item.getAttachmentsAsync(options, AttachmentCallback);
   }
   function AttachmentCallback(result) {
       try{
        var proceedToServiceCall=false;
        if (result.value.length > 0) {
            var resultLength=result.value.length-1;
          for (i = 0 ; i < result.value.length ; i++) {
            //result.asyncContext.currentItem.getAttachmentContentAsync(result.value[i].id, handleAttachmentsCallback);
            result.asyncContext.currentItem.getAttachmentContentAsync(result.value[i].id, function(r){  
                setTimeout(function () {     
                    var attachmentObj=new Object();
                attachmentObj.AttachmentContent=r.status=="succeeded"?r.value.content:"";
                var calculatedSize=FindAttachmentSize(r.value.content);
                attachmentObj.Name=result.value.filter(function(attachmnt){return attachmnt.size===calculatedSize})[0].name; 
                serviceRequest.Attachments.push(attachmentObj);
                  }, 1000);       
            });
          }
          setTimeout(function () { makeServiceRequest();  }, 2000);        
        }
        else{
            makeServiceRequest();  
        }
       }
       catch(err){
		testMethod(err.message);  
	  }
  }
  function FindAttachmentSize(attachmentContent){
    var n = attachmentContent.length;
    var lastTwoChar=attachmentContent.substr(attachmentContent.length -2);
    var lastOneChar=attachmentContent.substr(attachmentContent.length -1);
    var lC=0;
    if(lastTwoChar=="=="){
        lC=2
    }
    else if(lastOneChar=="="){
        lC=1
    }
    return (n * (3/4)) - lC;


  }
    // Invoke by Contoso Subject and CC Checker add-in before send is allowed.
    // <param name="event">MessageSend event is automatically passed by BlockOnSend code to the function specified in the manifest.</param>
    function validateSubjectAndCC(event) {
        shouldChangeSubjectOnSend(event);
    }

    // Check if the subject should be changed. If it is already changed allow send. Otherwise change it.
    // <param name="event">MessageSend event passed from the calling function.</param>
    function shouldChangeSubjectOnSend(event) {
        mailboxItem.subject.getAsync(
            { asyncContext: event },
            function (asyncResult) {
                addCCOnSend(asyncResult.asyncContext);
                //console.log(asyncResult.value);
                // Match string.
                var checkSubject = (new RegExp(/\[Checked\]/)).test(asyncResult.value)
                // Add [Checked]: to subject line.
                subject = '[Checked]: ' + asyncResult.value;

                // Check if a string is blank, null or undefined.
                // If yes, block send and display information bar to notify sender to add a subject.
                if (asyncResult.value === null || (/^\s*$/).test(asyncResult.value)) {
                    mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Please enter a subject for this email.' });
                    asyncResult.asyncContext.completed({ allowEvent: false });
                }
                else {
                    // If can't find a [Checked]: string match in subject, call subjectOnSendChange function.
                    if (!checkSubject) {
                        subjectOnSendChange(subject, asyncResult.asyncContext);
                        //console.log(checkSubject);
                    }
                    else {
                        // Allow send.
                        asyncResult.asyncContext.completed({ allowEvent: true });
                    }
                }

            }
          )
    }

    // Add a CC to the email.  In this example, CC contoso@contoso.onmicrosoft.com
    // <param name="event">MessageSend event passed from calling function</param>
    function addCCOnSend(event) {
        mailboxItem.cc.setAsync(['Contoso@contoso.onmicrosoft.com'], { asyncContext: event });        
    }

    function subjectOnSendChange(subject, event) {
        mailboxItem.subject.setAsync(
            subject,
            { asyncContext: event },
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Unable to set the subject.' });

                    // Block send.
                    asyncResult.asyncContext.completed({ allowEvent: false });
                }
                else {
                    // Allow send.
                    asyncResult.asyncContext.completed({ allowEvent: true });
                }

            });
    }
    function makeServiceRequest11111(mailId) {
        var attachment;
        xhr = new XMLHttpRequest();
        // Update the URL to point to your service location.
        //xhr.withCredentials = true;
        xhr.open("POST", "http://localhost:63428/api/Common/SaveEmailWithAttachmentsFromComposeMode", true);
        //xhr.setRequestHeader("Access-Control-Allow-Origin", "http://localhost:63428/api/SendEmailProcess");
        xhr.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        xhr.onreadystatechange = requestReadyStateChange;
    
        // Send the request. The response is handled in the 
        // requestReadyStateChange function.
        xhr.send(JSON.stringify(serviceRequest));
    }

    function userinfo(){
        var userinfo= JSON.parse(localStorage.getItem("PluginUserInfo"));
        if(userinfo==null){
           userinfo=new Object();
           userinfo.UserName="";
           userinfo.AccessToken="";
        }
        return userinfo;
      }

    function makeServiceRequest(mailId){
        var uinfo=userinfo();
        $.ajax({
          url: apiUrl+"api/Common/SaveEmailWithAttachmentsFromComposeMode",
          method: 'POST',
          data: serviceRequest,
          headers:  
                      {  
                          Authorization: 'Basic ' + btoa(uinfo.UserName + ':' + uinfo.AccessToken)  
                      },
          success: function (data) {
            sendEvent.completed({ allowEvent: true });
          },
         fail : function( jqXHR, textStatus,err ) {
            sendEvent.completed({ allowEvent: true });
         }
      })
      }

    function requestReadyStateChange() {
        sendEvent.completed({ allowEvent: true });
    }

    function testMethod(s){
        var reqobj=new Object();
        reqobj.s1=s;
        $.ajax({
          url: apiUrl+"api/AttachmentService/ProcessEmail",
          method: 'POST',
          data: reqobj,
          success: function (data) {
          },
         fail : function( jqXHR, textStatus,err ) {
          alert( "Request failed: " + textStatus );
         }
      })
      }

  (function () {
  "use strict";

  var messageBanner;
  var xhr;
  //var apiUrl="aspirevm10-20:9092/"; //live
  var apiUrl="http://localhost:63428/"; //local
  
var  serviceRequest = new Object();
    Office.initialize = function (reason) {
    $(document).ready(function () {
       UpdateTaskPaneUI(Office.context.mailbox.item);
       MulitiSelectableDropDown();
        $("#idmainmenu").click(function(){
          InitialLoadSettings();
        });
        $("#idmainmenu1").click(function(){
          InitialLoadSettings();
        });
        $("#spnDashboard1").click(function(){
          InitialLoadSettings();
        });
        $("#btnLogin").click(function(){
          APIlogin();
          });
        $("#btnCreateTask").click(function(){
        CreateNewTask(Office.context.mailbox.item);
        });

        $("#btnExportEmailModal").click(function(){
          $('#exportEmailModal').modal('hide');
        });
  
        $("#btnAddToCRM").click(function(){
          SaveContactInformation(Office.context.mailbox.item);
        });
        
        $("#btnExportEmail").click(function(){
          FindAttachmentToken(Office.context.mailbox.item);
        });
        
  
        $("#btnLogout").click(function(){
          APIlogout();
        });
        
      setTimeout(function(){
         
        testMethod();  
      }, 2000);
      
      InitialLoadSettings();
    initApp();
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);
    });
  };

  function MulitiSelectableDropDown(){
    $('#drpdwnType').select2({
      "placeholder": "Select Type",
       width: '100%',
    });
    $('#drpdwnChecklist').select2({
      "placeholder": "Select Checklist",
      width: '100%',
    });
    $('#drpdwnFund').select2({
      "placeholder": "Select Fund",
      width: '100%',
    });
  }

  function itemChanged(eventArgs) {
    // Update UI based on the new current item
    UpdateTaskPaneUI(Office.context.mailbox.item);
  }

  function UpdateTaskPaneUI(item)
  {
  // Assuming that item is always a read item (instead of a compose item).
  if (item != null) {
    testMethod();
    GetContactInformation(Office.context.mailbox.item);
  }
  }

  function testMethod(){
    var uinfo=userinfo();
    var reqobj=new Object();
    reqobj.s1="";
    $.ajax({
      url: apiUrl+"api/AttachmentService/ProcessEmail",
      method: 'POST',
      data: reqobj,
      headers:  
                  {  
                      Authorization: 'Basic ' + btoa(uinfo.UserName + ':' + uinfo.AccessToken)  
                  },
      success: function (data) {
      },
     fail : function( jqXHR, textStatus,err ) {
      alert( "Request failed: " + textStatus );
     }
  })
  }

  function GetCreateTaskData(){
    var uinfo=userinfo();
    var reqobj=new Object();
    $.ajax({
      url: apiUrl+"api/Inbox/CreateTaskData",
      method: 'POST',
      data: reqobj,
      headers:  
                  {  
                      Authorization: 'Basic ' + btoa(uinfo.UserName + ':' + uinfo.AccessToken)  
                  },
      success: function (response) {
      if(response.ActivityTypeList!=null){
        var activity="";
        response.ActivityTypeList.forEach(element => {
          activity+='<option value="' + element.KeywordValueID + '">' + element.Value + '</option>';
        });
        $('#drpdwnType').empty().append(activity);
        $('#drpdwnType').val('');
      }
      if(response.CheckLists!=null){
        var checkList="";
        response.CheckLists.forEach(element => {
          checkList+='<option value="' + element.KeywordID + '">' + element.Value + '</option>';
          $('#drpdwnChecklist').append('<option value="' + element.KeywordID + '">' + element.Value + '</option>');
        });
        $('#drpdwnChecklist').empty().append(checkList);
        $('#drpdwnChecklist').val('');
      }
      
      },
     fail : function( jqXHR, textStatus ) {
     }
  })
  }
  
      function initApp() {
        if (Office.context.mailbox.item.attachments == undefined) {

        } else if (Office.context.mailbox.item.attachments.length == 0) {

        } else {

            // Initalize a context object for the app.
            //   Set the fields that are used on the request
            //   object to default values.
            serviceRequest = new Object();
            serviceRequest.attachmentToken = "";
            serviceRequest.ewsUrl = Office.context.mailbox.ewsUrl;
            serviceRequest.attachments = new Array();
        }
    };
	
	function FindAttachmentToken() {
    Office.context.mailbox.getCallbackTokenAsync(attachmentTokenCallback);
};
function userinfo(){
  var userinfo= JSON.parse(localStorage.getItem("PluginUserInfo"));
  if(userinfo==null){
     userinfo=new Object();
     userinfo.UserName="";
     userinfo.AccessToken="";
  }
  return userinfo;
}

function InitialLoadSettings(){
  $("#wrapper").slideDown();
  $("#containerDashboard").slideDown();
  $("#loginContainer").slideUp();
  $("#containerSettings").slideUp();
  GetContactInformation(Office.context.mailbox.item);
}

function APIlogin(){
      
  var username=document.getElementById("InputUserName").value;
  var password=document.getElementById("InputPassword").value;
  if(username==null || username==""){
    document.getElementById("LoginError").innerHTML = "Please Enter Username";
    return;
  }
  else if (password==null || password==""){
    document.getElementById("LoginError").innerHTML = "Please Enter Password";
    return;
  }
  else{
    document.getElementById("LoginError").innerHTML = "";
  }
  var loginRequest=new Object();
  loginRequest.UserName=btoa(username);
  loginRequest.Password=btoa(password);
  $.ajax({
    url: apiUrl+"api/Login",
    method: 'POST',
    dataType: 'json',
    data: loginRequest,
    success: function (response) {
      if(response.IsAuthorized){
        var userinfo=new Object();
        userinfo.AccessToken=response.AccessToken;
        userinfo.UserName=response.UserName;
        localStorage.setItem("PluginUserInfo", JSON.stringify(userinfo));
        document.getElementById("InputUserName").value="";
        document.getElementById("InputPassword").value="";
        InitialLoadSettings();
      }
     else{
      document.getElementById("LoginError").innerHTML = "Invalid Credentails";
     }
    },
   fail : function( jqXHR, textStatus ) {
    console.log( "Request failed: " + textStatus );
   }
})
}

function GetContactInformation(item){
  var uinfo=userinfo();
  var reqobj=new Object();
  var fromDetails=GetEmailAddressString(item.from);
  reqobj.EmailId=fromDetails.EmailId;
  reqobj.ISComposeMode=false;
  $.ajax({
    url: apiUrl+"api/Contact/ContactInformation",
    method: 'POST',
    data: reqobj,
    headers:  
                {  
                    Authorization: 'Basic ' + btoa(uinfo.UserName + ':' + uinfo.AccessToken)  
                },
    success: function (response) {
      BuildInboxData(response,fromDetails);
    },
   fail : function( jqXHR, textStatus ) {
    console.log( "Request failed: " + textStatus );
   }
})
}

function SaveContactInformation(item){
  var uinfo=userinfo();
  var reqobj=new Object();
  var fromDetails=GetEmailAddressString(item.from);
  reqobj.EmailId=fromDetails.EmailId;
  var namearray=fromDetails.Name.split(" ");
  reqobj.FirstName=namearray[0];
  reqobj.LastName=namearray[1];
  reqobj.FullName=fromDetails.Name;
  $.ajax({
    url: apiUrl+"api/Contact/SaveContactInformation",
    method: 'POST',
    data: reqobj,
    headers:  
                {  
                    Authorization: 'Basic ' + btoa(uinfo.UserName + ':' + uinfo.AccessToken)  
                },
    success: function (response) {
      BuildInboxData(response,fromDetails);
    },
   fail : function( jqXHR, textStatus ) {
    console.log( "Request failed: " + textStatus );
   }
})
}

function BodyCallback(asyncResult) {
  serviceRequest.EmailBody = asyncResult.value;
  SaveEmailServiceRequest();
}


function attachmentTokenCallback(asyncResult, userContext) {
  if (asyncResult.status == "succeeded") {
      serviceRequest.attachmentToken = asyncResult.value;
      SaveEmailWithAttachments(Office.context.mailbox.item);
  }
  else {
      showToast("Error", "Could not get callback token: " + asyncResult.error.message);
  }
}

function SaveEmailServiceRequest(){
  var uinfo=userinfo();
  var attachment;
  for (var i = 0; i < Office.context.mailbox.item.attachments.length; i++) {
    attachment = Office.context.mailbox.item.attachments[i];
    attachment = attachment._data$p$0 || attachment.$0_0;

    if (attachment !== undefined) {
        serviceRequest.attachments[i] = JSON.parse(JSON.stringify(attachment));
    }
}
  $.ajax({
    url: apiUrl+"api/Common/SaveEmailWithAttachments",
    method: 'POST',
    data: serviceRequest,
    headers:  
                {  
                    Authorization: 'Basic ' + btoa(uinfo.UserName + ':' + uinfo.AccessToken)  
                },
    success: function (response) {
      if(response.IsSuccess){
        $('#exportEmailModal').modal('show');
      }
    },
   fail : function( jqXHR, textStatus ) {
    console.log( "Request failed: " + textStatus );
   }
})
}

function AttachmentCallback(result) {
if (result.value.length > 0) {
 for (i = 0 ; i < result.value.length ; i++) {
   result.asyncContext.currentItem.getAttachmentContentAsync(result.value[i].id, (r) => {
       setTimeout(function () {     
           var attachmentObj=new Object();
       attachmentObj.AttachmentContent=r.status=="succeeded"?r.value.content:"";
       var calculatedSize=FindAttachmentSize(r.value.content);
       attachmentObj.Name=result.value.filter(attachmnt=>(attachmnt.size===calculatedSize))[0].name; 
       serviceRequest.Attachments.push(attachmentObj);
         }, 1000);       
   });
 }
 setTimeout(function () { 
   
  makeServiceRequest();  }, 2000);        
}
else{
   makeServiceRequest();  
}
}

function SaveEmailWithAttachments(item){
  var emailFromDetails=GetEmailAddressString(item.from);
  serviceRequest.EmailFrom=emailFromDetails.EmailId;
  serviceRequest.EmailTo=buildMailAddressesString(item.to);
  serviceRequest.EmailSubject=item.subject;
  serviceRequest.EmailReceivedDate=item.dateTimeCreated.toLocaleString();
  serviceRequest.EmailCC=buildMailAddressesString(item.cc);
  Office.context.mailbox.item.body.getAsync(
    "text",
    { asyncContext: "This is passed to the callback" },
    BodyCallback);
}

function CreateNewTask(item){
  var uinfo=userinfo();
  var reqobj=new Object();
  var fromDetails=GetEmailAddressString(item.from);
  reqobj.EmailId=fromDetails.EmailId;
  reqobj.AssignedBy=uinfo.UserName;
  reqobj.AssignedTo=fromDetails.Name;
  reqobj.FundID=$('#drpdwnFund').val();
  reqobj.ActivityTypeId=$('#drpdwnType').val();
  var dueDate=reqobj.DueDateString=$('#inputDueDate').val()+"";
  dueDate=dueDate.split("-");
  reqobj.DueDateString=dueDate[1]+"/"+dueDate[0]+"/"+dueDate[2];
  reqobj.DiligenceID=$('#drpdwnChecklist').val();
  reqobj.DiligenceName=$("#drpdwnChecklist :selected").text();
  reqobj.Comments=$('#inputComment').val();
  $.ajax({
    url: apiUrl+"api/Inbox/CreateNewTask",
    method: 'POST',
    data: reqobj,
    headers:  
                {  
                    Authorization: 'Basic ' + btoa(uinfo.UserName + ':' + uinfo.AccessToken)  
                },
    success: function (response) {
      BuildTaskSection(response);
    },
   fail : function( jqXHR, textStatus ) {
    console.log( "Request failed: " + textStatus );
   }
})
}

function BuildTaskSection(response){
var taskList=response.TaskList;
if(taskList!=null && taskList.length>0){
      var tasks="";
    for(var i=0;i<taskList.length;i++){
      if(i%2==1){
        tasks+='<div style="background-color: #f0f0f5;" class="row"><div class="col-4">Type</div><div class="col-8">'+taskList[i].Type +'</div><div class="col-4">Date</div><div class="col-8">'+taskList[i].Date.split(" ")[0] +'</div><div class="col-4">Checklist</div><div class="col-8">'+taskList[i].Purpose +'</div><div class="col-4">Note</div><div class="col-8">'+taskList[i].Comments +'</div></div></br>';
      }
      else{
        tasks+='<div class="row"><div class="col-4">Type</div><div class="col-8">'+taskList[i].Type +'</div><div class="col-4">Date</div><div class="col-8">'+taskList[i].Date.split(" ")[0] +'</div><div class="col-4">Checklist</div><div class="col-8">'+taskList[i].Purpose +'</div><div class="col-4">Note</div><div class="col-8">'+taskList[i].Comments +'</div></div></br>';
      }
    }
    }
    else{
      tasks='<span>No Tasks</span>';
    }
    $('#divTasks').html(tasks);
    $("#AddNewTaskModal .close").click();
    $('#inputComment').val("");
}

function BuildInboxData(response,fromDetails){
  if(response.IsExistingContact){
    $("#spnFromFullName").text(response.ContactInfo.FullName);
    $("#spnFromPhone").text(response.AddressInfo.Phone1);
    $("#divFromPhone").css("display", "");
    $("#divExistingContact").css("display", "");
    $("#divAddToCRM").css("display", "none");
    $("#divExportEmail").css("display", "");
    if(response.FundInfoList!=null && response.FundInfoList.length>0){
      var Deals="";
    for(var i=0;i<response.FundInfoList.length;i++){
      Deals+='<span>'+ (i+1) +"."+response.FundInfoList[i].FundName + '</span></br>'
    }
    var fund="";
    response.FundInfoList.forEach(element => {
      fund+='<option value="' + element.FundId + '">' + element.FundName + '</option>';
    });
    $('#drpdwnFund').html(fund);
    $('#drpdwnFund').val('');
    }
    else{
      Deals='<span>No Deals</span>';
      $("#drpdwnFund").empty();
    }
    $('#divDeals').empty().append(Deals);
    BuildTaskSection(response);
    GetCreateTaskData();
   
  }else{
    $("#spnFromFullName").text(fromDetails.Name);
    $("#divExistingContact").css("display", "none");
    $("#divFromPhone").css("display", "none");
    $("#divAddToCRM").css("display", "");
    $("#divExportEmail").css("display", "none");
  } 
  $("#spnFromEmailId").text(fromDetails.EmailId); 
}

function APIlogout(){
  var uinfo=userinfo();
  var logoutRequest=new Object();
  logoutRequest.UserName=btoa(uinfo.UserName);
  $.ajax({
    url: apiUrl+"api/Logout",
    method: 'POST',
    dataType: 'json',
    data: logoutRequest,
    success: function (data) {
    },
   fail : function( jqXHR, textStatus ) {
    console.log( "Request failed: " + textStatus );
   }
})
$('#logoutModal').modal('hide');
localStorage.removeItem("PluginUserInfo");
$("#wrapper").slideUp();
$("#loginContainer").slideDown();
}


function makeServiceRequest() {
  var uinfo=userinfo();
    var attachment;
    xhr = new XMLHttpRequest();

    // Update the URL to point to your service location.
    //xhr.withCredentials = true;
    xhr.open("POST", apiUrl+"api/AttachmentService", true);
    //xhr.setRequestHeader("Access-Control-Allow-Origin", "http://localhost:63428/api/AttachmentService");
    xhr.setRequestHeader("Content-Type", "application/json; charset=utf-8","Authorization", "Basic " + btoa(uinfo.UserName + ':' + uinfo.AccessToken));
    xhr.onreadystatechange = SaveEmailWithAttachmentCallBack;

    // Translate the attachment details into a form easily understood by WCF.
    for (var i = 0; i < Office.context.mailbox.item.attachments.length; i++) {
        attachment = Office.context.mailbox.item.attachments[i];
        attachment = attachment._data$p$0 || attachment.$0_0;

        if (attachment !== undefined) {
            serviceRequest.attachments[i] = JSON.parse(JSON.stringify(attachment));
        }
    }

    // Send the request. The response is handled in the 
    // requestReadyStateChange function.
    xhr.send(JSON.stringify(serviceRequest));
};

function SaveEmailWithAttachmentCallBack(){

}
// Handles the response from the JSON web service.
function requestReadyStateChange() {
    if (xhr.readyState == 4) {
        if (xhr.status == 200) {
            var response = JSON.parse(xhr.responseText);
            if (!response.isError) {
                // The response indicates that the server recognized
                // the client identity and processed the request.
                // Show the response.
                var names = "<h2>Attachments processed: " + response.attachmentsProcessed + "</h2>";

                for (var i = 0; i < response.attachmentNames.length; i++) {
                    names += response.attachmentNames[i] + "<br />";
                }
                document.getElementById("names").innerHTML = names;
            } else {
                showToast("Runtime error", response.message);
            }
        } else {
            if (xhr.status == 404) {
                showToast("Service not found", "The app server could not be found.");
            } else {
                showToast("Unknown error", "There was an unexpected error: " + xhr.status + " -- " + xhr.statusText);
            }
        }
    }
};

// Shows the service response.
function showResponse(response) {
    showToast("Service Response", "Attachments processed: " + response.attachmentsProcessed);
}

// Displays a message for 10 seconds.
function showToast(title, message) {

    var notice = document.getElementById("notice");
    var output = document.getElementById('output');

    notice.innerHTML = title;
    output.innerHTML = message;

    $("#footer").show("slow");

    window.setTimeout(function () { $("#footer").hide("slow") }, 10000);
};

  // Take an array of AttachmentDetails objects and build a list of attachment names, separated by a line-break.
  function buildAttachmentsString(attachments) {
    if (attachments && attachments.length > 0) {
      var returnString = "";
      
      for (var i = 0; i < attachments.length; i++) {
        if (i > 0) {
          returnString = returnString + "<br/>";
        }
        returnString = returnString + attachments[i].name;
      }

      return returnString;
    }

    return "None";
  }

  // Format an EmailAddressDetails object as
  // GivenName Surname <emailaddress>

  function GetEmailAddressString(address) {
    var emailObj=new Object();
    emailObj.EmailId=address.emailAddress;
    emailObj.Name=address.displayName;
    return emailObj;
  }

  function buildEmailAddressString(address) {
    return address.displayName + " &lt;" + address.emailAddress + "&gt;";
  }

  function buildMailAddressesString(addresses) {
    if (addresses && addresses.length > 0) {
      var returnString = "";

      for (var i = 0; i < addresses.length; i++) {
        var emaiDetails=GetEmailAddressString(addresses[i]);
        returnString+=  emaiDetails.EmailId + ",";
      }

      return returnString;
    }

    return "None";
  }

  // Take an array of EmailAddressDetails objects and
  // build a list of formatted strings, separated by a line-break
  function buildEmailAddressesString(addresses) {
    if (addresses && addresses.length > 0) {
      var returnString = "";

      for (var i = 0; i < addresses.length; i++) {
        if (i > 0) {
          returnString = returnString + "<br/>";
        }
        returnString = returnString + buildEmailAddressString(addresses[i]);
      }

      return returnString;
    }

    return "None";
  }

  // Load properties from the Item base object, then load the
  // message-specific properties.
  function loadProps() {
    var item = Office.context.mailbox.item;
    item.body.getAsync(
      "text",
      { asyncContext: "This is passed to the callback" },
      function callback(result) {
        $('#mailContent').text(result.value);
      });
    
    $('#dateTimeCreated').text(item.dateTimeCreated.toLocaleString());
    $('#dateTimeModified').text(item.dateTimeModified.toLocaleString());
    $('#itemClass').text(item.itemClass);
    $('#itemId').text(item.itemId);
    $('#itemType').text(item.itemType);

    $('#message-props').show();

    $('#attachments').html(buildAttachmentsString(item.attachments));
    $('#cc').html(buildEmailAddressesString(item.cc));
    $('#conversationId').text(item.conversationId);
    $('#from').html(buildEmailAddressString(item.from));
    $('#internetMessageId').text(item.internetMessageId);
    $('#normalizedSubject').text(item.normalizedSubject);
    $('#sender').html(buildEmailAddressString(item.sender));
    $('#subject').text(item.subject);
    $('#to').html(buildEmailAddressesString(item.to));
  }

  // Helper function for displaying notifications
  function showNotification(header, content) {
    $("#notificationHeader").text(header);
    $("#notificationBody").text(content);
    messageBanner.showBanner();
    messageBanner.toggleExpansion();
  }
})();
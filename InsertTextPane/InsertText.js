// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.
/**$("#testspan").text("satz");
(function () {
  "use strict";
  var item;
  
  // The Office initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
    
    $(document).ready(function () {
      item=Office.context.mailbox.item;
      $('#insertText').click(insertText);
      $('#insertTOTO').click(AddEmailToTo);
      $('#insertTOBCC').click(AddEmailToBCC);
      $('#insertTOCC').click(AddEmailToCC);
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
      $("#drpdwnContacts").change(function() {
        GetContactInformation();
      });
      setTimeout(function(){
         
        testMethod();  
      }, 2000);
      InitialLoadSettings();
      MulitiSelectableDropDown();
  });
  };

  function testMethod(){
    var uinfo=userinfo();
    var reqobj=new Object();
    reqobj.s1="";
    $.ajax({
      url: "http://localhost:63428/api/AttachmentService/ProcessEmail",
      method: 'POST',
      data: reqobj,
      headers:  
                  {  
                      Authorization: 'Basic ' + btoa(uinfo.UserName + ':' + uinfo.AccessToken)  
                  },
      success: function (data) {
       console.log("Saved successfully");
      },
     fail : function( jqXHR, textStatus,err ) {
      alert( "Request failed: " + textStatus );
     }
  })
  }

  function GetContactInformation(){
    var uinfo=userinfo();
    var reqobj=new Object();
    var fromDetails=new Object();
    fromDetails.Name=$( "#drpdwnContacts option:selected" ).text();
    fromDetails.EmailId=$('#drpdwnContacts').val();
    reqobj.EmailId=$('#drpdwnContacts').val();
    $.ajax({
      url: "http://localhost:63428/api/Contact/ContactInformation",
      method: 'POST',
      data: reqobj,
      headers:  
                  {  
                      Authorization: 'Basic ' + btoa(uinfo.UserName + ':' + uinfo.AccessToken)  
                  },
      success: function (response) {
        BuildInboxData(response,fromDetails);
        $(".nUserVisibility").css("display","");
      },
     fail : function( jqXHR, textStatus ) {
      console.log( "Request failed: " + textStatus );
     }
  })
  }

  function BuildInboxData(response,fromDetails){
    if(response.IsExistingContact){
      $("#spnFromFullName").text(response.ContactInfo.FullName);
      $("#spnFromPhone").text(response.AddressInfo.Phone1);
      $("#divFromPhone").css("display", "");
      $("#divExistingContact").css("display", "");
      $("#divAddToButtons").css("display", "");
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

  function GetCreateTaskData(){
    var uinfo=userinfo();
    var reqobj=new Object();
    $.ajax({
      url: "http://localhost:63428/api/Inbox/CreateTaskData",
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
      }
      if(response.CheckLists!=null){
        var checkList="";
        response.CheckLists.forEach(element => {
          checkList+='<option value="' + element.KeywordID + '">' + element.Value + '</option>';
          $('#drpdwnChecklist').append('<option value="' + element.KeywordID + '">' + element.Value + '</option>');
        });
        $('#drpdwnChecklist').empty().append(checkList);
      }
      
      },
     fail : function( jqXHR, textStatus ) {
      console.log( "Request failed: " + textStatus );
     }
  })
  }

  function LoadAllContacts(item){
    var uinfo=userinfo();
    var reqobj=new Object();
    $.ajax({
      url: "http://localhost:63428/api/Contact/AllContactInformation",
      method: 'POST',
      data: reqobj,
      headers:  
                  {  
                      Authorization: 'Basic ' + btoa(uinfo.UserName + ':' + uinfo.AccessToken)  
                  },
      success: function (response) {
        var contacts="";
        response.ContactInformationList.forEach(element => {
          contacts+='<option value="' + element.Email1 + '">' + element.ContactName + '</option>';
        });
        $('#drpdwnContacts').html(contacts);
      },
     fail : function( jqXHR, textStatus ) {
      console.log( "Request failed: " + textStatus );
     }
  })
  }

  function APIlogout(){
    var uinfo=userinfo();
    var logoutRequest=new Object();
    logoutRequest.UserName=btoa(uinfo.UserName);
    $.ajax({
      url: "http://localhost:63428/api/Logout",
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
      url: "http://localhost:63428/api/Login",
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

function AddEmailToTo(){
  item.to.getAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed){
        write(asyncResult.error.message);
    }
    else {
        // Async call to get cc-recipients of the item completed.
        // Display the email addresses of the cc-recipients.
        var toObjects=asyncResult.value;
        var toobj1=new Object();
        toobj1.emailAddress="allieb@contoso.com";
        toobj1.displayName=""; 
        toObjects.push(toobj1);
        Office.context.mailbox.item.to.setAsync(toObjects, function(result) {
          if (result.error) {
              console.log(result.error);
              $("#testspan").text(JSON.stringify(result.error));
          } else {
              console.log("Recipients overwritten");
          }
      });
    }
});
}

  function AddEmailToCC(){
    item.cc.getAsync(function (asyncResult) {
      if (asyncResult.status == Office.AsyncResultStatus.Failed){
          write(asyncResult.error.message);
      }
      else {
          // Async call to get cc-recipients of the item completed.
          // Display the email addresses of the cc-recipients.
          var ccObjects=asyncResult.value;
          var ccobj1=new Object();
          ccobj1.emailAddress="allieb@contoso.com";
          ccObjects.push(ccobj1);
          ccobj1.displayName=""; 
          Office.context.mailbox.item.cc.setAsync(ccObjects, function(result) {
            if (result.error) {
                console.log(result.error);
                $("#testspan").text(JSON.stringify(result.error));
            } else {
                console.log("Recipients overwritten");
            }
        });
      }
  });
  }

  function AddEmailToBCC(){
    item.bcc.getAsync(function (asyncResult) {
      if (asyncResult.status == Office.AsyncResultStatus.Failed){
          write(asyncResult.error.message);
      }
      else {
          // Async call to get cc-recipients of the item completed.
          // Display the email addresses of the cc-recipients.
          var bccObjects=asyncResult.value;
          var bccobj1=new Object();
          bccobj1.emailAddress="allieb@contoso.com";
          bccobj1.displayName=""; 
          bccObjects.push(bccobj1);
          Office.context.mailbox.item.bcc.setAsync(bccObjects, function(result) {
            if (result.error) {
                console.log(result.error);
                $("#testspan").text(JSON.stringify(result.error));
            } else {
                console.log("Recipients overwritten");
            }
        });
      }
  });
  }

  function MulitiSelectableDropDown(){
    $('#drpdwnType').select2({
      width: '100%',
    });
    $('#drpdwnChecklist').select2({
      width: '100%',
    });
    $('#drpdwnFund').select2({
      width: '100%',
    });
    $('#drpdwnContacts').select2({
      width: '80%',
    });
    
  }

  function InitialLoadSettings(){
    $("#wrapper").slideDown();
    $("#containerDashboard").slideDown();
    $("#loginContainer").slideUp();
    $("#containerSettings").slideUp();
    LoadAllContacts();
  }


})(); **/



(function () {
  "use strict";
  var item;
  //var apiUrl="aspirevm10-20:9092/"; //live
  var apiUrl="http://localhost:63428/"; //local
  // The Office initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
    
    $(document).ready(function () {
      item=Office.context.mailbox.item;
      $("#idmainmenu").click(function(){
        InitialLoadSettings();
      });
      $("#btnAddtoTo").click(function(){
        CallAddEmailToTo();
      });
      $("#btnAddtoCc").click(function(){
        CallAddEmailToCc();
      });
      $("#btnAddtoBcc").click(function(){
        CallAddEmailToBcc();
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
      $("#drpdwnContacts").change(function() {
        GetContactInformation();
      });
      $("#btnLogout").click(function(){
        APIlogout();
      });
        testMethod();  
      InitialLoadSettings();
      MulitiSelectableDropDown();
  });
  };

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
       console.log("Saved successfully");
      },
     fail : function( jqXHR, textStatus,err ) {
      alert( "Request failed: " + textStatus );
     }
  })
  }

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
    $('#drpdwnContacts').select2({
      "placeholder": "Select Contact",
      width: '80%',
    });
    
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
          $("#spnFromFullName").text("");
          $("#spnFromEmailId").text("");
          $("#spnFromPhone").text("");
          $("#divFromPhone").css("display", "none");
          $("#divFundDetails").css("display", "none");
          $("#divExistingContact").css("display", "none");
          $("#divTaskDetails").css("display", "none");
          $("#divContactProfile").css("display", "none");
          $(".nUserVisibility").css("display","none");
          $("#divAddToButtons").css("display","none");
          
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

  function InitialLoadSettings(){
    $("#wrapper").slideDown();
    $("#containerDashboard").slideDown();
    $("#loginContainer").slideUp();
    $("#containerSettings").slideUp();
    LoadAllContacts();
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

  function GetContactInformation(){
    var uinfo=userinfo();
    var reqobj=new Object();
    var fromDetails=new Object();
    fromDetails.Name=$( "#drpdwnContacts option:selected" ).text();
    fromDetails.EmailId=$('#drpdwnContacts').val();
    reqobj.EmailId=$('#drpdwnContacts').val();
    reqobj.ISComposeMode=true;
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
        $(".nUserVisibility").css("display","");
      },
     fail : function( jqXHR, textStatus ) {
      console.log( "Request failed: " + textStatus );
     }
  })
  }

  function LoadAllContacts(item){
    var uinfo=userinfo();
    var reqobj=new Object();
    $.ajax({
      url: apiUrl+"api/Contact/AllContactInformation",
      method: 'POST',
      data: reqobj,
      headers:  
                  {  
                      Authorization: 'Basic ' + btoa(uinfo.UserName + ':' + uinfo.AccessToken)  
                  },
      success: function (response) {
        //$("#testspan").text(JSON.stringify(response.ContactInformationList));
        var contacts="";
        for(var i=0;i<response.ContactInformationList.length;i++){
          contacts+='<option value="' + response.ContactInformationList[i].Email1 + '">' + response.ContactInformationList[i].ContactName + '</option>';
        }
        $('#drpdwnContacts').html(contacts);
        $('#drpdwnContacts').val('');
      },
     fail : function( jqXHR, textStatus ) {
      console.log( "Request failed: " + textStatus );
     }
  })

  }

  function BuildInboxData(response,fromDetails){
    $("#divAddToButtons").css("display", "");
    if(response.IsExistingContact){
      $("#spnFromFullName").text(response.ContactInfo.FullName);
      $("#spnFromPhone").text(response.AddressInfo.Phone1);
      $("#divFromPhone").css("display", "");
      $("#divFundDetails").css("display", "");
      $("#divContactProfile").css("display", "");
      
      if(response.FundInfoList!=null && response.FundInfoList.length>0){
        var Deals="";
      for(var i=0;i<response.FundInfoList.length;i++){
        Deals+='<span>'+ (i+1) +"."+response.FundInfoList[i].FundName + '</span></br>'
      }
      var fund="";
      for(var i=0;i<response.FundInfoList.length;i++){
        fund+='<option value="' + response.FundInfoList[i].FundId + '">' + response.FundInfoList[i].FundName + '</option>';
      }
      $('#drpdwnFund').html(fund);
      $('#drpdwnFund').val('');
      }
      else{
        Deals='<span>No Deals</span>';
        $("#drpdwnFund").empty();
      }
      $('#divDeals').empty().append(Deals);
     
    }else{
      $("#spnFromFullName").text(fromDetails.Name);
      $("#divFundDetails").css("display", "none");
      $("#divFromPhone").css("display", "none");
      $("#divAddToCRM").css("display", "");
      $("#divExportEmail").css("display", "none");
    } 
      BuildTaskSection(response);
      GetCreateTaskData();
    $("#divTaskDetails").css("display", "");
    $("#spnFromEmailId").text(fromDetails.EmailId); 
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

    function GetCreateTaskData(){
      var uinfo=userinfo();
      var reqobj=new Object();
      $.ajax({
        url: apiUrl+"apiUrlapi/Inbox/CreateTaskData",
        method: 'POST',
        data: reqobj,
        headers:  
                    {  
                        Authorization: 'Basic ' + btoa(uinfo.UserName + ':' + uinfo.AccessToken)  
                    },
        success: function (response) {
        if(response.ActivityTypeList!=null){
          var activity="";
          for(var i=0;i<response.ActivityTypeList.length;i++){
            activity+='<option value="' + response.ActivityTypeList[i].KeywordValueID + '">' + response.ActivityTypeList[i].Value + '</option>';
          }
          $('#drpdwnType').empty().append(activity);
          $('#drpdwnType').val('');
        }
        if(response.CheckLists!=null){
          var checkList="";
          for(var i=0;i<response.CheckLists.length;i++){
            checkList+='<option value="' + response.CheckLists[i].KeywordID + '">' + response.CheckLists[i].Value + '</option>';
          }
          $('#drpdwnChecklist').empty().append(checkList);
          $('#drpdwnChecklist').val('');
        }
        
        },
       fail : function( jqXHR, textStatus ) {
        console.log( "Request failed: " + textStatus );
       }
    })
    }

    function CreateNewTask(item){
      var uinfo=userinfo();
      var reqobj=new Object();
      var fromDetails=new Object();
      fromDetails.Name=$( "#drpdwnContacts option:selected" ).text();
      fromDetails.EmailId=$('#drpdwnContacts').val();
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
    function CallAddEmailToTo(){
    var mailList=[];
    var mail=new Object();
    mail.emailAddress=$('#drpdwnContacts').val();
    mail.displayName=$( "#drpdwnContacts option:selected" ).text();
    mailList.push(mail);
    AddEmailToTo(mailList);
    }

    function CallAddEmailToCc(){
      var mailList=[];
      var mail=new Object();
      mail.emailAddress=$('#drpdwnContacts').val();
      mail.displayName=$( "#drpdwnContacts option:selected" ).text();
      mailList.push(mail);
      AddEmailToCC(mailList);
      }

      function CallAddEmailToBcc(){
        var mailList=[];
        var mail=new Object();
        mail.emailAddress=$('#drpdwnContacts').val();
        mail.displayName=$( "#drpdwnContacts option:selected" ).text();
        mailList.push(mail);
        AddEmailToBCC(mailList);
        }

    function AddEmailToTo(mailList){
      item.to.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get cc-recipients of the item completed.
            // Display the email addresses of the cc-recipients.
            var toObjects=asyncResult.value;
            for(var i=0;i<mailList.length;i++){
              toObjects.push(mailList[i]);
            }
            Office.context.mailbox.item.to.setAsync(toObjects, function(result) {
              if (result.error) {
                  console.log(result.error);
                  $("#testspan").text(JSON.stringify(result.error));
              } else {
                  console.log("Recipients overwritten");
              }
          });
        }
    });
    }

    function AddEmailToCC(mailList){
      item.cc.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get cc-recipients of the item completed.
            // Display the email addresses of the cc-recipients.
            var ccObjects=asyncResult.value;
            for(var i=0;i<mailList.length;i++){
              ccObjects.push(mailList[i]);
            }
            Office.context.mailbox.item.cc.setAsync(ccObjects, function(result) {
              if (result.error) {
                  console.log(result.error);
                  $("#testspan").text(JSON.stringify(result.error));
              } else {
                  console.log("Recipients overwritten");
              }
          });
        }
    });
    }

    function AddEmailToBCC(mailList){
      item.bcc.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get cc-recipients of the item completed.
            // Display the email addresses of the cc-recipients.
            var bccObjects=asyncResult.value;
            for(var i=0;i<mailList.length;i++){
              bccObjects.push(mailList[i]);
            }
            Office.context.mailbox.item.bcc.setAsync(bccObjects, function(result) {
              if (result.error) {
                  console.log(result.error);
                  $("#testspan").text(JSON.stringify(result.error));
              } else {
                  console.log("Recipients overwritten");
              }
          });
        }
    });
    }

  

})();
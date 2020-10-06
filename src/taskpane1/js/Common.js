(function($) {
  "use strict"; // Start of use strict

$.ajaxSetup({
               error: function (XMLHttpRequest, textStatus, errorThrown)
   {
     if(errorThrown=="Unauthorized"){
      RedirectToLogin();
     }
               }
          });
		  
function RedirectToLogin(){
  $("#wrapper").css("display", "none");
  $("#loginContainer").css("display", "");
}

 



})(jQuery); // End of use strict

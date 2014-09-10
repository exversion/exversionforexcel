/// <reference path="../App.js" />
/*global app*/

(function () {
	'use strict';

	// The initialize function must be run each time a new page is loaded
	Office.initialize = function (reason) {
		$(document).ready(function () {
			app.initialize();

			//$('#get-data-from-selection').click(getDataFromSelection);
			if(sessionStorage.getItem("exv-token")){
				//if there are unsynced changes show sync button
				var changes = sessionStorage.getItem('exv-changes');
				if(changes){
					$('#work').append('<br><a href="../sync/sync.html" class="btn btn-warning btn-block"><i class="fa fa-retweet fa-lg"></i> Sync changes with Exversion</a>');	
				}
				advance('#login','#work');
				$('#logout').show();
			}
	
			$('#login_submit').click(function(){
				var u = $('#login_username').val();
				var p = $('#login_password').val();
			if(u && p){
				var button = $(this).html();
				$(this).html('<img src="../../images/ajax-loader.gif">');
				//Send Login Request
				$.ajax({
				type: 'POST',
				data: {'username':u,'password':p},
				url:'https://exversion.com/api/v1/excel/login/',
				success: function(data){
					$(this).html(button);
					if(data.status == 200){
						app.closeNotification();
					//Exversion must reply with Token, Key, and private repo status (0-5 repos or unlimited) 
					sessionStorage.setItem("exv-token", data.body[0].token);
					sessionStorage.setItem("exv-key", data.body[0].key);
					sessionStorage.setItem("exv-private", data.body[0].private);
					sessionStorage.setItem("exv-access_list",JSON.stringify(data.body[0].datasets));
					advance('#login','#work');
					$('#logout').show();
					$('.user_display').html(u);
					}else{
						app.showNotification(data.message);
					}
				}
			});
			}
		});
		
		$('#register_id').click(function(){
			advance('#login','#register');
		});
		
		$('#register_submit').click(function(){
			//Check that all values are present
			var u = $('#register_username').val();
			var p = $('#register_password').val();
			var p_again = $('#register_password_confirm').val();
			var e = $('#register_email').val();
			if(u && p && p_again && e){
				if(!validEmail(e)){
					app.showNotification("This does not appear to be a valid email");
				}
				else if(p == p_again){
					//Send registration request
					var button = $(this).html();
				$(this).html('<img src="../../images/ajax-loader.gif">');
					$.ajax({
				type: 'POST',
				data: {'username':u,'password':p, 'confirm_password':p_again, 'email':e},
				url:'https://exversion.com/api/v1/excel/register/',
				success: function(data){
					$(this).html(button);
					if(data.status == 200){
						app.closeNotification();
					//Exversion must reply with Token, Key, and private repo status (0-5 repos or unlimited) 
					sessionStorage.setItem("exv-token", data.body[0].token);
					sessionStorage.setItem("exv-key", data.body[0].key);
					sessionStorage.setItem("exv-private", data.body[0].private);
					advance('#register','#work');
					$('#logout').show();
					$('.user_display').html(u);
					}else{
						app.showNotification(data.message);
					}
				}
			});
				}else{
					app.showNotification('Password confirm does not match');
				}
			}else{
				app.showNotification('Please fill in all fields');
			}
		});
		
		$('#register_redo').click(function(){
			advance('#register','#login');
		});
		
		});
	};
})();

function validEmail(email){
	var re = /^(([^<>()[\]\\.,;:\s@\"]+(\.[^<>()[\]\\.,;:\s@\"]+)*)|(\".+\"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
    return re.test(email);
}
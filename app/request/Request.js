/// <reference path="../App.js" />
/*global app*/

(function () {
	'use strict';

	// The initialize function must be run each time a new page is loaded
	Office.initialize = function (reason) {
		$(document).ready(function () {
			app.initialize();

			if(!sessionStorage.getItem("exv-token")){
				window.location = '../home/home.html';
				}
	
			$('#request_submit').click(function(){
				var d = $('#description').val();
				var s = $('#sources').val();
				var u = $('#uses').val();
				
			if(d && u && s && d != 'Describe the data you need...'){
				var button = $(this).html();
				$(this).html('<img src="../../images/ajax-loader.gif">');
				//Send Login Request
				$.ajax({
				type: 'POST',
				data: {'access_token':sessionStorage.getItem("exv-token"), 'request':{'description':d,'sources':s, 'uses':u}},
				url:'https://exversion.com/api/v1/request/create/',
				success: function(data){
					$(this).html(button);
					if(data.status == 200){
						app.closeNotification();
					$('#request_link').attr('href', 'https://www.exversion.com/requests/view/'+data['body'][0]['id']);
					advance('#request','#complete');
					}else{
						app.showNotification(data.message);
					}
				}
			});
			}
		});
		
		$('#request_redo').click(function(){
			window.location = '../home/home.html';
		});
		
		});
	};
})();

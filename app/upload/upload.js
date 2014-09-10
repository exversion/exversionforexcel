/// <reference path="../App.js" />
/*global app*/
var dataset;
var metadata;

(function () {
	'use strict';

	// The initialize function must be run each time a new page is loaded
	Office.initialize = function (reason) {
		$(document).ready(function () {
			app.initialize();
			
			//If number of private repos is unlimited
			var repo = sessionStorage.getItem("exv-private");
			if(repo == 'unlimited'){
				$('#free_account').remove();
			}
			else if(repo == 0){
				$('#free_account').show();
				$('#private').hide();
			}
			else{
				if(repo == 1){
					$('#repo_plural').html('repo');
				}
				$('#private_repo_available').html(repo);
				$('#free_account').show();
			}
			
			//Is anything selected?
			if(test_select() === false){
				app.showNotification('Please select some data first');
			}
			$('#confirm_submit').click(function(){
				//Create a dataset
				advance('#confirm','#createSet');
			});
			
			$('#confirm_redo').click(function(){
				if(test_select() === false){
					app.showNotification('Please select some data first');
				}
			});
			
			$('#create_set').click(function(){
				metadata = {'name': $('#create_name').val(),
							'description':$('#create_description').val(),
							'private': $('#create_private').is(':checked')};
			if(form_validation(metadata)){
				//Send Request to Exversion
				var button = $(this).html();
				$(this).html('<img src="../../Images/ajax-loader.gif">');
				$.ajax({
				type: 'POST',
				data: {"access_token":sessionStorage.getItem('exv-token'), "name":$('#create_name').val(), "description":$('#create_description').val(),"source_url":"N/A", "org":0,"source_author":"N/A","source_date":"N/A", "source_contact":"N/A", "private":($('#create_private').is(':checked') ? 1 : 0)},
				url:'https://exversion.com/api/v1/dataset/create/',
				success: function(data){
					$(this).html(button);
					if(data.status == 200){
						app.closeNotification();
						dataset = data.body[0].dataset;
						//Read data
						//Send it
						//Bind selection
						readData();
					}else{
						app.showNotification(data.message);
					}
				}
			});
			}
		});
		});
	}
})();

function send_data(metadata, data){
	advance('#createSet', '#Sending');
	var button = $(this).html();
	$(this).html('<img src="../../Images/ajax-loader.gif">');
	//Push data to exversion
	$.ajax({
		type: 'POST',
		data: {"access_token":sessionStorage.getItem('exv-token'), "dataset":dataset,"data":data, "meta":1},
		url:'https://exversion.com/api/v1/dataset/push/',
		success: function(data){
			$(this).html(button);
			if(data.status == 200){
				app.closeNotification();
				var sets = [];
				var hashes = [];
				sets = sets.concat(data.body.inserted);
				sets = sets.concat(data.body.updated);
				for(var d in sets){
					hashes[sets[d]['*'+dataset+'_row']] = {'_id':sets[d]['_id'], 'hash': sets[d]['*'+dataset+'_hash']};
				}
				//Store dataset name and sheetid in Excel
				datasets = app.getDatasets(dataset);
				datasets[dataset] = hashes;
				app.setDatasets(datasets);
				
				//bind the selection
				Office.context.document.bindings.addFromSelectionAsync("table", { id: dataset }, 
        			function (asyncResult) {
            		if (asyncResult.status === "failed") {
						app.showNotification('Error:', asyncResult.error.message);
					}else{
						readBoundDataSend(dataset);
					}
					});
				
				//Load success page
				$('#dataset_link').attr('href','https://exversion.com/data/view/'+dataset);
				$('#dataset_link').html(dataset);
				advance('#Sending', '#complete');
			}else{
				app.showNotification(data.message);
			}
		}
	});
}

function form_validation(data){
	for(var k in data){
		if(!data[k] && k != 'private'){
			$('#create_'+k).addClass('has-error');
			return false;
		}
	}
	return true;
}

function test_select(){
		Office.context.document.getSelectedDataAsync(Office.CoercionType.Matrix,
			function (result) {
				if (result.status === Office.AsyncResultStatus.Succeeded) {
					//Is the matrix empty?
					if(test_empty_data(result.value)){
						//Assume result.value[0] is header
						var header = result.value[0];
						var items = result.value
						items.shift();
						//Zip remaining items
						var objects = zip_object(header, items);
						//Display objects
						$('#preview').html(JSON.stringify(objects.slice(0,2), null, 4));
						$('#confirm').show();
					
					}else{
						return false;
					}
				} else {
					return false;
				}
			}
		);
	}
	
function test_empty_data(array){
		for(var a in array){
			if(a instanceof Array){
				test_empty_data(a);
			}else if (a){
				return true;
			}
		}
		return false;
	}

function zip_object(keys, values){
		if (keys == null) return [];
		var result = [];
		for(var i = 0, l = values.length; i < l; i++){
			var obj = {};
			var v = values[i];
			for (var j = 0, len = v.length; j < len; j++) {
				obj[keys[j]] = v[j];
			}
			result.push(jQuery.extend({},obj));
		}
		return result;
	};
	
function zip_object_send(keys, values){
		if (keys == null) return close_out();
		
		//Store columns for later
		if(Office.context.document.settings.get('schemas')){
			schemas = app.getSchema(dataset);
			schemas[dataset] = keys;
		}else{
			schemas = {};
			schemas[dataset] = keys;
		}
		app.setSchema(schemas);
		
		var result = [];
		for(var i = 0, l = values.length; i < l; i++){
			var obj = {};
			obj['*'+dataset+'_row'] = i;
			obj['*'+dataset+'_hash'] = generate_hash(all_string(values[i]));
			var v = values[i]; 
			for (var j = 0, len = v.length; j < len; j++) {
				obj[keys[j]] = v[j];
			}
			result.push(jQuery.extend({},obj));
		}
		
		return send_data(metadata, result);
	};

function close_out(){
	$('#dataset_link').attr('href','https://exversion.com/data/view/'+dataset);
	//$('#dataset_link').html(dataset);
	advance('#Sending', '#complete');
}
	
function bindData(dataset, callback) {
	Office.context.document.bindings.addFromSelectionAsync("table", { id: dataset }, 
        function (asyncResult) {
            if (asyncResult.status === "failed") {
				app.showNotification('Error:', asyncResult.error.message);
            }else{
				//Add event handlers
				readBoundDataSend(dataset, callback);
    }
		});
	}

function readData() {
	Office.context.document.getSelectedDataAsync(Office.CoercionType.Matrix, function (asyncResult) {
        if (asyncResult.status === "failed") {
            app.showNotification('Error: ' + asyncResult.error.message);
        } 
        else{
			var header = asyncResult.value[0];
			var items = asyncResult.value;
			items.shift();
             return zip_object_send(header, items);
        }
    });
    }
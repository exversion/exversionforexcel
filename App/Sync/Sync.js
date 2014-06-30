/// <reference path="../App.js" />
/*global app*/

(function () {
	'use strict';
	// The initialize function must be run each time a new page is loaded
	Office.initialize = function (reason) {
		$(document).ready(function () {
			app.initialize();
			advance('synced','change');
			var changes = JSON.parse(sessionStorage.getItem('exv-changes'));
			display_well(changes);
			
			$('#change_submit').click(function(){
				var changes = JSON.parse(sessionStorage.getItem('exv-changes'));
				run_sync(changes);
			});
		
		
		$('#change_edit').click(function(){
			var changes = JSON.parse(sessionStorage.getItem('exv-changes'));
			var edit_list = '';
			for(var i in changes){
				if(changes[i].hasOwnProperty('delete')){
						edit_list += '<div><p>Remove datapoint with id '+changes[i]['_id']+'</p><div style="text-align:center"><button row_id="'+i+'" class="btn btn-warning edit_ch" disabled style="width:20%;"><i class="fa fa-edit fa-lg"></i></button><button row_id="'+i+'" class="btn btn-danger del_ch" style="width:20%;margin-left:20px;"><i class="fa fa-times fa-lg"></i></button></div><br style="clear:both;"></div>';
					continue;
					}
				if(!changes[i]['_id']){
					if(changes[i].hasOwnProperty('data')){
						var change_str = JSON.stringify(changes[i]['data']);
						edit_list += '<div><p>'+change_str.substring(0,50)+'</p><div style="text-align:center"><button row_id="'+i+'" class="btn btn-warning edit_ch" style="width:20%;"><i class="fa fa-edit fa-lg"></i></button><button row_id="'+i+'" class="btn btn-danger del_ch" style="width:20%;margin-left:20px;"><i class="fa fa-times fa-lg"></i></button></div><br style="clear:both;"></div>';
					continue;
					}
					//column change
					if(changes[i]['type'] == 'delete'){
						edit_list += '<div><p>Remove column '+changes[i]['column_name']+'</p><div style="text-align:center"><button row_id="'+i+'" class="btn btn-warning edit_ch_col" style="width:20%;"><i class="fa fa-edit fa-lg"></i></button><button row_id="'+i+'" class="btn btn-danger del_ch" style="width:20%;margin-left:20px;"><i class="fa fa-times fa-lg"></i></button></div><br style="clear:both;"></div>';
					}else{
						edit_list += '<div><p>Add column '+changes[i]['column_name']+'</p><div style="text-align:center"><button row_id="'+i+'" class="btn btn-warning edit_ch_col" style="width:20%;"><i class="fa fa-edit fa-lg"></i></button><button row_id="'+i+'" class="btn btn-danger del_ch" style="width:20%;margin-left:20px;"><i class="fa fa-times fa-lg"></i></button></div><br style="clear:both;"></div>';
					}
				}else{
				//otherwise pull up changes
					var change_str = JSON.stringify(changes[i]['changes']);
					edit_list += '<div><p>'+change_str.substring(0,50)+'</p><div style="text-align:center"><button row_id="'+i+'" class="btn btn-warning edit_ch" style="width:20%;"><i class="fa fa-edit fa-lg"></i></button><button row_id="'+i+'" class="btn btn-danger del_ch" style="width:20%;margin-left:20px;"><i class="fa fa-times fa-lg"></i></button></div><br style="clear:both;"></div>';
				}
			}
			$('#sync_edit_inner').html(edit_list);
			advance('#change','#sync_edit');
		});
		
		$('#sync_edit').on('click','.edit_ch_col', function(){
			var row_id = $(this).attr('row_id');
			var changes = JSON.parse(sessionStorage.getItem('exv-changes'));
			var item = '<div><input type="text" disabled class="form-control obj_params" value="column_name" style="width:40%; float:left;"/><input type="text" id="column_name" value="'+changes[row_id]['column_name']+'" class="filter_value form-control" style="width:54%; float:left; margin-left:2px;"/><br style="clear:both;"></div>';
			$('#edit_item_col_submit').attr('row_id',row_id);
			$('#sync_edit_item_col_inner').html(item);
			advance('#sync_edit', '#sync_edit_item_col');
		});
		
		$('#edit_item_col_submit').click(function(){
			var row_id = $(this).attr('row_id');
			var changes = JSON.parse(sessionStorage.getItem('exv-changes'));
			changes[row_id]['column_name'] = $('#column_name').val();			
			sessionStorage.setItem('exv-changes', JSON.stringify(changes));
			//Reset well
			display_well(changes);
			advance('#sync_edit_item_col', '#change')
		});
		
		$('#changes_raw').click(function(){
			var changes = sessionStorage.getItem('exv-changes');
			$('#changes_raw_text').val(changes);
			advance('#sync_edit','#sync_edit_raw');
		});
		
		$('#sync_edit').on('click','.edit_ch', function(){
			//Load edit item form
			var item = '';
			var row_id = $(this).attr('row_id');
			var changes = JSON.parse(sessionStorage.getItem('exv-changes'));
			var restricted = ["_id","_hash","sets","extension"];
			if(changes[row_id].hasOwnProperty('changes')){
				for(var k in changes[row_id]['changes']){
					if($.inArray(k, restricted) == -1 && k.substring(0,1) != '*'){
						item += '<div><input type="text" disabled class="form-control obj_params" value="'+k+'" style="width:30%; float:left;"/><input type="text" id="'+k+'" value="'+changes[row_id]['changes'][k]+'" class="filter_value form-control" style="width:54%; float:left; margin-left:2px;"/><br style="clear:both;"></div>';
					}
				}
			}
			if(changes[row_id].hasOwnProperty('data')){
				for(var k in changes[row_id]['data']){
					if($.inArray(k, restricted) == -1 && k.substring(0,1) != '*'){
						item += '<div><input type="text" disabled class="form-control obj_params" value="'+k+'" style="width:30%; float:left;"/><input type="text" id="'+k+'" value="'+changes[row_id]['data'][k]+'" class="filter_value form-control" style="width:54%; float:left; margin-left:2px;"/><br style="clear:both;"></div>';
					}
				}
			}
			$('#edit_item_submit').attr('row_id',row_id);
			$('#sync_edit_item_inner').html(item);
			advance('#sync_edit', '#sync_edit_item');
		});
		
		$('#sync_edit').on('click','.del_ch', function(){
			var row_id = $(this).attr('row_id');
			var changes = JSON.parse(sessionStorage.getItem('exv-changes'));
			var new_changes = []
			for(var i in changes){
				if(i !=  row_id){
					new_changes.push(changes[i]);
				}
			}
			sessionStorage.setItem('exv-changes', JSON.stringify(new_changes));
			//Reset well
			display_well(new_changes);
			advance('#sync_edit', '#change')
		});
		
		$('#change_raw_submit').click(function(){
			var new_changes = $('#changes_raw_text').val();
			//Is this string JSON?
			try {
        		new_changes = JSON.parse(new_changes);
    		} catch (e) {
        		app.showNotification('Error: Changes not valid JSON');
				return;
    		}
			
			//Does this follow the correct format?
			for(var i in new_changes){
				var dataset = new_changes[i]['dataset'];
				if($.inArray(dataset, bindings) == -1){
					app.showNotification('Error: Changes dataset id '+new_changes[i]['dataset']+' not part of this document.');
					return;
				}
				
				if(!new_changes[i]['_id']){
					if(new_changes[i]['column_name'] && new_changes[i]['type']){
					}else if(new_changes[i]['data']){
					}else{
						app.showNotification('Error: Change number '+i+' not properly formatted');
						return;
					}
				}
			}
			
			run_sync(new_changes);	
			
		});
		
		$('#edit_item_submit').click(function(){
			var row_id = $(this).attr('row_id');
			var changes = JSON.parse(sessionStorage.getItem('exv-changes'));
			var obj = {};			
			//Collect and assemble obj
			$('.obj_params').each(function(){
				var obj_key = $(this).val();
				obj[obj_key] = $('#'+obj_key).val();
			});
			if(changes[row_id].hasOwnProperty('changes')){
				changes[row_id]['changes'] = obj;
			}else{
				changes[row_id]['data'] = obj;	
			}			
			sessionStorage.setItem('exv-changes', JSON.stringify(changes));
			//Reset well
			display_well(changes);
			advance('#sync_edit_item', '#change')
		});
		
		$('#sync_off').click(function(){
			//Build dataset list
			var sync_list = '';
			Office.context.document.bindings.getAllAsync(function (asyncResult) {
        		for (var i in asyncResult.value) {
					sync_list += '<div style="margin-bottom:5px;"><a href="https://www.exversion.com/data/view/'+asyncResult.value[i].id+'" target="new" class="btn btn-transparent">'+asyncResult.value[i].id+'</a><button dataset_id="'+asyncResult.value[i].id+'" class="btn btn-danger delete_binding" style="width:20%;margin-left:20px;"><i class="fa fa-times fa-lg"></i></button></div>';
				}
				$('#syncing_off_sets').html(sync_list);
				advance('#change','#syncing_off');
			});
		});
		
		$('#syncing_off_sets').on('click','.delete_binding', function(){
			var dataset = $(this).attr('dataset_id');
			Office.context.document.bindings.releaseByIdAsync(dataset, function (asyncResult) {});
			$(this).parent().remove();
		});
		
		$('#syncing_off_all').click(function(){
			Office.context.document.bindings.getAllAsync(function (asyncResult) {
        		for (var i in asyncResult.value) {
					Office.context.document.bindings.releaseByIdAsync(asyncResult.value[i].id, function (asyncResult) {});
				}
				$('#sync_off').attr('disabled',true);
				advance('#syncing_off','#change');
			});
		});
		
		/*$('#add_change').click(function(){
			advance('#sync_edit','#sync_add_item');
		});
		
		$('#add_item_submit').click(function(){
			//Grab the highlighted area
			Office.context.document.getSelectedDataAsync(Office.CoercionType.Matrix,
			function (result) {
				if (result.status === Office.AsyncResultStatus.Succeeded) {
					var items = result.value;
					
					}
				});
			});*/
		
		$('#change_redo').click(function(){
			advance('#sync_edit','#change');
		});
		
		$('#change_raw_redo').click(function(){
			advance('#sync_edit_raw','#sync_edit');
		});
		
		$('#edit_item_redo').click(function(){
			advance('#sync_edit_item','#sync_edit');
		});
		
		$('#edit_item_col_redo').click(function(){
			advance('#sync_edit_item_col','#sync_edit');
		});
		
		$('#syncing_off_redo').click(function(){
			advance('#syncing_off','#change');
		});
		
		$('#add_item_redo').click(function(){
			advance('#sync_add_item','#sync_edit')
		});
		
		});
	};
})();

function getReady() {
	var deferredReady = $.Deferred();
	$(document).ready(function() {
		deferredReady.resolve();
		});
	return deferredReady.promise();
}

function addAllData(position, stop, keys, add_data){
	var d = $.Deferred();
	var msgs = {};
	$.ajax({
		type: 'POST',
		url:'https://exversion.com/api/v1/dataset/push/', 
		data:{'access_token':sessionStorage.getItem('exv-token'), 'dataset':keys[position], 'data':add_data[keys[position]], 'meta':1},
		success: function(msg){
			msgs['success'] = true;
			msgs[keys[position]] = msg;
			if(position+1 != stop){
				addAllData(position+1, stop, keys, add_data);
			}else{
				d.resolve(msgs);
			}
		},
		error: function(msg){
			msgs['success'] = false;
			msgs['error'] = keys[position];
			d.resolve(msgs);
		}
		});
	return d.promise();
}

function blankEvent(){
	return true;
}

function reconstitute_adds(add_data){
	var rechanges = [];
	for(var i in add_data){
		for(var d in add_data[i]){
			rechanges.push({'_id':null,'dataset':i,'data':add_data[i][d]});
		}
	}
	return rechanges;
	
}

function display_well(changes){
	$('#changes').html(JSON.stringify(changes, null, 2));
	return true;
}

function assign_hashes(responses){
	var datasets = app.getDatasets();
	for(var d in responses){
		if(d == 'success' || d == 'error'){
			continue;
		}
		var rows = responses[d]['body']['inserted'];
		rows = rows.concat(responses[d]['body']['updated']);
		var hashes = datasets[d];
		//Find hashes without ids
		for(var h in hashes){
			if(!hashes[h]['_id']){
				//Find the row with the right row number
				for(var rw in rows){
					if(rows[rw]['*'+d+'_row'] == h){
						//Assign the id
						hashes[h]['_id'] = rows[rw]['_id'];
						break;
					}
				}
			}
		}
	datasets[d] = hashes;
	}
	//Reset hashes
	app.setDatasets(datasets);
}

function run_sync(changes){
	var schema_changes = [];
	var add_data = {};
	var deleted_data = [];
	var data_changes = [];
				//Sort changes into the appropriate requests
				for( var i in changes){
					if(!changes[i]['_id'] && changes[i].hasOwnProperty('data')){
						if(add_data.hasOwnProperty(changes[i]['dataset'])){
							add_data[changes[i]['dataset']].push(changes[i]['data']);
						}else{
							add_data[changes[i]['dataset']] = [changes[i]['data']];
						}
					}else if(!changes[i]['_id']){
						schema_changes.push(changes[i]);
					}else if(changes[i].hasOwnProperty('delete')){
						deleted_data.push(changes[i]);
					}else{
						data_changes.push(changes[i]);
					}
				}
				//Run schema changes first
				//{"access_token":"", "changes":[{"dataset":"", "type":"delete", "column_name":"name"}]}
				var active_panel = $('.active')[0].id;
				var add_data_keys = Object.keys(add_data);
				
				if(schema_changes.length > 0){
					var schemaRequest = $.ajax({type: 'POST', url:'https://exversion.com/api/v1/dataset/schema/', data:{'access_token':sessionStorage.getItem('exv-token'), 'changes':schema_changes}});
					var dataset = schema_changes[0]['dataset'];
				}else{
					var schemaRequest = blankEvent();
				}
				
				if(data_changes.length > 0){
	    			var dataRequest = $.ajax({type: 'POST', url:'https://exversion.com/api/v1/dataset/edit/', data:{'access_token':sessionStorage.getItem('exv-token'), 'edits':data_changes}});
					var dataset = data_changes[0]['dataset'];
				}else{
					var dataRequest = blankEvent();
				}
				
				if(add_data_keys.length > 0){
					var addRequest = addAllData(0, add_data_keys.length, add_data_keys, add_data);
				}else{
					var addRequest = blankEvent();
				}
				
				if(deleted_data.length > 0){
					var deleteRequest = $.ajax({type: 'POST', url:'https://exversion.com/api/v1/dataset/delete/', data:{'access_token':sessionStorage.getItem('exv-token'), 'deletes':deleted_data}});
					var dataset = deleted_data[0]['dataset'];
				}else{
					var deleteRequest = blankEvent();
				}
				
				$.when( getReady(), schemaRequest, dataRequest, addRequest, deleteRequest).done( function( readyResponse, schemaResponse, dataResponse, addResponse, deleteResponse ) {
					//Check for failures
					var old_changes = [];
					if(typeof(dataset) != 'undefined'){
						$("#dataset_link").attr("href", "https://www.exversion.com/data/view/"+dataset);
						$("#dataset_link").text(dataset);
					}else{
						$('#synced_text').html('<strong>Data has been added to the appropriate repos.</strong>');
					}
					var response = '';
					if(typeof(schemaResponse[0]) != 'undefined' && schemaResponse[0]['status'] == 200){
						response += '<p>Changes to schema made successfully</p>';
					}else if(schemaResponse == true){
						response += '<p>No changes to schema detected</p>';
					}else{
						response += '<p>Changes to schema failed: '+schemaResponse[0]['message']+'</p>';
						//Perserve these changes for later syncing
						old_changes = old_changes.concat(schema_changes);
					}
					
					if(typeof(dataResponse[0]) != 'undefined' && dataResponse[0]['status'] == 200){
						response += '<p>Edits to existing data made successfully</p>';
					}else if(dataResponse == true){
						response += '<p>No changes to existing data detected</p>';
					}else{
						response += '<p>Changes to existing data failed: '+dataResponse[0]['message']+'</p>';
						//Perserve these changes for later syncing
						old_changes = old_changes.concat(data_changes);
					}
					
					if(typeof(deleteResponse[0]) != 'undefined' && deleteResponse[0]['status'] == 200){
						response += '<p>Data deleted successfully</p>';
					}else if(deleteResponse == true){
						response += '<p>No delete requests were detected</p>';
					}else{
						response += '<p>Data delete failed: '+deleteResponse[0]['message']+'</p>';
						//Perserve these changes for later syncing
						old_changes = old_changes.concat(deleted_data);
					}
					
					if(addResponse.hasOwnProperty('success') && addResponse['success'] == true){
						response += '<p>Data added successfully</p>';
						//assign hashes
						assign_hashes(addResponse);
					}else if(addResponse == true){
						response += '<p>No data to be added detected</p>';
					}else{
						response += '<p>Adding data failed for dataset '+addResponse['error']+'</p>';
						//Perserve these changes for later syncing
						var changes_add = reconstitute_adds(add_data);
						old_changes = old_changes.concat(changes_add);
						
						//assign whatever hashes were successful
						assign_hashes(addResponse);
					}
					$('#sync_report').html(response);
					advance('#'+active_panel,'#synced');
					//Clear changes
					sessionStorage.setItem('exv-changes', JSON.stringify(old_changes));
  				});
				
}
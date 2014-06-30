/* Common app functionality */
var bind_init = true;
var bindings = [];
var added_rows = [];

var app = (function () {
	'use strict';

	var app = {};

	// Common initialization function (to be called from each page)
	app.initialize = function () {
		
		//Need to reload the listeners because this API is dumb
		Office.context.document.bindings.getAllAsync(function (asyncResult) {
        for (var i in asyncResult.value) {
			bind_init = true;
            var bindingString = asyncResult.value[i].id;
			Office.select("bindings#"+asyncResult.value[i].id).addHandlerAsync("bindingDataChanged", bindingChanged, function (asyncResultII) {
    	    		if (asyncResultII.status === "failed") {
						app.showNotification('Error: ' + asyncResultII.error.message);
        			}
			});
		bindings.push(bindingString.trim());
        }
    });
		
		$('body').append(
			'<div id="notification-message">' +
				'<div class="padding">' +
					'<div id="notification-message-close"></div>' +
					'<div id="notification-message-header"></div>' +
					'<div id="notification-message-body"></div>' +
				'</div>' +
			'</div>');

		$('#notification-message-close').click(function () {
			$('#notification-message').hide();
		});


		// After initialization, expose a common notification function
		app.showNotification = function (header, text) {
			$('#notification-message-header').text(header);
			$('#notification-message-body').text(text);
			$('#notification-message').slideDown('fast');
		};
		app.closeNotification = function(){
			$('#notification-message').hide('fast');
		}
		
		//Improve storage system for better fault tolerance
		app.getDatasets = function(dataset){
			//Do we have the info in the document?
			var datasets = JSON.parse(Office.context.document.settings.get('datasets'));
			if(datasets && typeof(dataset) != 'undefined' && datasets.hasOwnProperty(dataset)){
				return datasets;
			}else{
				//If not, grab from backup
				var key = document_key_gen();
				if(key){
					var docs = JSON.parse(localStorage.getItem('exv-exceldocs-changes'));
					if(docs) return docs[key];
					return {};	
				}
				return {};
			}
		}
		
		app.setDatasets = function(datasets){
			//Update traditional document based storage
			Office.context.document.settings.set('datasets', JSON.stringify(datasets));
			//Update local backup
			var key = document_key_gen();
			if(key){
				if(localStorage.getItem('exv-exceldocs-changes') != null){
					var docs = JSON.parse(localStorage.getItem('exv-exceldocs-changes'));
				}else{
					var docs = {};
				}
				docs[key] = datasets;
				localStorage.setItem('exv-exceldocs-changes', JSON.stringify(docs));
				return true;
			}else{
				return false;
			}
		}
		
		app.getSchema = function(dataset){
			//Do we have the info in the document?
			var schema = JSON.parse(Office.context.document.settings.get('schema'));
			if(schema && typeof(dataset) != 'undefined' && schema.hasOwnProperty(dataset)){
				return schema;
			}else{
				//If not, grab from backup
				var key = document_key_gen();
				if(key){
					var docs = JSON.parse(localStorage.getItem('exv-exceldocs-schema'));
					if(docs) return docs[key];
					return {};	
				}
				return {};
			}
		}
		
		
		app.setSchema = function(schema){
			Office.context.document.settings.set('schema', JSON.stringify(schema));
			var key =  document_key_gen();
			if(key){
				if(localStorage.getItem('exv-exceldocs-schema') != null){
					var docs = JSON.parse(localStorage.getItem('exv-exceldocs-schema'));
				}else{
					var docs = {};
				}
				docs[key] = schema;
				localStorage.setItem('exv-exceldocs-schema', JSON.stringify(docs));
				return true;
			}else{
				return false;
			}
		}
		
		//app.logout = function(){
		$('document').ready(function(){
			$('#logout').click(function(){
				//Clear user data
				sessionStorage.removeItem("exv-token");
				sessionStorage.removeItem("exv-key");
				sessionStorage.removeItem("exv-private");
				//redirect back home
				window.location = "../Home/Home.html";
			});
		});
		//}
	};

	return app;
})();

function document_key_gen(){
	//Get document url
	var url = Office.context.document.url;
	//if url is null than document hasn't been saved yet
	if(url){
		//encode it so we can use it as a key
		var key = Base64.encode(url);
		return key
	}
	return false;
}

function advance(panel1, panel2){
	$(panel1).hide();
	$(panel2).removeClass('active');
	$(panel2).fadeIn();
	$(panel2).addClass('active');
}

function generate_hash(d){
	return Base64.encode(JSON.stringify(d));
}

function bindingChanged(eventArgs){
	eventArgs.binding.getDataAsync({ coerciontype: "table" }, function (asyncResult) {
        if (asyncResult.status === "failed") {
            app.showNotification('Error: ' + asyncResult.error.message);
        } else {
			if(typeof bind_init == 'undefined'){
				bind_init = true; 
			}
			if(!bind_init){
			dataset =  eventArgs.binding.id;
			//Check for access to this dataset
			//if(dataset_access_check(dataset, sessionStorage.getItem('exv-key'))){
            	
				//Is this a table or a matrix?
				if (typeof asyncResult.value.headers == 'undefined') {
					var headers = asyncResult.value[0];
					var rows = asyncResult.value.slice(1);
				}else{
					var headers = asyncResult.value.headers[0];
					var rows = asyncResult.value.rows;
				}
				look_for_change(dataset, headers, rows);
			}else{
				bind_init = false;
			}
        }
    });
}

function dataset_access_check(dataset, key){
	$.ajax({
				type: 'GET',
				url:'https://exversion.com/api/v1/metadata/'+dataset+'?key='+key,
				success: function(data){
					return data.body[0].access;
				}
});
}

function update_row_position(dataset, changes, row_id, increment){
	//Iterate through changes
	var len = changes.length;
	for(var c in changes){
		//if *_row exists and is greater than row_id or equal to (but not last value), increment
		if(c != (len-1) && typeof(changes[c]['data']) != 'undefined' && typeof(changes[c]['data']['*'+dataset+'_row']) != 'undefined' && changes[c]['data']['*'+dataset+'_row'] >= row_id){
			changes[c]['data']['*'+dataset+'_row'] = (changes[c]['data']['*'+dataset+'_row'] + increment);
		}
	}
	
	//Do it for any added_rows too
	for(var r in added_rows){
		if(added_rows[r] > row_id){
			added_rows[r] = added_rows[r] + increment;
		}		
	}
	return changes;
}

function last_row_deletes(hashes, values, i){
	//is the number of rows the same as it was before?
	if(hashes.length > values.length){
		//Are we on the last row?
		if((values.length-1) == i){
			return true;//position of where this is call will determine whether it's a last row or second to last row deletion (last row will not detect change at all)
		}
	}	
	return false;
}

function look_for_change(dataset, keys,values){
	//This is super annoying but Excel returns all data in the binding every time one thing changes
	datasets = app.getDatasets();
	hashes = datasets[dataset]
	var changes = sessionStorage.getItem('exv-changes');
	added_rows = sessionStorage.getItem('exv-added_rows') ? JSON.parse(sessionStorage.getItem('exv-added_rows')) : [];
	if(changes){
		changes = JSON.parse(changes);
	}else{
		changes = [];
	}
	//Check for changes in columns first
	schemas = app.getSchema();
	columns = schemas[dataset];
	var deleted_columns = $(columns).not(keys).get();
	var inserted_columns = $(keys).not(columns).get();
	if(deleted_columns.length > 0 || inserted_columns.length > 0){
		if(inserted_columns[0] != ""){
		//There's been a change in schema
		if(deleted_columns.length > 0){
			changes.push({'dataset':dataset, '_id':null,'type':'delete','column_name':deleted_columns[0]});	
		}else{
			changes.push({'dataset':dataset, '_id':null,'type':'insert','column_name':inserted_columns[0]});
		}
		//Reset schema
		schemas[dataset] = keys;
		app.setSchema(schemas);
		//Reset the hashes
		for(var i = 0, l = values.length; i < l; i++){
			var obj = {};
			var v = values[i]; 
			datasets[dataset][i]['hash'] = generate_hash(all_string(values[i]));
			}
		}else{
			return;
		}
	}else{
	var slide_down = false;
	var slide_up = false;
	var del_num = datasets[dataset].length - values.length; //Use for deleting multiple rows
	var splice_length = del_num;
	
	for(var i = 0, l = values.length; i < l; i++){
			var obj = {};
			var v = values[i]; 
			for (var j = 0, len = v.length; j < len; j++) {
				obj[keys[j]] = v[j];
			}
			if(typeof(datasets[dataset][i]) == 'undefined'){
				//If we add a row then the loop will overflow. Prevent that.
				temp_datasets[i] = {};
				temp_datasets[i]['hash'] = generate_hash(all_string(values[i]));
				temp_datasets[i]['_id'] = datasets[dataset][i-1]['_id'];
				continue;
			}
			
			if(last_row_deletes(datasets[dataset], values, i)){
				//Last row has been deleted
				for(var z = i; del_num > 0; z++, del_num--){//Modified for multiple last rows
					if(!slide_up){//Separate out because we still need to snip the last hash off
						if($.inArray(z, added_rows) != -1){
							added_rows = $(added_rows).not([z]).get();
							changes = remove_from_changes(dataset, changes, z);	
						}else{
							//Add delete
							changes.push({'dataset':dataset, '_id':datasets[dataset][z+1]['_id'], 'delete': true});
						}
						//Update any other changes that depend on row position
						changes = update_row_position(dataset, changes, z, -1);
					}
					datasets[dataset][z]['hash'] = generate_hash(all_string(values[z]));
					datasets[dataset][z]['_id'] = datasets[dataset][z+1]['_id'];
				}
				//Update hash
				datasets[dataset].splice(i+1, splice_length);
				break;
			}
			if(slide_down){
					temp_datasets[i] = {};
					temp_datasets[i]['hash'] = generate_hash(all_string(values[i]));
					temp_datasets[i]['_id'] = datasets[dataset][i-1]['_id'];
				}
				
			if(slide_up){
				datasets[dataset][i]['hash'] = generate_hash(all_string(values[i]));
				datasets[dataset][i]['_id'] = datasets[dataset][i+1]['_id'];
			}
		
			
			if(datasets[dataset][i]['hash'] != generate_hash(all_string(values[i])) && typeof(temp_datasets) == 'undefined'){
				//If _id equals null than this is a new row we're added data to
				if(!datasets[dataset][i]['_id'] && del_num == 0){
					//Go through our changes and find this row
					for(var r in changes){
						if(typeof(changes[r]['data']) != 'undefined' && typeof(changes[r]['data']['*'+dataset+'_row']) != 'undefined' && changes[r]['data']['*'+dataset+'_row'] == i){
							obj['*'+dataset+'_row'] = i;
							changes[r]['data'] = obj;
							
							//Update Hash
							datasets[dataset][i]['hash'] = generate_hash(all_string(values[i]))
							break;
						}
					}
					break;
				}
				
				//Found the change!
				if(del_num < 0){
					//If this is a new row, update all other rows
					slide_down = true;
					obj['*'+dataset+'_row'] = i;
					changes.push({'dataset':dataset, '_id':null, 'data': obj});
					//Update any other changes that depend on row position
					changes = update_row_position(dataset, changes, i, 1);
					added_rows.push(i);
					//Update hash
					var temp_datasets = datasets[dataset].slice(0,i);
					temp_datasets[i] = {};
					temp_datasets[i]['hash']=generate_hash(all_string(values[i]));
					temp_datasets[i]['_id']=null;
					continue;
				}
				
				//If this is a deleted row
				/* Is breaking when deleted row is the last or second to last row. What's the best way to identify this case? */
				else if(del_num > 0){
					if(!slide_up){//Don't delete all the rows after (duh)
						if(del_num == 1){
							slide_up = true; //If we have multiple rows deleted wait to turn this on
						}else{
							del_num = del_num - 1;
						}
						
						//If this is an added row we're deleting just remove it from changes instead
						if($.inArray(i, added_rows) != -1){
							added_rows = $(added_rows).not([i]).get();
							changes = remove_from_changes(dataset, changes, i);	
						}else{
							//Add delete
							changes.push({'dataset':dataset, '_id':datasets[dataset][i]['_id'], 'delete': true});
						}
						//Update any other changes that depend on row position
						changes = update_row_position(dataset, changes, i, -1);
						//Update hash
						datasets[dataset][i]['hash'] = generate_hash(all_string(values[i]));
						datasets[dataset][i]['_id'] = datasets[dataset][i+1]['_id'];
						continue;
					}
				}else{
				//Check to see if we've updated this row already so we cut back on duplicates
				var existing_changes = check_existing_changes(changes, datasets[dataset][i]['_id'], obj);
				if(!existing_changes){
					changes.push({'dataset':dataset, '_id':datasets[dataset][i]['_id'], 'changes':obj});
				}else{
					changes = existing_changes;
				}	
				//Update hash
				datasets[dataset][i]['hash'] = generate_hash(all_string(values[i]));
				break;
				}
								
			}
			
		}
	}
		if(temp_datasets){
			datasets[dataset] = temp_datasets;
		}
		app.setDatasets(datasets);
		sessionStorage.setItem('exv-changes', JSON.stringify(changes));
		sessionStorage.setItem('exv-added_rows',JSON.stringify(added_rows));
		window.location = '../Sync/Sync.html';
}

function remove_from_changes(dataset, changes, row_id){
	var new_changes = [];
	for(var c in changes){
		if(!changes[c].hasOwnProperty('data') || !changes[c]['data'].hasOwnProperty('*'+dataset+'_row') || changes[c]['data']['*'+dataset+'_row'] != row_id){
			new_changes.push(changes[c]);
		}
	}
	return new_changes;
}

function check_existing_changes(changes, id, obj){
	for(var c in changes){
		if(changes[c]['_id'] == id){
			changes[c]['changes'] = obj
			return changes
		}
	}
	return false;
}

function readBoundData(dataset, callback) {
    Office.select("bindings#"+dataset).getDataAsync({ coercionType: "table" }, 
        function (asyncResult) {
            if (asyncResult.status === "failed") {
                console.log('Error: ' + asyncResult.error.message);
            } else {
                return callback(asyncResult.value.headers[0], asyncResult.value.rows, send_data);
            }
        });
}

function readBoundDataSend(dataset){
	bind_init = true;
	Office.select("bindings#"+dataset).addHandlerAsync("bindingDataChanged", bindingChanged, function (asyncResult) {
    if (asyncResult.status === "failed") {
		app.showNotification('Error: ' + asyncResult.error.message);
     }else{
        return true;
     }
});
}

function isEmpty(obj){
    for(p in obj){
        if(obj[p] != ""){
            return false;
        }
    }
	return true;
}

function isJson(obj){
	try{ var json = JSON.parse(obj); return true;}
	catch(e){ return false;}
}

function isString(obj){ if (obj.substring) { return true;}else{return false;}}

function all_string(obj){
	//Force everything to render as a string to prevent hashes from throwing false positives
	var data = [];
	for(var i in obj){
		if(isJson(obj[i])){
			data.push(JSON.stringify(obj[i]));
		}else{
			data.push(obj[i].toString());
		}
	}
	return data;
}
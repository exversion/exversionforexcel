/// <reference path="../App.js" />
/*global app*/
var dataset = '';
var columns = [];
var column_count = 0;
var row_count = 0;
var query = '';
var column_alpha = ['','A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'];
var remove_columns = [];
var no_bind = false;
var dataset_access = false;
var hashes = [];
var row_order = [];

(function () {
	'use strict';

	// The initialize function must be run each time a new page is loaded
	Office.initialize = function (reason) {
		$(document).ready(function () {
			app.initialize();
			//Grab datasets from Exversion
			var datasets_access_list = JSON.parse(sessionStorage.getItem("exv-access_list"));
			$('#dataset_dropdown').html(populate_dropdown(datasets_access_list));
		
		$('#dataset_dropdown_button').click(function(){
			dataset = $('#dataset_dropdown').val();
			data_import();
		});
		
		$('#search_submit').click(function(){
			var button = $(this).html();
			$(this).html('<img src="../../images/ajax-loader.gif">');
				
			//Grab input value
			var q = $('#dataset_search').val();
			var list = '';
			//run search
			$.ajax({
				type: 'GET',
				url:'https://exversion.com/api/v1/search?_public=1&q='+q,
				success: function(data){
					$('#search_submit').html(button);
					for(var d in data.body){
						list += '<div class="search_item"><strong>'+data.body[d].name+'</strong><p>'+data.body[d].description+'</p><center><a href="https://exversion.com/data/view/'+data.body[d].dataset+'" target="new" class="btn btn-warning">View</a> <a class="btn btn-danger import" id="'+data.body[d].dataset+'">Import</a></center></div>';
					}
					$('#search_results').html(list);
				}
			});
		});
		
		$('#search_results').on('click','.import', function(){
			dataset = this.id;
			data_import();
		});
		
		$('#filters_inner').on('click','button', function(){
			remove_columns.push(this.id);
			//Remove column from column array
			columns = without(columns, this.id);
			column_count -= 1;
			$(this).parent().remove();
		});
		
		$('#filter_submit').click(function(){
			var button = $(this).html();
			$(this).html('<img src="../../images/ajax-loader.gif">');
			
			if($('#create_branch').is(':checked')){
				//Create a branch and change dataset value before going any further
				var this_private = 0;
				if($('#create_private').is(':checked')){
					this_private = 1;
				}
				$.ajax({
				type: 'POST',
				url:'https://exversion.com/api/v1/dataset/fork/',
				data: {"access_token":sessionStorage.getItem('exv-token'),'parent':dataset,'name':"",'forkchanges':"",'track':0,'private':this_private, 'clear':0},
				success: function(data){
					$(this).html(button);
					dataset = data.body[0].dataset;
				},
				error: function(){
					$(this).html(button);
					app.showNotification('Error: Unable to create branch');
				}
			});
			}else if(dataset_access){
				no_bind = false;
			}else{
				no_bind = true;
			}
			query = '?key='+key;
			var form = $('#filter .filter_value').each(function(i){
				if($(this).val()){
					query += '&'+this.id+'='+$(this).val();
				}
			});
			query += '&_remove='+remove_columns.join(',');
			
			//Get count
			$.ajax({
				type: 'GET',
				url:'https://exversion.com/api/v1/count/'+dataset+query,
				success: function(data){
					$(this).html(button);
					$('#filtered_row_count').html(data.body[0].count);
					row_count = data.body[0].count+1;//Plus one to account for header	
				}
			});
			
			//Get sample
			$.ajax({
				type: 'GET',
				url:'https://exversion.com/api/v1/dataset/'+dataset+query+'&_limit=2',
				success: function(data){
					$(this).html(button);
					$('#preview').html(JSON.stringify(data.body, null, 4));
					advance('#filter','#confirm');	
				}
			});
		});
		
		$('#confirm_submit').click(function(){
			$('#select_space_column_count').html(column_count);
			$('#select_space_row_count').html(row_count);
			$('#column_alpha').html(column_alpha_calc(column_count));
			advance('#confirm', '#select_space');
		});
		
		$('#filter_redo').click(function(){
			$('#preview').html('');
			advance('#filter','#importData');
		});
		
		$('#confirm_redo').click(function(){
			$('#preview').html('');
			advance('#confirm','#filter');
		});
		
		$('#select_space_submit').click(function(){
			//If the binding exists already, clear it
			Office.context.document.bindings.releaseByIdAsync(dataset, function (asyncResult) {});
			//Create binding and check proper space
			Office.context.document.bindings.addFromSelectionAsync("matrix", { id: dataset }, function (asyncResult) {
            if (asyncResult.status === "failed") {
				app.showNotification('Error:', asyncResult.error.message);
            }else{
				//Add event handlers
				if(!no_bind){
					writeBoundData(dataset,asyncResult);
				}else{
					checkSize(asyncResult.value.columnCount, asyncResult.value.rowCount, grabData);
				}
			}
        });
		});
		
		$('#select_space_redo').click(function(){
			advance('#select_space','#filter');
		});
		
		});
	}
})();

function data_import(){
	$.ajax({
				type: 'GET',
				url:'https://exversion.com/api/v1/metadata/'+dataset+'?key='+key,
				success: function(data){
					//$('#preview').html(JSON.stringify(data.body, null, 4));
					
					advance('#importData','#filter');
					//Create filters
					var filters = '';
					dataset_access = data.body[0].access;
					if(!dataset_access){
						$('#branch').show();
						
						//Private repos
						var repo = sessionStorage.getItem("exv-private");
						if(repo == 'unlimited'){
							$('#free_account').remove();
							$('private').show();
						}
						else if(repo == 0){
							$('#free_account').show();
							$('#private').hide();
						}else{
						if(repo == 1){
							$('#repo_plural').html('repo');
						}
						$('#private_repo_available').html(repo);
						$('#free_account').show();
						$('private').show();
					}
					}else{
						$('#branch').hide();
						$('#private').hide();
					}
					for(var f in data.body[0].columns){
						filters += '<div><input type="text" disabled class="form-control" value="'+data.body[0].columns[f]+' = " style="width:30%; float:left;"/><input type="text" id="'+data.body[0].columns[f]+'" class="filter_value form-control" style="width:54%; float:left; margin-left:2px;"/><button id="'+data.body[0].columns[f]+'" class="btn btn-danger remove_col" style="width:15%;float:right;"><i class="fa fa-times fa-lg"></i></button><br style="clear:both;"></div>';
					}
					$('#filters_inner').html(filters);
					$('#row_count').html(data.body[0].rows);
					columns = data.body[0].columns;
					column_count = columns.length;
				}
			});
}

function without(array, item){
	var array_without = [];
	for(i in array){
		if(array[i] != item){
			array_without.push(array[i]);
		}
	}
	return array_without;
}

function populate_dropdown(datasets){
	var options = '';
	for(var d in datasets){
		options += '<option value="'+datasets[d]['dataset']+'">'+datasets[d]['name']+'</option>';
	}
	return options;
}

function column_alpha_calc(col){
	var col_advance = 1;
	if(col > 26){
		var col_advance = 0;
		while(col > 26){
			col_advance += 1;
			var col = col - 26 
		}
		return column_alpha[col_advance]+column_alpha[col];
	}else{
		return column_alpha[col];
	}
	
}

function checkSize(col, row, callback){
	if(col < column_count || row < row_count){
			app.showNotification('Selection is not big enough for all the data. Needs to be '+column_count+' columns by '+row_count+' rows');
	}else if(col > column_count || row > row_count){
			app.showNotification("Selection is too big for the data and the people who built Microsoft's API are morons. Needs to be "+column_count+" columns by "+row_count+" rows");
	}else{
			grabData(dataset, writeData);
	}
}

function writeBoundData(dataset, asyncResult){
	Office.select("bindings#"+dataset).addHandlerAsync("bindingDataChanged", bindingChanged, function (asyncResultII) {
   	if (asyncResultII.status === "failed") {
		app.showNotification('Error: ' + asyncResultII.error.message);
    } else {
    	checkSize(asyncResult.value.columnCount, asyncResult.value.rowCount, grabData);
    }
 });
}

function grabData(dataset, callback){
	var fetch = true;
	var i = 1;
		$.ajax({
				type: 'GET',
				url:'https://exversion.com/api/v1/dataset/'+dataset+query+'&_limit=10000&_page='+i,
				success: function(data){
					if(data.status == 204){
						fetch = false;
					}else{
						callback(arrangeRows(data.body, i));		
					}
				}
			});
}

function arrangeRows(data, i){
	var rows = [columns];
	var keys = [];
	for(var d in data){
		var row_data = [];
		row_order.push(data[d]['_id']);
		for(var c in columns){
			var k = columns[c];
				if(isString(data[d][k])){
					var data_str = data[d][k];
				}else if(isJson(data[d][k])){
					var data_str = JSON.stringify(data[d][k]);
				}else{
					//Force everything to evaluate as a string
					var data_str = data[d][k].toString();
				}
			row_data.push(data_str);
		}
		rows.push(row_data);
	}
	return rows;
}

function isNumber(obj) { return !isNaN(parseFloat(obj)); }

function isNumberMoney(num){
    if(!isNaN(parseFloat(num.replace(/[^0-9-.,]/g, ''))) && num.match(/[^0-9-.,]/g) && num.match(/[^0-9-.,]/g).length < 5){
       return true;
       }else{
       return false;
       }
}


function writeData(data){
	//Store columns for later
		if(Office.context.document.settings.get('schemas')){
			schemas = app.getSchema(dataset);
			schemas[dataset] = columns;
		}else{
			schemas = {};
			schemas[dataset] = columns;
		}
		app.setSchema(schemas);
				
	 Office.context.document.setSelectedDataAsync(data, function (asyncResult) {
        if (asyncResult.status === "failed") {
            app.showNotification('Error: ' + asyncResult.error.message);
        }else{
			//Generate hashes after insert as a work around for Excel's autoformatting
			set_hashes(dataset);
		}
    });
}

function set_hashes(dataset){
	hashes = []; 
	Office.select("bindings#"+dataset).getDataAsync({ coercionType: "matrix" }, 
        function (asyncResult) {
            if (asyncResult.status === "failed") {
                console.log('Error: ' + asyncResult.error.message);
            } else {
		var data = asyncResult.value.slice(1);
        for(var d in data){
			var hash = generate_hash(all_string(data[d]))
			hashes.push({"_id":row_order[d], "hash":hash});
            }
			datasets = app.getDatasets(dataset);
			datasets[dataset] = hashes;
			app.setDatasets(datasets);
			advance('#select_space','#complete');
        }
	});	
}
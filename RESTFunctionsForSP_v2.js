/* Collection of REST functions to be used with SharePoint. Reference up to date jQuery in addition to this file. */


/****************** WORKING WITH LISTS AND LIST ITEMS ******************/

/* Gets SharePoint list items and returns an array of objects containing list data for all returned items. 
Must specify listName obj value. Additional obj values are optional to retrieve lookup values and add a filter to the query. 
By default, all non lookup fields are selected. */
function getSharePointListItemData(obj) {
	
	if (!obj.siteUrl) {
		obj.siteUrl = _spPageContextInfo.webAbsoluteUrl;
	}
	
	//Build the query url from parameters here
	var url = obj.siteUrl + "/_api/web/lists/getbytitle('" + obj.listName + "')/items?$select=*";
	if (obj.selectLookupStr !== undefined && obj.selectLookupStr !== "") {
		url = url + "," + obj.selectLookupStr;
	}
	if (obj.expandLookupStr !== undefined && obj.expandLookupStr !== "") {
		url = url + "&$expand=" + obj.expandLookupStr;
	}
	if (obj.filterStr !== undefined && obj.filterStr !== "") {
		url = url + "&$filter=" + obj.filterStr;
	}
	if (obj.orderBy !== undefined && obj.orderBy !== "") {
		url = url + "&$orderBy=" + obj.filterStr;
	}
		
	//Perform query
	$.ajax({
		url: url,				
		method: "GET",
		headers: { "Accept": "application/json; odata=verbose" },
		success: obj.success,
		error: obj.error,
	});
}

/* Gets SharePoint list items based on a SharePiont view. Must specify webUrl (usually current site), listName, and viewTitle parameters. 
Also uses getJson and getListItems functions to get list items from the view */
function getListItemsForView(obj) {
	var viewQueryUrl = obj.siteUrl + "/_api/web/lists/getByTitle('" + obj.listName + "')/Views/getbytitle('" + obj.viewTitle + "')/ViewQuery";
	return getJson(viewQueryUrl).then(
	function(data){         
		var viewQuery = data.d.ViewQuery;
		return getListItems(obj.siteUrl, obj.listName, viewQuery); 
	});
}

/* Get List Items for View - Gets JSON of the view to retrieve ViewQuery (CAML) */	
function getJson(url) {
	return $.ajax({
		url: url,
		type: "GET",
		contentType: "application/json;odata=verbose",
		headers: { 
			"Accept": "application/json;odata=verbose"
		}
	});
}

/* Get List Items for View - Performs query to the specified list using retrieved CAML */
function getListItems(webUrl, listTitle, queryText) {
	var viewXml = '<View><Query>' + queryText + '</Query></View>';
	var url = webUrl + "/_api/web/lists/getbytitle('" + listTitle + "')/getitems?$select=*,FieldValuesAsText/Title&$expand=FieldValuesAsText&@target='" + webUrl + "'";
	var queryPayload = {  
	   'query' : {
			'__metadata': { 'type': 'SP.CamlQuery' }, 
			'ViewXml' : viewXml  
	   }
	};

	return $.ajax({
		url: url,
		method: "POST",
		data: JSON.stringify(queryPayload),
		headers: {
			"X-RequestDigest": $("#__REQUESTDIGEST").val(),
			"Accept": "application/json; odata=verbose",
			"content-type": "application/json; odata=verbose"
		}
	});
}

/* Updates a list item using the listName, itemId, and itemProperties 
Also uses getItemTypeForListName to build itemProperties for update. Must build out success and failure functions when calling. */
function updateListItem(obj) {
	var itemType = getItemTypeForListName(obj.listName);
	obj.itemProperties["__metadata"] = { "type": itemType };

	$.ajax({
			url: _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getbytitle('" + obj.listName + "')/items(" + obj.itemId + ")",
			type: "POST",
			contentType: "application/json;odata=verbose",
			data: JSON.stringify(obj.itemProperties),
			headers: {
					"Accept": "application/json;odata=verbose",
					"X-RequestDigest": $("#__REQUESTDIGEST").val(),
					"IF-MATCH": "*",
					"X-HTTP-Method": "MERGE",
			},
			success: obj.success,
			error: obj.error
		});
}

/* Update List Item  - Gets List Item Type metadata */
function getItemTypeForListName(name) {
		return "SP.Data." + name.charAt(0).toUpperCase() + name.split(" ").join("").slice(1) + "ListItem";
}


/****************** WORKING WITH LIBRARIES, DOCUMENTS, AND FOLDERS/DOC SETS ******************/

/* Gets SharePoint library items from a specific folder and returns an array of objects containing file data for all returned items.
Must specify a folder url. Additional parameter to retrieve lookup values. By default, all non lookup fields are selected. */
function getSharePointFolderFileData(obj) {
	
	//Build the query url from parameters here
	var url = siteUrl + "/_api/web/GetFolderByServerRelativeUrl('" + obj.folderUrl + "')/Files?$select=*";
	if (obj.selectLookupStr !== undefined && obj.selectLookupStr !== "") {
	url = url + "," + obj.selectLookupStr;
	}
	if (obj.expandLookupStr !== undefined && obj.expandLookupStr !== "") {
		url = url + "&$expand=" + obj.expandLookupStr;
	}
		
	//Perform query
	$.ajax({  
		url: url,
		method: "GET",  
		headers: { "Accept": "application/json; odata=verbose" },  
		success: obj.success,  
		error: obj.error
	});
}


/* Creates a document set within a library. Must specify a webUrl (usually current site, but can be another site), a list name, a document set name, and the success/error functions. 
Also uses createFolder and getListUrl functions to fully create document set */
function createSharePointDocumentSet(obj) {
	createFolder(obj.siteUrl, obj.listName, obj.folderName,'0x0120D520', obj.success, obj.error);
} 
	
/* Document Set Creation - Creates folder in library with document set content type */
function createFolder(webUrl, listName, folderName, folderContentTypeId, success, error) {  
	getListUrl(webUrl, listName,
	function(listUrl) {
		var folderPayload = {
			'Title' : folderName,
			'Path' : listUrl
		};

		//Create Folder resource
		$.ajax({
			url: webUrl + "/_vti_bin/listdata.svc/" + listName,
			type: "POST",
			contentType: "application/json;odata=verbose",
			data: JSON.stringify(folderPayload),
			headers: {
			   "Accept": "application/json;odata=verbose",
			   "Slug": listUrl + "/" + folderName + "|" + folderContentTypeId
			},
			success: success,
			error: error
		});
	},
	error);
}

/* Document Set Creation  - Gets list url to use as location for document set creation */
function getListUrl(webUrl, listName, success, error) {
		var headers = {};
		$.ajax({       
		   url: webUrl + "/_api/lists/getbytitle('" + listName + "')/rootFolder?$select=ServerRelativeUrl",   
		   type: "GET",   
		   contentType: "application/json;odata=verbose",
		   headers: { 
			  "Accept": "application/json;odata=verbose"
		   },   
		   success: function(data){
			   success(data.d.ServerRelativeUrl);
		   },
		   error: error
		});
	}
	
	
/* Updates a document's data/properties. Must specify a webUrl (usually current site), a file url, and the item properties (JSON object) that you want to update.
Also uses executeJson function to update document properties. */
function updateFileProperties(obj) {
	var endpointUrl = obj.siteUrl + "/_api/web/getFileByServerRelativeUrl('" + obj.fileUrl + "')/ListItemAllFields";
	return executeJson(endpointUrl).then(function(data){
		var updateHeaders = {
			'X-HTTP-Method' : 'MERGE',
			'If-Match': data.d['__metadata']['etag']
		};
		var itemPayload = itemProperties;
		itemPayload['__metadata'] = {'type': data.d['__metadata']['type']};
		return executeJson(endpointUrl, 'POST', updateHeaders, itemPayload);    
	});
}

/* Update File Properties - Performs file property update using passed arguements */
function executeJson(url, method, headers, payload) {
	method = method || 'GET';
	headers = headers || {};
	headers["Accept"] = "application/json;odata=verbose";
	if(method == "POST") {
		headers["X-RequestDigest"] = $("#__REQUESTDIGEST").val();
	}      
	var ajaxOptions = 
	{       
	   url: url,   
	   type: method,  
	   contentType: "application/json;odata=verbose",
	   headers: headers
	};
	if (typeof payload != 'undefined') {
	  ajaxOptions.data = JSON.stringify(payload);
	}  
	return $.ajax(ajaxOptions);
}


/* Uploads a document to a library or folder using a url parameter and an html file input and upload button.
Also uses onError to alert error messages during file upload */
function upload(obj) {
	// Define the folder path
	var serverRelativeUrlToFolder = obj.destinationUrl;

	// Get test values from the file input and text input page controls.
	var fileInput = jQuery('#getFile');
	var newName = jQuery('#displayName').val();
	var fileCount = fileInput[0].files.length;
	if (fileCount > 0) {
		$('.loader').show();
	}
	// Get the server URL.
	var serverUrl = _spPageContextInfo.webAbsoluteUrl;
	var filesUploaded = 0;
	for(var i = 0; i < fileCount; i++){
		// Initiate method calls using jQuery promises.
		// Get the local file as an array buffer.
		var getFile = getFileBuffer(i);
		getFile.done(function (arrayBuffer,i) {

			// Add the file to the SharePoint folder.
			var addFile = addFileToFolder(arrayBuffer,i);
			addFile.done(function (file, status, xhr) {
				//$("#msg").append("<div>File : "+file.d.Name+" ... uploaded sucessfully</div>");
				filesUploaded++;
				if(fileCount == filesUploaded){
					alert("All files uploaded successfully");
					//$("#msg").append("<div>All files uploaded successfully</div>");
					$("#getFile").value = null;
					filesUploaded = 0;
				}
			});
			addFile.fail(onError);
		});
		getFile.fail(onError);

	}

	// Get the local file as an array buffer.
	function getFileBuffer(i) {
		var deferred = jQuery.Deferred();
		var reader = new FileReader();
		reader.onloadend = function (e) {
			deferred.resolve(e.target.result,i);
		}
		reader.onerror = function (e) {
			deferred.reject(e.target.error);
		}
		reader.readAsArrayBuffer(fileInput[0].files[i]);
		return deferred.promise();
	}

	// Add the file to the file collection in the Shared Documents folder.
	function addFileToFolder(arrayBuffer,i) {
		var index = i;

		// Get the file name from the file input control on the page.
		var fileName = fileInput[0].files[index].name;

		// Construct the endpoint.
		var fileCollectionEndpoint = String.format(
				"{0}/_api/web/getfolderbyserverrelativeurl('{1}')/files" +
				"/add(overwrite=true, url='{2}')",
				serverUrl, serverRelativeUrlToFolder, fileName);

		// Send the request and return the response.
		// This call returns the SharePoint file.
		return jQuery.ajax({
			url: fileCollectionEndpoint,
			type: "POST",
			data: arrayBuffer,
			processData: false,
			headers: {
				"accept": "application/json;odata=verbose",
				"X-RequestDigest": jQuery("#__REQUESTDIGEST").val(),
				"content-length": arrayBuffer.byteLength
			}
		});
	}
}

/* Display error messages. */
function onError(error) {
	alert(error.responseText);
}


/* Copies a document to a document set or folder. Must specify the webUrl (usually current site, but can be another site), a url for the document to be copied, and the destination url. */
function copy(obj) {	
	var url = obj.siteUrl + "/_api/web/GetFileByServerRelativeUrl('" + obj.docToCopyUrl + "')/copyTo(strNewUrl = '" + obj.destinationUrl + "')";
	$.ajax({
		type: "POST",
		contentType: "application/json;odata=verbose",
		url: encodeURI(url),
		headers: {
			"Accept": "application/json;odata=verbose",
			"X-RequestDigest": jQuery("#__REQUESTDIGEST").val()
		},
		success: obj.success,
		error: obj.error,
	});
}


/* Deletes a file using its url after a user confirms from prompt. Can be specified on onlick javascript of a button.*/
function deleteFile(fileUrl) {
	var r = confirm("Are you sure you want to delete this item?");
	if (r == true) {
		$.ajax({  
			url: _spPageContextInfo.webAbsoluteUrl + "/_api/web/GetFileByServerRelativeUrl('" + fileUrl + "')/recycle()",
			method: "POST",  
			headers: { 
				"Accept": "application/json; odata=verbose" ,
				"Content-Type": "application/json; odata=verbose" ,
				"X-HTTP-Method": "DELETE",
				"If-Match": "*",
				"X-RequestDigest": $("#__REQUESTDIGEST").val()
			},
			success: obj.success,
			error: obj.error,
		});	
	}
}


/****************** WORKING WITH USERS AND PERMISSIONS ******************/

/* Gets current user data and returns it in an object. */
function getCurrentUser() {  
	$.ajax({  
		url: _spPageContextInfo.webAbsoluteUrl + "/_api/web/CurrentUser",  
		method: "GET",  
		headers: { "Accept": "application/json; odata=verbose" },  
		success: success,  
		error: error
	});
}  

/* Gets a user's group affiliation from the current site using a userId parameter and returns group titles in an array. */
function getCurrentUserGroupColl(userID , callback) {  
	$.ajax({  
		url: _spPageContextInfo.webAbsoluteUrl + "/_api/web/GetUserById(" + userID + ")/Groups",  
		method: "GET",  
		headers: { "Accept": "application/json; odata=verbose" },  
		success: callback,
		error: function (data) {  
			failure(data);  
		}  
	});
} 
 
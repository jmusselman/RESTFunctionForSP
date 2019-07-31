# Read Me - REST Functions for SharePoint

This file outlines usage of functions within **RESTFunctionsForSP.js**. The functions listed here that retrieve data will return it as a JavaScript Object. For info on these, [click here](https://www.w3schools.com/js/js_json_objects.asp).

For more detailed info on REST API, refer to the below links: 

- [Complete Basic operations using SharePoint REST endpoints](https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/complete-basic-operations-using-sharepoint-rest-endpoints)
- [Using REST API for Selecting, Filtering, Sorting, and Pagination](https://social.technet.microsoft.com/wiki/contents/articles/35796.sharepoint-2013-using-rest-api-for-selecting-filtering-sorting-and-pagination-in-sharepoint-list.aspx)

The functions for the following actions will be outlined:

**Working with Lists**

- Get SharePoint List Item Data
- Get SharePoint List Item Data Using Existing SharePoint View
- Update SharePoint List Items

**Working with Libraries and Files**

- Get SharePoint Folder File Data / Properties
- Create SharePoint Document Set
- Update File Data / Properties
- Upload a Document Using a File Input
- Copy Library Document to Another Library
- Delete a File

**Working with Users and Permissions**

- Get Current User
- Get Current User Group Collection
- Get Current User Profile Picture


## Get SharePoint List Item Data

Call the function `getSharePointListItemData` to retrieve SharePoint list item data from the current site. By default, this function will retrieve all non Lookup and People fields for an item. Lookup and People fields must be specified in the query. You may do this by passing additional values in the obj argument to the function. You may also filter your query. 

### Object Parameters 

- **listName (Required)** - A required string value that is the display name (not the internal name) of the list you are querying.
- **success (Required)** - A required function to specify what to do on successful query.
- **error (Required)** - A required function to specify what to do on unsuccessful query.
- **selectLookupStr** and **expandLookupStr** - String values used to select and expand desired Lookup and People fields.
- **filterStr** - A string value used to apply a filter to your query. 

### Usage

    getSharePointListItemData({
		listName: 'My SharePoint List',
		success: function (data) {  
			var items = [];  
			$.each(data.d.results, function(i,result) {
				items.push(result);
			});	
			console.log(items);
			return items;
		},
		error: function (data) { 
			alert('Error');      
		}
	});

In this example, no lookup fields are selected. The query will by default return all fields (that aren't lookups) and all list items.

	getSharePointListItemData({
		listName: 'My SharePoint List',
		selectLookupStr: 'LookupField/Title,PeopleField/Id', //Selects the Title of the LookupField and Id of the PeopleField
		expandLookupStr: 'LookupField,PeopleField', //Need to expand the fields to access the selected values
		success: function (data) {  
			var items = [];  
			$.each(data.d.results, function(i,result) {
				items.push(result);
			});	
			console.log(items);
			return items;
		},
		error: function (data) { 
			alert('Error');      
		}
	});

In this example, a Lookup field and a People field have been selected and expanded in addition to the rest of the list item data.

	getSharePointListItemData({
		listName: 'My SharePoint List',
		selectLookupStr: 'LookupField/Title,PeopleField/Id', //Selects the Title of the LookupField and Id of the PeopleField
		expandLookupStr: 'LookupField,PeopleField', //Need to expand the fields to access the selected values
		filterStr: "Status eq 'In Progress'", //Filters list item data where the Status field is equal to 'In Progress' 
		success: function (data) {  
			var items = [];  
			$.each(data.d.results, function(i,result) {
				items.push(result);
			});	
			console.log(items);
			return items;
		},
		error: function (data) { 
			alert('Error');      
		}
	});

In this example, a Lookup field and a People field have been selected and expanded in addition to the rest of the list item data. A filter has also been applied to the query that only retrieves items that have a Status of 'In Progress'.

	getSharePointListItemData({
		listName: 'My SharePoint List',
		filterStr: "Status eq 'In Progress'", //Filters list item data where the Status field is equal to 'In Progress' 
		success: function (data) {  
			var items = [];  
			$.each(data.d.results, function(i,result) {
				items.push(result);
			});	
			console.log(items);
			return items;
		},
		error: function (data) { 
			alert('Error');      
		}
	});

In this example, a filter has been applied to the query that only retrieves items that have a Status of 'In Progress'.


## Get SharePoint List Item Data Using Existing SharePoint View

Call the function `getListItemsForView` to retrieve SharePoint list item data from the current site based on an existing SharePoint List View. By default, this function will retrieve all fields values as text of an item. The functions `getJson` and `getListItems` are also used to complete this process.

### Object Parameters 

- **siteUrl (Required)** - A required string value that is the site url that contains the list you want to query. 
- **listTitle (Required)** - A required string value that is the display name (not the internal name) of the list you are querying. 
- **viewTitle (Required)** - A required string value used to specify the SharePoint View you want to apply to the query. This will be the title of the view. 

### Usage

	var listData = getListItemsForView({
		siteUrl: _spPageContextInfo.webAbsoluteUrl, //SharePoint object that returns current site url
		listName: 'My SharePoint List',
		viewTitle: 'All Active Items',
	}).done(function(data){
		var items = data.d.results;
	}).fail(function(error){
		console.log(error)
	});


## Update SharePoint List Items

Call the function `updateListItem` to update a SharePoint list item in the current site. You must specify your own success and failure functions. The function `getItemTypeForListName` is used to complete this process.

### Parameters

- **listName (Required)** - A required string value that is the display name (not the internal name) of the list you are querying. 
- **itemId (Required)** - A required integer value that is the ID of the item you want to update.
- **itemProperties (Required)** - a required object that contains the data you want to update. Object keys must be the same as internal SharePoint field names to correctly update.
- **success (Required)** - a required callback function that is called on successful query.
- **error (Required)** - a required callback function that is call on unsuccessful query.

### Usage

	updateListItem({
		listName: "My SharePoint List",
		itemId: 12,
		itemProperties: {
			'Status': 'Complete',
			'Concert_x0020_Tickets': 20
		} //Object keys must match sharePoint field internal names to update correctly
		success: function() {
			alert('Your item has been saved!');
		},
		error: function(error) {
			alert(JSON.stringify(error));
		}
	});


## Get SharePoint Folder File Data / Properties

Call the function` getSharePointFolderFileData` to retrieve library data from a specific folder. By default, this function will retrieve all non Lookup and People fields for an item. Lookup and People fields must be specified in the query. You may do this by passing arguments to the function.

### Parameters 

- **folderUrl (Required)** - A required string value that is the url of the folder you are querying. 
- **selectLookupStr** and **expandLookupStr** - String values used to select and expand desired Lookup and People fields.
- **success (Required)** - a required callback function that is called on successful query.
- **error (Required)** - a required callback function that is call on unsuccessful query.

### Usage

	getSharePointFolderFileData({
		folderUrl: 'sites/MySharePointSite/MyLibrary/MyFolder',
		success: function(data) { 
			var items = []; //Array that will contain data objects
			$.each(data.d.results, function(i,result) {
				items.push(result);
			});	
			items = items.sort(function(a, b) {
				if(a.Title < b.Title) {return -1;}
				if(a.Title > b.Title) {return 1;}
				return 0;
			});
			console.log(items);
			return items;
		},
		error: function(data) {  
			alert("Error");  
		}  
	});

In this example, the remaining parameters can be left blank. The query will return all non Lookup and People fields for items within the folder.

	getSharePointFolderFileData({
		folderUrl: 'sites/MySharePointSite/MyLibrary/MyFolder',
		selectLookupStr: 'LookupField/Title,PeopleField/Id', //Selects the Title of the LookupField and Id of the PeopleField
		expandLookupStr: 'LookupField,PeopleField'; //Need to expand the fields to access the selected values
		success: function(data) { 
			var items = []; //Array that will contain data objects
			$.each(data.d.results, function(i,result) {
				items.push(result);
			});	
			items = items.sort(function(a, b) {
				if(a.Title < b.Title) {return -1;}
				if(a.Title > b.Title) {return 1;}
				return 0;
			});
			console.log(items);
			return items;
		},
		error: function(data) {  
			alert("Error");  
		}  
	});

In this example, a Lookup field and a People field have been selected and expanded in addition to the rest of the returned item data.


## Create SharePoint Document Set

Call the function `createSharePointDocumentSet` to create a document set within a library at a given url. The functions `createFolder` and `getListUrl` are also used to complete this process.

### Parameters

- **siteUrl (Required)** - A required string value that is the site url that contains the library where you want to create the document set. This can be in another site collection as long as you have appropriate permissions.
- **listName (Required)** - A required string value that is the display name (not the internal name) of the library where you will create the document set.
- **folderName (Required)** - A required string value that is the name of the document set you will create.
- **success (Required)** - a required callback function that is called on successful creation.
- **error (Required)** - a required callback function that is call on unsuccessful creation.

### Usage

	createSharePointDocumentSet({
		siteUrl: _spPageContextInfo.webAbsoluteUrl, //SharePoint object that returns current site url
		listName: "My SharePoint Library",
		folderName: "New Document Set",
		success: function(folder) {
			alert('Document Set ' + folder.Name + ' has been created successfully');
		},
		error: function(error) {
			alert(JSON.stringify(error));
		}
	);


## Update File Data / Properties

Call the function `updateFileProperties` to update a file's data / properties. The function `executeJson` is also used to complete this process. Optional `.done()` and `.fail()` callbacks can also be added (see Usage).

### Parameters

- **siteUrl (Required)** - A required string value that is the site url that contains the library where you want to update the file data / properties.
- **fileUrl (Required)** - A required string value that is the url of the file you want to update.
- **itemProperties (Required)** - a required object that contains the data you want to update. Object keys must be the same as internal SharePoint field names to correctly update.

### Usage

	updateFileProperties({
		siteUrl: _spPageContextInfo.webAbsoluteUrl, //SharePoint object that returns current site url
		fileUrl: '/sites/MySharePointSite/MyLibrary/MyFolder/MyFile.pdf',
		itemProperties: {
			'Status': 'Complete',
			'Concert_x0020_Tickets': 20
		} //Object keys must match sharePoint field internal names to update correctly
	})
	.done(function() {
		alert('Your file properties have been successfully updated!');
	})
	.fail(function(error) {
		alert(JSON.stringify(error));
	});


## Upload a File Using a File Input

Add an html file input and upload button, and call the function `upload` on button click to upload the selected file to a library within the current site. The functions `getFileBuffer`, `addFileToFolder`, and `OnError` are also used to complete this process.

### Parameters

- **destinationUrl (Required)** - A required string url that is the library location where you want to upload the file.

### Usage

    <!-- HTML Markup -->

	<input id="getFile" type="file" multiple="multiple"/>
	<input id="addFileButton" type="button" value="Upload"/>

	$("document").ready(function(){
		//add click event to upload button to trigger upload function
		$("#addFileButton").click(function() {
			upload({
				destinationUrl: 'MySharePointLibrary/', //add folder to url here to upload to a specific folder
			});
		});	
	});


## Copy Library File to Another Library

Call the function `copy` to copy a file from one library to another.

###Parameters

- **webUrl (Required)** - A required string value that is the site url that contains the library where you want to update the file data / properties.
- **docToCopyUrl (Required)** - A required string value that is the url of the file you want to copy.
- **destinationUrl (Required)** - A required string url that is the library location where you want to upload the file.
- **success (Required)** - a required callback function that is called on successful copy.
- **error (Required)** - a required callback function that is call on unsuccessful copy.

### Usage

	copy({
		siteUrl: _spPageContextInfo.webAbsoluteUrl, //SharePoint object that returns current site url
		docToCopyUrl: '/sites/MySharePointSite/MyLibrary/MyFolder/MyFile.pdf',
		destinationUrl: 'NewLibrary/', //add folder to url here to upload to a specific folder
		success: function(sender, args) {
			console.log("Success! Your file was copied.");
		},
		error: function() {
			console.log("Something went wrong with document copy.");
		}
	});


## Delete a File

Call the function `deleteFile` to delete a file from a library within the current site.

### Parameters

- **fileUrl (Required)** - A required string value that is the url of the file you want to delete.
- **success (Required)** - a required callback function that is called on successful delete.
- **error (Required)** - a required callback function that is call on unsuccessful delete.

### Usage

	deleteFile({
		fileUrl: '/sites/MySharePointSite/MyLibrary/MyFolder/MyFile.pdf',
		success: function() {
			alert("Item successfully deleted.");
		},
		error: function() {
			console.log("Unable to delete item");
		}
	});


## Get Current User

Call the function `getCurrentUser` to retrieve current user information.

### Parameters

- **success (Required)** - a required callback function that is called on successful query.
- **error (Required)** - a required callback function that is call on unsuccessful query.

### Usage

    getCurrentUser({
		success: function (data) {  
			return data.d;  
		},
		error: function (data) {  
			alert(JSON.stringify(data));  
		}  
	});


## getCurrentUserGroupColl(userId)

Call the function `getCurrentUserGroupColl` to get all of the current user's group affiliation data from the current site collection.

### Parameters

- **userId (Required)** - A required integer value of the current user's ID.
- **success (Required)** - a required callback function that is called on successful query.
- **error (Required)** - a required callback function that is call on unsuccessful query.

### Usage

	getCurrentUserGroupColl({
		userId: _spPageContextInfo.userId, //SharePoint object that returns current user ID
		success: function (data) {   
			var results = data.d.results; 
			return results; 
		},
		error: function (data) {  
			alert(JSON.stringify(data));  
		}    
	});


## Get Current User Profile Picture

Call the function `getCurrentUserProfilePicture()` to retrieve the current user's profile picture url.

### Usage

    var currentUserProfilePictureUrl = getCurrentUserProfilePicture();
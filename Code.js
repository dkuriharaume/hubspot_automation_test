/* Testing */
function onOpen() {
   var ui = SpreadsheetApp.getUi(); 
   var menu = ui.createMenu('Test Scripts');
   menu.addItem('Test Authenticate', 'authenticate').addItem('Test Refresh', 'refresh');
   menu.addToUi();                            // メニューをUiクラスに追加する
}

/* End of Testing*/

/* Authentication */
var CLIENT_ID = 'a9e54a19-6ac5-44a4-b93e-19b423c55b95';     // Enter your Client ID
var CLIENT_SECRET = '91ae1051-086b-4720-8e15-c8faec77ed9f'; // Enter your Client secret
var SCOPE = 'contacts';
var AUTH_URL = 'https://app.hubspot.com/oauth/authorize';
var TOKEN_URL = 'https://api.hubapi.com/oauth/v1/token';
var API_URL = 'https://api.hubapi.com';

function getService() {
   return OAuth2.createService('hubspot')
      .setTokenUrl(TOKEN_URL)
      .setAuthorizationBaseUrl(AUTH_URL)
      .setClientId(CLIENT_ID)
      .setClientSecret(CLIENT_SECRET)
      .setCallbackFunction('authCallback')
      .setPropertyStore(PropertiesService.getUserProperties())
      .setScope(SCOPE);
}
function authCallback(request) {
   var service = getService();
   var authorized = service.handleCallback(request);
   if (authorized) {
      return HtmlService.createHtmlOutput('Success!');
   } else {
      return HtmlService.createHtmlOutput('Denied.');
   }
}
function authenticate() {
   var service = getService();
   if (service.hasAccess()) {
      // … whatever needs to be done here …
      var authorizationUrl = service.getAuthorizationUrl();
      Logger.log(authorizationUrl);
   } else {
      var authorizationUrl = service.getAuthorizationUrl();
      Logger.log('Open the following URL and re-run the script: %s', authorizationUrl);
   }
}
/* End Authentication */
/* Getting Data from HubSpot */
function getStages() {
   // Prepare authentication to Hubspot
   var service = getService();
   var headers = { headers: { 'Authorization': 'Bearer ' + service.getAccessToken() } };

   // API request
   var pipeline_id = "default"; // Enter your pipeline id here.
   var url = API_URL + "/crm-pipelines/v1/pipelines/deals";
   var response = UrlFetchApp.fetch(url, headers);
   var result = JSON.parse(response.getContentText());
   var stages = Array();

   // Looping through the different pipelines you might have in Hubspot
   result.results.forEach(function (item) {
      if (item.pipelineId == pipeline_id) {
         var result_stages = item.stages;
         // Let's sort the stages by displayOrder
         result_stages.sort(function (a, b) {
            return a.displayOrder - b.displayOrder;
         });

         // Let's put all the used stages (id & label) in an array
         result_stages.forEach(function (stage) {
            stages.push([stage.stageId, stage.label]);
         });
      }
   });

   return stages;
}

function getDealTypes() {
   // Prepare authentication to Hubspot
   var service = getService();
   var headers = { headers: { 'Authorization': 'Bearer ' + service.getAccessToken() } };

   // API request
   var url = API_URL + "/properties/v1/deals/properties";
   var response = UrlFetchApp.fetch(url, headers);
   var result = JSON.parse(response.getContentText());
   var dealTypes = Array();

   result.sort(function (a, b) {
      return a.displayOrder - b.displayOrder;
   });
   result.forEach(function (item) {

      if (item.label == "Deal Type") {
         Logger.log("Title: " + item.label);
         item.options.forEach(function (option) {
            dealTypes.push([option.value, option.label]);
            Logger.log("value: " + option.value + ", label" + option.label);
         });
      }
   });
   return dealTypes;
}

function getDeals() {
   // Prepare authentication to Hubspot
   var service = getService();
   var headers = { headers: { 'Authorization': 'Bearer ' + service.getAccessToken() } };
   // Prepare pagination
   // Hubspot lets you take max 250 deals per request.
   // We need to make multiple request until we get all the deals.
   var keep_going = true;
   var offset = 0;
   var deals = Array();
   while (keep_going) {
      // We’ll take three properties from the deals: the source, the stage & the amount of the deal
      var url = API_URL + "/deals/v1/deal/paged?properties=dealstage&properties=dealname&properties=amount&properties=dealtype&properties=acquisition_channel&properties=hubspot_owner_id&limit=250&offset=" + offset;
      var response = UrlFetchApp.fetch(url, headers);
      var result = JSON.parse(response.getContentText());
      // Are there any more results, should we stop the pagination
      keep_going = result.hasMore;
      offset = result.offset;
      // For each deal, we take the stageId, source & amount
      result.deals.forEach(function (deal) {
         var stageId = (deal.properties.hasOwnProperty("dealstage")) ? deal.properties.dealstage.value : "unknown";
         var dealName = (deal.properties.hasOwnProperty("dealname")) ? deal.properties.dealname.value : "unknown";
         var amount = (deal.properties.hasOwnProperty("amount")) ? deal.properties.amount.value : 0;
         var dealType = (deal.properties.hasOwnProperty("dealtype")) ? deal.properties.dealtype.value : "unknown";
         var acquisitionChannel = (deal.properties.hasOwnProperty("acquisition_channel")) ? deal.properties.acquisition_channel.value : "unknown";
         var dealOwnerId = (deal.properties.hasOwnProperty("hubspot_owner_id")) ? deal.properties.hubspot_owner_id.value : "unknown";
         deals.push([stageId, dealName, amount, dealType, acquisitionChannel, dealOwnerId]);
      });
   }
   return deals;
}

/* End of Getting Data from HubSpot */
/* Output to Sheets */
var sheetNameStages = "Stages";
var sheetNameDeals = "Deals";
var sheetNameDealTypes = "Deal Types";

function writeStages(stages) {
   var ss = SpreadsheetApp.getActiveSpreadsheet();
   var sheet = ss.getSheetByName(sheetNameStages);
   // Let’s put some headers and add the stages to our table
   var matrix = Array(["StageID", "Label"]);
   matrix = matrix.concat(stages);
   // Writing the table to the spreadsheet
   var range = sheet.getRange(1, 1, matrix.length, matrix[0].length);
   range.setValues(matrix);
}
function writeDealTypes(dealTypes) {
   var ss = SpreadsheetApp.getActiveSpreadsheet();
   var sheet = ss.getSheetByName(sheetNameDealTypes);
   // Let’s put some headers and add the stages to our table
   var matrix = Array(["Deal Type value", "label"]);
   matrix = matrix.concat(dealTypes);
   // Writing the table to the spreadsheet
   var range = sheet.getRange(1, 1, matrix.length, matrix[0].length);
   range.setValues(matrix);
}
function writeDeals(deals) {
   var ss = SpreadsheetApp.getActiveSpreadsheet();
   var sheet = ss.getSheetByName(sheetNameDeals);
   // Let’s put some headers and add the deals to our table
   var matrix = Array(["StageID", "Deal Name", "Amount", "Deal Type", "獲得チャネル", "担当者ID"]);
   matrix = matrix.concat(deals);
   // Writing the table to the spreadsheet
   var range = sheet.getRange(1, 1, matrix.length, matrix[0].length);
   range.setValues(matrix);
}

function refresh() {
   var service = getService();
   if (service.hasAccess()) {
      var stages = getStages();
      writeStages(stages);
      var deals = getDeals();
      writeDeals(deals);
      var dealTypes = getDealTypes();
      writeDealTypes(dealTypes);
      LogDate();
      mapValueLabel(stages,dealTypes);
   } else {
      var authorizationUrl = service.getAuthorizationUrl();
      Logger.log('Open the following URL and re-run the script: %s', authorizationUrl);
   }
}

var sheetNameMetadata = "my metadata";
function LogDate() {
   var date = new Date();
   Logger.log(date);
   var ss = SpreadsheetApp.getActiveSpreadsheet();
   var sheet = ss.getSheetByName(sheetNameMetadata);
   var matrix = Array();
   matrix.push(["最終更新日"]); matrix.push([date]);
   var range = sheet.getRange(1, 1, matrix.length, matrix[0].length);
   range.setValues(matrix);
}

var sheetNameMapping = "mapping";
function mapValueLabel(stage, dealtype) {
   var ss = SpreadsheetApp.getActiveSpreadsheet();
   var sheet = ss.getSheetByName(sheetNameMapping);
   var matrix = Array(["Deal stage ID", "Deal stage label", "Deal type value", "Deal type label"]);
   var maxRowCount = getMaxRowCount([stage, dealtype]);
   var dataRows = Array();
   for(var i=0; i<maxRowCount; i++){
      var row = Array();
      if(stage.length<i+1){
         row.push();row.push();
      }
      else
         row.push(stage[i][0],stage[i][1]);

      if(dealtype.length<i+1){
         row.push();row.push();
      }
      else
         row.push(dealtype[i][0],dealtype[i][1]);
      dataRows.push(row);
   }
   matrix.concat(dataRows);
   var range = sheet.getRange(1, 1, matrix.length, matrix[0].length);
   range.setValues(matrix);
}

function getMaxRowCount(pairArray){
   var counts=Array();
   for(var i=0; i<pairArray.length; i++){
      counts.push(pairArray[i].length);
   }
   return Math.max(counts);
}

/* End of Output to Sheets */
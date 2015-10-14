////oAuth related code

//hardcoded here for easily tweaking this. should move this to ScriptProperties or better parameterize them
//step 1. we can actually start directly here if that is necessary
var AUTHORIZE_URL = 'https://api.sandbox.freeagent.com/v2/approve_app'; 
//step 2. after we get the callback, go get token
var TOKEN_URL = 'https://api.sandbox.freeagent.com/v2/token_endpoint'; 

//PUT YOUR OWN SETTINGS HERE
var CLIENT_ID = 'your_client_id';
var CLIENT_SECRET='your_secret_key';
var REDIRECT_URL= ScriptApp.getService().getUrl();

//this is the user propety where we'll store the token, make sure this is unique across all user properties across all scripts
var tokenPropertyName = 'FREEAGENT_OAUTH_TOKEN5'; 
var baseURLPropertyName = 'FREEAGENT_INSTANCE_URL5';

var apiurl = 'https://api.sandbox.freeagent.com/v2/';
var perpage = 100;  // max 100
var days = 30;

// Process specific contacts    
var processContactArray = ['LLR', 'Cool Earth'];

// Colorize ranges
minRed = 0;
maxRed = minBlue = 20;
maxBlue = minAmber = 30;
maxAmber = minGreen = 40;

//Title Row styling
titleRowHeight = 30;
titleHorizontalAlignment = "center";
titleVerticalAlignment = "middle";
titleFontSize   = 13;
titleFontFamily = "Times New Roman";
titleFontWeight = "bold";
titleFontColor  = "black";
/* NOTE: any time sheet values are converted to an array instead of accessed directly, the indexing changes from 1 indexing to 0 indexing */
var documentProperties = PropertiesService.getDocumentProperties();

var orderSheet = SpreadsheetApp.getActive().getSheetByName('Form Responses 1');
var orders = orderSheet.getDataRange().getValues();
var orderHeaders = orders[0];

var configOptions = ['PayPal client ID', 'PayPal secret',
                     "Your first name", "Your last name (optional)",
                     "Your email (MUST match PayPal)", "Your website",
                     "Logo image URL (optional)", "Accept tips? (y/n)",
                     "US Shipping Cost", "International Shipping Cost  (optional)",
                     "Sales Tax Percent (optional)", "Notes for invoice"];
// associate human-readable names with machine-readable names
var configObjLookup = {
  'PayPal client ID': 'client',
  'PayPal secret': 'secret',
  'Your first name': 'firstName',
  'Your last name (optional)': 'lastName',
  'Your email (MUST match PayPal)': 'email',
  'Your website': 'website',
  'Logo image URL (optional)': 'logoUrl',
  'Accept tips? (y/n)': 'tips',
  'US Shipping Cost': 'usShippingCost',
  'International Shipping Cost  (optional)': 'internationalShippingCost',
  'Sales Tax Percent (optional)': 'salesTax',
  'Notes for invoice': 'invoiceNotes'
};
var configSheet = SpreadsheetApp.getActive().getSheetByName('Configuration');

var inventorySheet = SpreadsheetApp.getActive().getSheetByName('Inventory');

var isSandbox = false;

var baseUrl = 'https://api.paypal.com';

if (isSandbox) {
  baseUrl = 'https://api.sandbox.paypal.com';
}

var accessToken = {};

var itemInfo = {};

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('PayPal Invoicing')
      .addItem('Get Started', 'setup')
      .addSeparator()
      .addItem('Update Seller Config', 'parseSellerInfo')
      .addItem('Add New Inventory Items', 'parseAllItems')
      .addItem('Recount Inventory', 'recountInventory')
      .addToUi();
}

/**
  * Set up sheets.
  * Add trigger to send invoice when new submission is added.
  *
  */
function setup() {
  var triggers = ScriptApp.getProjectTriggers();
  if (triggers.length == 0) {
    ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(SpreadsheetApp.getActive()).onFormSubmit().create();
  }
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (inventorySheet == null) {
    inventorySheet = activeSpreadsheet.insertSheet();
    inventorySheet.setName("Inventory");
    inventorySheet.appendRow(['Item', 'Price', 'Qty Sold']);
  }

  if (configSheet == null) {
    configSheet = activeSpreadsheet.insertSheet();
    configSheet.setName("Configuration");
    var configHelpText = [["SELLER CONFIGURATION", ""],
                      ["This information will be displayed on each invoice and should be considered public.", ""],
                      ["To generate the PayPal client and secret tokens, follow the instructions on the 'Get Started' tab.", ""],
                      ["Make sure values are entered in column B, where it says 'VALUE'.", ""],
                      ["IMPORTANT: Email entered MUST match the email on your PayPal account.", ""],
                      ["IMPORTANT: After changing or updating values, click 'Update Seller Config' in the PayPal Invoicer menu.", ""],
                      ["", ""],
                      ["NAME", "VALUE"]];
    // build a 2D array of values from our previously defined configOptions and configHelpText
    // configOptions is used elsewhere and needs to stay as a 1D array otherwise
    var configValues = configHelpText.concat(configOptions.map(function(val, index, arr){ return [val, ""]; }));
    configSheet.getRange(1, 1, configValues.length, 2).setValues(configValues);
    configSheet.autoResizeColumns(1, 1);
  }

  var getStartedSheet = SpreadsheetApp.getActive().getSheetByName('Get Started');
  if (getStartedSheet == null) {
    getStartedSheet = activeSpreadsheet.insertSheet();
    getStartedSheet.setName("Get Started");
    var getStartedHelpText = [['Welcome to PayPal Invoicer!'],
                              [''],
                              ['There are a few steps you\'ll need to take to get this app working.'],
                              [''],
                              ['1) Create a new app for your PayPal account by following these instructions: https://developer.paypal.com/docs/api/overview/#get-credentials'],
                              ['Make sure to click the "Live" button at the top (A PAYPAL BUSINESS ACCOUNT IS REQUIRED FOR THIS). You can call your app whatever you like. I recommend something like "Google Form App".'],
                              ['Next, restrict the areas your new app has access to. Go to "App Settings" and uncheck everything except "Invoicing". This will help your credentials stay safe.'],
                              ['Creating a new app will give you a special client ID and secret. Copy and paste these in the appropriate fields on the "Configuration" tab.'],
                              [''],
                              ['2) Fill out the rest of the "Configuration" tab. If a field isn\'t marked as optional, you MUST fill it out.'],
                              [''],
                              ['3) When you\'re done filling out the "Configuration" tab, click "Update Seller Configuration" in the PayPal Invoicer menu.'],
                              ['If you change any of your information (like shipping price), make sure to click "Update Seller Configuration" again.'],
                              [''],
                              ['4) If you add new items for sale on your form, click "Add New Inventory Items". This will not remove any deleted items or update the price of existing items; only add new ones.'],
                              ['Don\'t modify anything on the "Inventory" tab by hand. Either your changes will be lost or you\'ll break the invoicer :)'],
                              [''],
                              ['5) If someone wants to change their order later, update the spreadsheet row with their form submission to reflect their new order. Then click "Recount Inventory" for the new number of items sold to repopulate.'],
                              ['Again, don\'t modify the "Inventory" tab by hand. Update the "Form Responses 1" tab instead, then run this function. This will also rebuild the inventory and add any new items.'],
                              ['You can also resend the invoice for that row if needed.'],
                              [''],
                              ['Need help? Message me on Telegram (@shirgoldbird) or email me@shirgoldbird.com']];
    getStartedSheet.getRange(1, 1, getStartedHelpText.length, 1).setValues(getStartedHelpText);
    getStartedSheet.autoResizeColumns(1, 1);
  }

  parseAllItems();
  inventorySheet.autoResizeColumns(1, 1);
}

/**
example of itemInfo obj:

{Stickers - $5 each [5. Trixie]:
   {
     price: $5,
     itemName: Trixie (Stickers),
     itemType: Stickers
    },
Tiny Prints - $5 each [4. Frog Ice Cream]:
  {
    itemName: Frog Ice Cream (Tiny Prints),
    price: $5,
    itemType: Tiny Prints
   }
}
*/

// parses a row representing an item named like "Holographic Stickers - $5 each [1. Mob-kun]"
// the [1. NAME] part is automatically generated by the Google Form grid
function parseItemInfo(item) {
  var itemNameRegex = /\[[0-9]+[a-z]?\.? (.*)\]$/;

  var splitText = item.split(" - ");

  var itemInfo = {};

  // if there's a hypen it's an item, otherwise it's not
  if (splitText[1]) {
    var itemType = splitText[0]; // Holographic Stickers
    var price = splitText[1].match(/[0-9]+/g)[0]; // $5
    var match = itemNameRegex.exec(splitText[1]);
    var itemName = match[1]; // Mob-kun

    // invoice will read like "Mob-kun (Holographic Stickers)
    var itemNameForInvoice = itemName + " (" + itemType + ")";

    itemInfo[item] = {
      'itemType': itemType,
      'price': price,
      'itemName': itemNameForInvoice
    };

    documentProperties.setProperty(itemInfo, price);
  }

  return itemInfo;
}

/**
  * Delete the current inventory, rebuild the item list, then recount the number of items sold.
  * Useful if a customer wants to change their order later.
  */
function recountInventory() {
  // delete inventory data, except the headers
  inventorySheet.getRange('A2:C').clearContent();

  // rebuild the inventory
  parseAllItems();

  // recount items
  var ordersRange = orderSheet.getDataRange();
  var orders = ordersRange.getValues();
  // Start at row 2, skipping headers in row 1
  // parseOrder will subtract 1 from this value
  for (var rowsWithoutHeader = 2; rowsWithoutHeader < orders.length + 1; rowsWithoutHeader++) {
    parseOrder(rowsWithoutHeader);
  }
}

// figure out what items we have for sale and how much they're selling for
// this will be run on initial setup and whenever the user clicks the "Rebuild Inventory" menu item
function parseAllItems() {
  var orders = orderSheet.getDataRange().getValues();
  var orderHeaders = orders[0];

  Logger.log(orderHeaders);

  for (var i = 0; i < orderHeaders.length; i++) {
    var item = orderHeaders[i];
    var itemObj = parseItemInfo(item);
    // check if parseItemInfo found an item
    if (itemObj[orderHeaders[i]]) {
      // we already have this item
      if (findTextInColumn(inventorySheet, 1, itemObj[item]['itemName'], false)) {
        continue;
      } else {
        inventorySheet.appendRow([itemObj[item].itemName, itemObj[item].price, 0]);
      }
    }
  }
}


// parses the Configuration tab in the sheet and stores the values in the documentProperties object
// https://developers.google.com/apps-script/reference/properties/properties#setProperties(Object)
function parseSellerInfo() {
  var config = configSheet.getDataRange().getValues();
  for (var i = 0; i < configOptions.length; i++) {
    var valueRow = findTextInColumn(configSheet, 1, configOptions[i], false).getRow() - 1;
    var value = config[valueRow][1];
    documentProperties.setProperty(configObjLookup[configOptions[i]], value);
  }
}

function parseBuyerInfo(rowId) {
  var orderIndex = rowId - 1;
  var buyerInfo = {};
  buyerInfo['name'] = {};
  buyerInfo['address'] = {};

  buyerInfo['name']['given_name'] = orders[orderIndex][getColumnIndexFromName(orderHeaders, 'First Name')];
  buyerInfo['name']['surname'] = orders[orderIndex][getColumnIndexFromName(orderHeaders, 'Last Name')];
  buyerInfo['address']['address_line_1'] = orders[orderIndex][getColumnIndexFromName(orderHeaders, 'Address Line 1')];
  buyerInfo['address']['address_line_2'] = orders[orderIndex][getColumnIndexFromName(orderHeaders, 'Address Line 2')];
  buyerInfo['address']['admin_area_2'] = orders[orderIndex][getColumnIndexFromName(orderHeaders, 'City')];
  // we don't want to try and infer the country code (like "US" or "CA") from this field, which is what PayPal requires
  var country = orders[orderIndex][getColumnIndexFromName(orderHeaders, 'Country')];
  // instead, we'll pass a fake country code to PayPal so it doesn't show anything for that field
  buyerInfo['address']['country_code'] = 'AA';
  // and then append the country to the state so the information still appears on the invoice
  buyerInfo['address']['admin_area_1'] = orders[orderIndex][getColumnIndexFromName(orderHeaders, 'State')] + ' ' + country;
  buyerInfo['address']['postal_code'] = orders[orderIndex][getColumnIndexFromName(orderHeaders, 'Zip')];
  buyerInfo['email_address'] = orders[orderIndex][getColumnIndexFromName(orderHeaders, 'Email')];

  // prepare to later apply different shipping costs if this order is international
  if (orders[orderIndex][getColumnIndexFromName(orderHeaders, 'international shipping')]) {
    var isInternational = orders[orderIndex][getColumnIndexFromName(orderHeaders, 'international shipping')].includes('Yes');
  } else {
    var isInternational = false;
  }
  buyerInfo['isInternational'] = isInternational;

  Logger.log('Buyer: ', buyerInfo);
  return buyerInfo;
}

// figure out how many items are being ordering from one row and format it for PayPal
function parseOrder(rowId) {
  var orderIndex = rowId - 1;
  var rowOrder = [];
  for (var i = 0; i < orders[orderIndex].length; i++) {
    var currCell = orders[orderIndex][i];
    // orderHeaders is 0 indexed
    var item = orderHeaders[i];
    var itemObj = parseItemInfo(item);
    if (currCell !== "None" && currCell !== "" && !isEmpty(itemObj)) {
      var order = {};
      var quantity = currCell.split(" ")[0];
      order['name'] = itemObj[item]['itemName'];
      order['quantity'] = quantity;
      order['unit_amount'] = {
        'currency_code': 'USD',
        'value': itemObj[item].price
      };
      rowOrder.push(order);

      var inventory = inventorySheet.getDataRange().getValues();
      var valueRow = findTextInColumn(inventorySheet, 1, order['name'], false).getRow();
      var qtySold = parseInt(inventory[valueRow - 1][2]);
      var updatedQtySold = qtySold + parseInt(quantity);

      inventorySheet.getRange(valueRow, 3).setValue(updatedQtySold);
    }
  };
  Logger.log('Order: ', rowOrder);
  return rowOrder;
}


function getToken() {
  var path = "/v1/oauth2/token";

  var headers = {
      'Authorization': 'Basic ' + Utilities.base64Encode(documentProperties.getProperty('client') + ':' + documentProperties.getProperty('secret')),
      'Accept': 'application/json',
      'Accept-Language': 'en_US'
  };

  var payload = {
    'grant_type': "client_credentials"
  };
  var payload_json = encodeURIComponent(JSON.stringify(payload));

  var options = {
      headers: headers,
      method: "POST",
      contentType: "application/json",
      payload: payload,
      muteHttpExceptions: true
  };

  var url = baseUrl + path;

  var expiration = new Date();

  var response = UrlFetchApp.fetch(url, options);
  var json = response.getContentText();
  var parsedResponse = JSON.parse(json);

  expiration.setSeconds(expiration.getSeconds() + parsedResponse["expires_in"]);

  Logger.log('attempted to generate token, response was ', parsedResponse);

  if (parsedResponse['error']) {
    Logger.log("Error generating token, exiting");
    return;
  }

  documentProperties.setProperty('accessToken', parsedResponse["access_token"]);
  documentProperties.setProperty('accessTokenExpiration', expiration);

  Logger.log('access token generated');
}

// max out at 5 retries for a single invoice
var retries = 0;

// cannot be used for initial token call as it requires an Authorization Basic header
function callPaypalUrl(path, payload) {
  Logger.log(documentProperties.getProperty('accessToken'));
  // check if token exists and is not expired
  if (documentProperties.getProperty('accessTokenExpiration')) {
    // get current time in seconds
    // https://stackoverflow.com/questions/3830244/get-current-date-time-in-seconds
    var currTime = Math.round(new Date() / 1000);
    if (documentProperties.getProperty('accessTokenExpiration') <= currTime) {
      Logger.log("token expired, generating new token...")
      getToken();
    }
  } else {
      Logger.log("no token found, generating new token...")
      getToken();
  }

  var headers = {
    'Authorization': 'Bearer ' + documentProperties.getProperty('accessToken'),
    'Content-Type': 'application/json'
  };

  var payload_json = JSON.stringify(payload);

  var options = {
    headers: headers,
    method: 'POST',
    payload: payload_json,
    contentType: 'application/json',
    muteHttpExceptions: true
  };

  var url = baseUrl + path;
  var response = UrlFetchApp.fetch(url, options);
  var json = response.getContentText();
  var parsedResponse = JSON.parse(json);

  Logger.log("PayPal URL ", path, " called, response is: ", parsedResponse);

  // sometimes the token may not be expired but may still be invalidated
  // if so, regenerate token and retry request
  if (parsedResponse['error'] == 'invalid_token' && retries < 5) {
    Logger.log("token marked invalid, generating new token and retrying...")
    getToken();
    retries++;
    callPaypalUrl(path, payload);
  }
  retries = 0;
  return parsedResponse;
}

function sendInvoice(order) {
  var currDate = new Date();
  var dateString = new Date(currDate.getTime() - (currDate.getTimezoneOffset() * 60000 ))
                    .toISOString()
                    .split("T")[0];

  var path = "/v2/invoicing/invoices";


  // payload schema at https://developer.paypal.com/docs/api/invoicing/v2/#invoices_create
  var payload = {
    "detail": {
      "invoice_date": dateString,
      "currency_code": "USD",
      "note": "Thanks, and enjoy your merch!",
      "payment_term": {
        "term_type": "DUE_ON_RECEIPT"
      }
    },
    "invoicer": {
      "name": {
      }
    },
    "configuration": {
      "allow_tip": true,
      "tax_inclusive": false
    },
    "amount": {
      "breakdown": {
        "shipping": {
          "amount": {
            "currency_code": "USD"
          },
          "tax": {
            "name": "Sales Tax",
            "percent": "0"
          }
        }
      }
    }
  };

  payload['detail']['note'] = documentProperties.getProperty('invoiceNotes');
  payload['invoicer']['name']['given_name'] = documentProperties.getProperty('firstName');
  payload['invoicer']['name']['surname'] = documentProperties.getProperty('lastName');
  payload['invoicer']['email_address'] = documentProperties.getProperty('email');
  payload['invoicer']['website'] = documentProperties.getProperty('website');
  payload['invoicer']['logo_url'] = documentProperties.getProperty('logoUrl');
  payload['configuration']['allow_tip'] = documentProperties.getProperty('tips') == 'y' ? true : false;

  payload['items'] = order['items'];

  // set billing and shipping info
  var buyerObj = {};
  buyerObj['billing_info'] = order['buyer'];
  buyerObj['shipping_info'] = order['buyer'];
  payload['primary_recipients'] = [buyerObj];

  // apply shipping costs and sales tax
  if (order['buyer']['isInternational']) {
    payload['amount']['breakdown']['shipping']['amount']['value'] = documentProperties.getProperty('internationalShippingCost');
  } else {
    payload['amount']['breakdown']['shipping']['amount']['value'] = documentProperties.getProperty('usShippingCost');
  }

  if (documentProperties.getProperty('salesTax') != '') {
    payload['amount']['breakdown']['shipping']['amount']['tax']['percent'] = documentProperties.getProperty('salesTax');
  }

  var draftInvoiceResponse = callPaypalUrl(path, payload);

  if ((draftInvoiceResponse.name && draftInvoiceResponse.name == "INVALID_REQUEST") || ('error' in draftInvoiceResponse)) {
    Logger.log("Error creating invoice, exiting")
    Logger.log(draftInvoiceResponse);
    return;
  }

  // the href will be something like https://api.sandbox.paypal.com/v2/invoicing/invoices/INV2-PYYR-W2ZN-ES7Y-THFZ
  var draftInvoiceUrl = draftInvoiceResponse.href;
  var draftInvoiceId = draftInvoiceUrl.split("/")[6]

  var sendInvoiceUrl = "/v2/invoicing/invoices/" + draftInvoiceId + "/send";

  var sendInvoice = callPaypalUrl(sendInvoiceUrl, { "send_to_invoicer": true });
}


function onFormSubmit(e) {
  if (SpreadsheetApp.getActiveSheet().getName() !== "Form Responses 1") return;
  var order = {};
  order['items'] = parseOrder(e.range.getRow());
  order['buyer'] = parseBuyerInfo(e.range.getRow());
  sendInvoice(order);
}

/** HELPER FUNCTIONS **/
// looks for partial match
// returns a 1-indexed value
function getColumnIndexFromName(headers, name) {
  for (var i = 0; i < headers.length; i++) {
    var re = new RegExp(name, 'g');
    if (headers[i].match(re)) return i;
  }
  return -1;
}

// an optimized twist on the createTextFinder method when we know our text will be in a specific column
// used to get the user's config values
// returns a Range or null
function findTextInColumn(sheet, columnNum, text, matchCase) {
  return sheet.getRange(columnNum,columnNum,sheet.getLastRow()).createTextFinder(text).matchCase(matchCase).findNext();
}

function isEmpty(obj) {
  for (var key in obj) {
    if (obj.hasOwnProperty(key)) {
      return false;
    }
  }
  return true;
}

/**
 * Test function for Spreadsheet Form Submit trigger functions.
 * Loops through content of sheet, creating simulated Form Submit Events.
 *
 * Check for updates: https://stackoverflow.com/a/16089067/1677912
 *
 * See https://developers.google.com/apps-script/guides/triggers/events#google_sheets_events
 */
function test_onFormSubmit_All() {
  var ordersRange = orderSheet.getDataRange();
  var orders = ordersRange.getValues();
  var headers = orders[0];
  // Start at row 1, skipping headers in row 0
  for (var offsetFromFirstRow = 1; offsetFromFirstRow < orders.length; offsetFromFirstRow++) {
    var e = {};
    e.range = ordersRange.offset(offsetFromFirstRow, 0, 1, headers.length);
    onFormSubmit(e);
  }
}

function test_onFormSubmit_Single() {
  if (SpreadsheetApp.getActiveSheet().getName() !== "Form Responses 1") return;
  var order = {};
  order['items'] = parseOrder(4);
  order['buyer'] = parseBuyerInfo(4);
  sendInvoice(order);
}

function onFormSubmit_Range() {
  var startRow = 3;
  var endRow = 5;

  // mess with the indexing so user can enter the number AS SHOWN IN THE SHEET
  startRow -= 1;
  endRow -= 1;
  var ordersRange = orderSheet.getDataRange();
  var orders = ordersRange.getValues();
  var headers = orders[0];
  // Start at row 1, skipping headers in row 0
  for (startRow; startRow <= endRow; startRow++) {
    var e = {};
    e.range = ordersRange.offset(startRow, 0, 1, headers.length);
    onFormSubmit(e);
  }
}

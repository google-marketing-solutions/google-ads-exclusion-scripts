const spreadsheetUrl = 'INSERT_SPREADSHEET_URL_HERE';
const sheetName = '3 - Topics exclusion';

/**
 * The code to execute when running the script.
 */
function main() {
  let readAccountListSheet = extractCustomerIds(getAccountListSheet());
  let customerIds = readAccountListSheet.account_ids;
  let limit = readAccountListSheet.limit;
  
  if (customerIds.length > 0) {
    for(let i = 0; i < customerIds.length; i = i+limit) {
      let customerIdsToProcess = customerIds.slice(i, i+limit) //process accounts in batches of limit
      AdsManagerApp
        .accounts()
        .withIds(customerIdsToProcess)
        .executeInParallel('processAccounts', 'postProcess');
    }

  } else {
    console.log ('No more accounts to process.');
  }
}

/**
 * Log the completion of the script and update number of accounts
 * processed in accounts list sheet
 */
function postProcess() {  
  console.log('Finished processing accounts');
}

/**
 * Applies the topic ids to all display and video campaigns
 * in the managed account
 *
 */
function processAccounts() {
  let topics =  readTopicsFromSheet();

  let campaigns = 
      AdsApp
        .campaigns()
        .withCondition("campaign.status IN (ENABLED, PAUSED)")
        .withCondition("campaign.advertising_channel_type IN (DISPLAY)")
        .get();
    
 for(const campaign of campaigns){         
    let campaignDisplay = campaign.display();
    let excludedTopics = campaignDisplay.excludedTopics().get();
    
    while(excludedTopics.hasNext()) {
      excludedTopics.next().remove();
    }    

    for(let topic of topics) {
      campaignDisplay.newTopicBuilder()
        .withTopicId(topic)
        .exclude();
    }
  }
  
  let videoCampaigns =
      AdsApp
        .videoCampaigns()
        .withCondition("campaign.status IN (ENABLED, PAUSED)")
        .get();
  

  while(videoCampaigns.hasNext()) {
    let videoCampaignTargeting = videoCampaigns.next().videoTargeting();

    let alreadyExcludedTopics = videoCampaignTargeting.excludedTopics().get();
    while(alreadyExcludedTopics.hasNext()) {
      alreadyExcludedTopics.next().remove();
    }
    
    for(let topic of topics) {
      videoCampaignTargeting
        .newTopicBuilder()
        .withTopicId(topic.toString())
        .exclude();
    }
  }
}

/**
 * Extracts topic ids from the exclusion sheet.
 *
 * @return {Array.<number>} A list of topic ids
 */
function readTopicsFromSheet() {
  let topics = 
      validateAndGetSpreadsheet(spreadsheetUrl)
        .getSheetByName(sheetName)
        .getDataRange()
        .getValues()
  
  return topics
        .slice(2)  //exclude first two rows of the topics sheet
        .map(topic => topic[1])
        .filter(topic => topic != '');
}

/**
 * Extracts customerIds from the account ids sheet.
 *
 * @param {string} data the input.
 * @return {Object<Array.<number>, number>} A list of customer IDs and configured limit.
 */
function extractCustomerIds(account_sheet) {
  
  let accountListSheet = account_sheet.getActiveSheet();
  const limit = accountListSheet.getRange('C2').getValue();

  let numOfRows = accountListSheet.getLastRow();
  let startRow = Number(2);

  let accountIds = accountListSheet
                    .getRange(startRow, 2, numOfRows - 1 , 1)
                    .getValues()
                    .map(accountId => accountId.toString());
  
  return {
    account_ids: accountIds,
    limit: limit
  };
}

/**
* Get the spreadsheet containing list of sub-accounts under this MCC
*
* @return {SpreadsheetApp.Spreadsheet}
* @throws error if no spreadsheet with the name is found
*/
function getAccountListSheet() {
  const accountName = AdsApp.currentAccount().getName();
  const name = '[' + accountName + '] Account List'
  const files = DriveApp.getFilesByName(name);

  let accountList = {};
  if (files.hasNext()) {
    let existingSheet = files.next();
    console.log('Fetched file with name', name, 'at ', existingSheet.getUrl());
    return SpreadsheetApp.open(existingSheet);
  } else {
    throw {
      'message': 'Cannot find spreadsheet with name: ' + name
    }
  }
}

/**
 * DO NOT EDIT ANYTHING BELOW THIS LINE.
 * Please modify your spreadsheet URL at the top of the file only.
 */

/**
 * Validates the provided spreadsheet URL
 * to make sure that they're set up properly. Throws a descriptive error message
 * if validation fails.
 *
 * @param {string} spreadsheeturl The URL of the spreadsheet to open.
 * @return {Spreadsheet} The spreadsheet object itself, fetched from the URL.
 */
function validateAndGetSpreadsheet(spreadsheeturl) {
  if (spreadsheeturl == 'INSERT_SPREADSHEET_URL_HERE') {
    throw new Error('Please specify a valid Spreadsheet URL. You can find' +
        ' a link to a template in the associated guide for this script.');
  }
  return SpreadsheetApp.openByUrl(spreadsheeturl);
}

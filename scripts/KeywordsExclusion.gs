/**
 * Configuration to be used for the Common Negative List Script for Google Ads
 * manager accounts.
 */
const SPREADSHEET_URL = 'INSERT_SPREADSHEET_URL_HERE'; //Replace here

/**
 * Keep track of the spreadsheet names for various criteria types, as well as
 * the criteria type being processed.
 */
const Criteria = {
  KEYWORDS: {
    type: 'Keywords',
    sheet: '2 - Keywords exclusion'
  },
  LISTNAME: '[Script-Created] Common List'
};

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
        .withIds(customerIds)
        .executeInParallel('processAccounts', 'postProcess');
    }

  } else {
    console.log ('No more accounts to process.');
  }
}

/**
 * Process an account when processing multiple accounts under a Google Ads
 * manager account in parallel.
 *
 * @return {string} A JSON string that summarizes the number of keywords
 *     synced, and the number of campaigns processed.
 */
function processAccounts() {
  return JSON.stringify(syncMasterLists());
}

/**
 * Callback method after processing accounts, when processing multiple accounts
 * under a Google Ads manager account in parallel.
 *
 * @param {Array.<AdsManagerApp.ExecutionResult>} results The execution results
 *     from the accounts that were processed by this script.
 */
function postProcess(results) {

  const resultParams = {
    // Number of keywords that were synced.
    KeywordCount: 0,    
    // Summary of customers who were synced.
    Customers: {
      // How many customers were synced?
      Success: 0,
      // How many customers failed to sync?
      Failure: 0,
      // Details of each account processed. Contains 3 properties:
      // CustomerId, CampaignCount, Status.
      Details: []
    }
  };

  for (const result of results) {
    const customerResult = {
      // The customer ID that was processed.
      CustomerId: result.getCustomerId(),
      // Number of campaigns that were synced.
      CampaignCount: 0,
      // Status of processing this account - OK / ERROR / TIMEOUT
      Status: result.getStatus()
    };

    if (result.getStatus() == 'OK') {
      let retval = JSON.parse(result.getReturnValue());
      customerResult.CampaignCount = retval.CampaignCount;
      if (resultParams.Customers.Success == 0) {
        resultParams.KeywordCount = retval.KeywordCount;
      }
      resultParams.Customers.Success++;
    } else {
      resultParams.Customers.Failure++;
    }
    resultParams.Customers.Details.push(customerResult);
  }
  logResults(resultParams);
  
  if (!AdsApp.getExecutionInfo().isPreview()) {
    let accountListSheet = getAccountListSheet().getActiveSheet();
    let limit = accountListSheet.getRange('C2').getValue();
    let numOfProcessedAccounts = accountListSheet.getRange('D2');
    numOfProcessedAccounts.setValue(numOfProcessedAccounts.getValue() + limit);
  }
}

/**
 * Logs the changes that this script made.
 *
 * @param {Object} resultParams Contains details of the result
 */
function logResults(resultParams) {
  console.log("Processing completed. Results:");
  console.log("Keywords applied: ", resultParams.KeywordCount);
  console.log("CIDs successfully applied: ", resultParams.Customers.Success);
  console.log("CIDs failed to apply: ", resultParams.Customers.Failure);

  for (const detail of resultParams.Customers.Details) {
    console.log('CID: ' + detail.CustomerId +
                ', campaigns processed: ' + detail.CampaignCount +
                ', status: ' + detail.Status);
  }
}

/**
 * Synchronizes the negative criteria list in an account with the master list
 * in the user spreadsheet.
 *
 * @return {Object} A summary of the number of keywords synced,
 *     and the number of campaigns to which these lists apply.
 */
function syncMasterLists() {
  let syncedCampaignCount = 0;
  
  const keywordListDetails = syncCriteriaInNegativeList(Criteria.KEYWORDS.type, Criteria.KEYWORDS.sheet);
  syncedCampaignCount = syncCampaignList(keywordListDetails.SharedList, Criteria.KEYWORDS.type);
  
  return {
    'CampaignCount': syncedCampaignCount,
    'KeywordCount': keywordListDetails.CriteriaCount
  };
}

/**
 * Synchronizes the list of campaigns covered by a negative list against the
 * desired list of campaigns to be covered by the master list.
 *
 * @param {AdsApp.NegativeKeywordList}
 *    sharedList The shared negative criterion list to be synced against the
 *    master list.
 * @param {String} criteriaType The criteria type for the shared negative list.
 *
 * @return {Number} The number of campaigns synced.
 */
function syncCampaignList(sharedList, criteriaType) {
  const campaignIds = getCampaigns();
  const totalCampaigns = Object.keys(campaignIds).length;
  
  const listedCampaigns = sharedList.campaigns().get();
  const listedShoppingCampaigns = sharedList.shoppingCampaigns().get();
  let allAvailableCampaigns = [].push(listedCampaigns, listedShoppingCampaigns);

  for (let index in allAvailableCampaigns) {
    for (const availableCampaign of allCampaignsToAddList[index]) {
      if (availableCampaign.getId() in campaignIds) {
        delete campaignIds[allAvailableCampaigns[index]];
      } else {
        campaignsToRemove.push(allAvailableCampaigns[index]);
      }
    }
  }


  const campaignsToRemove = [];
  
 
  // Anything left over in campaignIds starts a new list.

  const campaignsToAdd = AdsApp.campaigns().withIds(
      Object.keys(campaignIds)).get();
  const videoCampaignsToAdd = AdsApp.videoCampaigns().withIds(
      Object.keys(campaignIds)).get();
  const shoppingCampaignsToAdd = AdsApp.shoppingCampaigns().withIds(
      Object.keys(campaignIds)).get();
  
  let allCampaignsToAddList = [];
  allCampaignsToAddList.push(campaignsToAdd, videoCampaignsToAdd, shoppingCampaignsToAdd);
  
  for(let index in allCampaignsToAddList) {
    for (const campaignToAdd of allCampaignsToAddList[index]) {
      campaignToAdd.addNegativeKeywordList(sharedList);
    }
  }

  for (const campaignToRemove of campaignsToRemove) {
    campaignToRemove.removeNegativeKeywordList(sharedList);
  }

  return totalCampaigns;
}

/**
 * Gets a list of campaigns.
 *
 *
 * @return {Array.<Number>} An array of campaign IDs 
 */
function getCampaigns() {

  const campaignIds = {};
  let campaigns;
  let videoCampaigns;
  let shoppingCampaigns;
  
  campaigns = AdsApp.campaigns().withCondition(
    'Status in [ENABLED, PAUSED]').get();
  videoCampaigns = AdsApp.videoCampaigns().withCondition(
    'Status in [ENABLED, PAUSED]').get()
  shoppingCampaigns = AdsApp.shoppingCampaigns().withCondition(
    'Status in [ENABLED, PAUSED]').get()
 
  for (const video of videoCampaigns) {
    campaignIds[video.getId()] = 1;
  }
  for (const campaign of campaigns) {
    campaignIds[campaign.getId()] = 1;
  }
  for (const shopping of shoppingCampaigns) {
    campaignIds[shopping.getId()] = 1;
  }

  return campaignIds;
}

/**
 * Synchronizes the criteria in a shared negative criteria list with the user
 * spreadsheet.
 *
 * @param {String} criteriaType The criteria type for the shared negative list.
 *
 * @return {Object} A summary of the synced negative list, and the number of
 *     criteria that were synced.
 */
function syncCriteriaInNegativeList(criteriaType, sheetName) {
  const criteriaFromSheet = loadCriteria(criteriaType, sheetName);
  const totalCriteriaCount = Object.keys(criteriaFromSheet).length;

  let sharedList = null;
  let listName = Criteria.LISTNAME;

  sharedList = createNegativeListIfRequired(listName, criteriaType);


  let negativeCriteria = null;

  try {
    negativeCriteria = sharedList.negativeKeywords().get();
  } catch (e) {
    console.error(`Failed to retrieve shared list. Error says ${e}`);
    if (AdsApp.getExecutionInfo().isPreview()) {
      let message = Utilities.formatString(`The script cannot create the ` +
          `negative ${criteriaType} list in preview mode. Either run the ` +
          `script without preview, or create a negative ${criteriaType} list ` +
          `with name "${listName}" manually before previewing the script.`);
      console.log(message);
    }
    throw e;
  }
  
  for (const negativeCriterion of negativeCriteria) {
    negativeCriterion.remove();
  }

  sharedList.addNegativeKeywords(Object.keys(criteriaFromSheet));

  return {
    'SharedList': sharedList,
    'CriteriaCount': totalCriteriaCount,
    'Type': criteriaType
  };
}


/**
 * Creates a shared negative criteria list if required.
 *
 * @param {string} listName The name of shared negative criteria list.
 * @param {String} listType The criteria type for the shared negative list.
 *
 * @return {AdsApp.NegativeKeywordList} An
 *     existing shared negative criterion list if it already exists in the
 *     account, or the newly created list if one didn't exist earlier.
 */
function createNegativeListIfRequired(listName, listType) {
  let negativeListSelector = null;
  if (listType == Criteria.KEYWORDS.type) {
    negativeListSelector = AdsApp.negativeKeywordLists();
  }
  let negativeListIterator = negativeListSelector.withCondition(
      `shared_set.name = '${listName}'`).get();

  if (negativeListIterator.totalNumEntities() == 0) {
    let builder = AdsApp.newNegativeKeywordListBuilder();
    let negativeListOperation = builder.withName(listName).build();

    return negativeListOperation.getResult();
  } else {
    return negativeListIterator.next();
  }
}

/**
 * Loads a list of criteria from the user spreadsheet.
 *
 * @param {string} sheetName The name of shared negative criteria list.
 *
 * @return {Object} A map of the list of criteria loaded from the spreadsheet.
 */
function loadCriteria(criteria, sheetName) {
  const spreadsheet = validateAndGetSpreadsheet(SPREADSHEET_URL);  
  const sheet = spreadsheet.getSheetByName(sheetName);
  let values = [];


  const range = sheet.getRange('C:J').getValues(); //Specify the columns to include
  values = range.slice(3).flat(); // exclude the first 3 rows from the keywords sheet
  
  const retval = {};
  for (const value of values) {
    let keyword = value.toString().trim();
    if (!isEmptyKeyword(keyword)) {
      retval[keyword] = 1;
    }
  }
  return retval;
}

/**
* Checks if the supplied keyword is empty
*
* @param {string} keyword to check
* @return {Boolean}
*/
function isEmptyKeyword(keyword) {
  return (keyword == '' || 
          keyword == "" ||
          keyword == '[]' || 
          keyword == "[]");
}

/**
 * Extracts customerIds from the account ids sheet.
 *
 * @param {string} data the input.
 * @return {Array.<number>} A list of customer IDs.
 */
function extractCustomerIds(account_sheet) {
  let accountListSheet = account_sheet.getActiveSheet();
  const LIMIT = accountListSheet.getRange('C2').getValue();

  let numOfRows = accountListSheet.getLastRow();
  let startRow = Number(2);

  let accountIds = accountListSheet
                    .getRange(startRow, 2, numOfRows - 1 , 1)
                    .getValues()
                    .map(accountId => accountId.toString());
  
  return { 
    limit: LIMIT,
    account_ids: accountIds
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
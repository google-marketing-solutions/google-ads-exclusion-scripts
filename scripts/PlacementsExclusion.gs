const PLACEMENT_LIST_NAME = 'INSERT_LIST_NAME_HERE'

/**
* The main function that is executed when the script is run
* 
* This function reads a placement exclusion list from the MCC
* and applies to child accounts based on the account list
*/
function main() {
  let placementLists = 
      AdsApp.search("SELECT shared_set.name, shared_set.id, shared_set.resource_name, shared_set.type " + 
                    "FROM shared_set WHERE shared_set.type = 'NEGATIVE_PLACEMENTS' AND shared_set.name LIKE '" + PLACEMENT_LIST_NAME + "'");  
  let placementList = placementLists.next();
  
  let placements = AdsApp.search("SELECT shared_criterion.type, shared_criterion.placement.url, shared_criterion.youtube_channel.channel_id, shared_criterion.youtube_video.video_id " + 
                                 "FROM shared_criterion WHERE shared_set.id = " + placementList.sharedSet.id);  

  const placementsMap = {
    'placement': [],
    'youtubeChannel': [],
    'youtubeVideo': []
    
  }
  while(placements.hasNext()) {    
    let placement = placements.next().sharedCriterion;
    
    if(placement.type == 'PLACEMENT') {
      placementsMap.placement.push(placement.placement.url)      
    }
    if(placement.type == 'YOUTUBE_CHANNEL') {
      placementsMap.youtubeChannel.push(placement.youtubeChannel.channelId)
    }
    if(placement.type == 'YOUTUBE_VIDEO') {
      placementsMap.youtubeVideo.push(placement.youtubeVideo.videoId)
    }
  }
  
  let placementsMapAsString = JSON.stringify(placementsMap);

  let readAccountListSheet = extractCustomerIds(getAccountListSheet());
  let customerIds = readAccountListSheet.account_ids;
  let limit = readAccountListSheet.limit;
  
  if (customerIds.length > 0) {
    for(let i = 0; i < customerIds.length; i = i+limit) {
      let customerIdsToProcess = customerIds.slice(i, i+limit) //process accounts in batches of limit
      AdsManagerApp
        .accounts()
        .withIds(customerIdsToProcess)
        .executeInParallel('processAccounts', 'postProcess', placementsMapAsString);
    }

  } else {
    console.log ('No more accounts to process.');
  }
}

function processAccounts(placementsMapAsString) {
  let placementsMap = JSON.parse(placementsMapAsString);
  
  let campaigns = AdsApp.campaigns().get();
  applyExclusions(placementsMap, campaigns, (campaign) => campaign.display());
  
  let videoCampaigns = AdsApp.videoCampaigns().get();  
  applyExclusions(placementsMap, videoCampaigns, (campaign) => campaign.videoTargeting()); 
}

/**
* 
* The function to call as the last step of the script 
*/
function post(result) {
  console.log('Execution Completed:', result)
}


/**
*
* Apply website, channel and video exclusions to campaigns
*
* @param {Object}placementsMap Map of all exclusion criterion in the MCC level list
* @param {Array<Campaign>}campaigns List of all campaigns in this account
* @param {Function}fnSelectTargeting Function to pick the right targeting for campaigns
*/
function applyExclusions(placementsMap, campaigns, fnSelectTargeting) {  
  while(campaigns.hasNext()) {
    let campaign = campaigns.next();
    let targeting = fnSelectTargeting(campaign);
    
    let alreadyExcludedChannels = targeting.youTubeChannels().get();  
    let channelsToExclude = filterAlreadyExcluded(placementsMap.youtubeChannel, alreadyExcludedChannels, (placement) => placement.getChannelId());
  
    let alreadyExcludedVideos = targeting.youTubeVideos().get();
    let videosToExclude = filterAlreadyExcluded(placementsMap.youtubeVideo, alreadyExcludedVideos, (placement) => placement.getVideoId())
    
    let alreadyExcludedPlacements = targeting.excludedPlacements().get();
    let urlsToExclude = filterAlreadyExcluded(placementsMap.placement, alreadyExcludedPlacements, (placement) => placement.getUrl());
            
    for(let index in channelsToExclude) {
      targeting
        .newYouTubeChannelBuilder()
        .withChannelId(channelsToExclude[index])
        .exclude();       
    }
    
    for(let index in videosToExclude) {
      targeting
        .newYouTubeVideoBuilder()
        .withVideoId(videosToExclude[index])
        .exclude();       
    }
    
    for(let index in urlsToExclude) {
      targeting
        .newPlacementBuilder()
        .withUrl(urlsToExclude[index])
        .exclude();       
    }
  }
}

/**
* Filter out placements that have already been applied to this campaign
*
* @param {Array<String>} placement List of placements for this type of placement (website/channel/video)
* @param {AdsApp.ExcludedPlacementIterator|ExcludedVideoPlacementIterator} alreadyExcluded Iterator of placement already excluded for this campaign
* @param {Function} fnGetIdentifier Function to get the identifier for this type of placement (url/channelId/videoId)
*/
function filterAlreadyExcluded(placements, alreadyExcluded, fnGetIdentifier) {
  let excluded = [];
  
  while(alreadyExcluded.hasNext()) {
   let value = fnGetIdentifier(alreadyExcluded.next());
   excluded.push(value);
  }
    
  return excluded.length > 0 ? 
    placements.filter(placement => !excluded.find(exc => exc == placement)) :
    placements;
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

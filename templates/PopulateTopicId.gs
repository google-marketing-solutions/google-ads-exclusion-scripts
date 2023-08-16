function populateTopicId() {
  let currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  let topicIdMappingSheet = currentSpreadsheet.getSheetByName('[DO NOT EDIT] Topic IDs');
  fetchLatestVerticalMapping(topicIdMappingSheet);
  let formattedTopicToId = getTopicIdMapping(topicIdMappingSheet);

  updateTopicsSheet(currentSpreadsheet, formattedTopicToId);
}

function fetchLatestVerticalMapping(topicIdMappingSheet) {
  let verticals = UrlFetchApp
    .fetch('https://developers.google.com/static/google-ads/api/data/tables/verticals.csv')
    .getAs('text/csv');

  let data = Utilities.parseCsv(verticals.getDataAsString());
  
  topicIdMappingSheet.getRange(1,1,data.length, 3).setValues(data);
}

function getTopicIdMapping(topicIdMappingSheet) {
  let topicIdMapping = 
    topicIdMappingSheet
      .getDataRange()
      .getValues();

  return topicIdMapping
  .slice(1) //exclude header row  
  .map(row => {
    let topic = row[2].split('/'); //map topic to criterion id
    return {
      topic: topic[topic.length - 1],
      id: row[0]
    }
  }); 
}

function updateTopicsSheet(currentSpreadsheet, formattedTopicToId) {
  let topicsSheet = 
    currentSpreadsheet
      .getSheetByName('3 - Topics exclusion');

  let topicsSheetDataRange = topicsSheet.getRange(3,1,topicsSheet.getLastRow(),2);

  let topicToIdMap = 
    topicsSheetDataRange     
      .getValues()
      .map(topic => {
        if(topic[1] == '') {
          let matchedTopic = formattedTopicToId.find(topicToId => topicToId.topic == topic[0]);
          return matchedTopic ? [topic[0], matchedTopic.id]: topic;
          } 
        return topic;   
      });

  topicsSheetDataRange.setValues(topicToIdMap);
}
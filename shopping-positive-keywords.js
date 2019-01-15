/**
* AutomatingAdWords.com - Auto-negative keyword adder
*
* This script will automatically add negative keywords based on specified criteria
* Go to automatingadwords.com for installation instructions & advice
*
* Version: 1.7.0
**/

var SPREADSHEET_URL = "your-settings-url-here"; //template here: https://docs.google.com/spreadsheets/d/1G8mjtESGW3O1Jtd3l9MvqPA93cwUvFdVwkWsHOvT0W4/edit#gid=0
var SS = SpreadsheetApp.openByUrl(SPREADSHEET_URL);

//You may also be interested in this Chrome keyword wrapper!
//https://chrome.google.com/webstore/detail/keyword-wrapper/paaonoglkfolneaaopehamdadalgehbb



function main() {

  var timeStampCol = 7

  var sheets = SS.getSheets();


  var times = new Object();
  var timesArray = [];
  //get timestamps
  var i = 0;
  for(var sheetNo in sheets){
    
    var sheet = sheets[sheetNo];

    var timestamp = sheet.getRange(1, timeStampCol).getValue();
    
    if(timestamp){
      times[timestamp] = sheet.getName();
      timesArray.push(timestamp);
    }else{
      //first run
      var oldenTimes = new Date();
      oldenTimes.setDate(oldenTimes.getDate() - (1000+i));
      times[oldenTimes] = sheet.getName();
      timesArray.push(oldenTimes);
    }
    
    i++;
    
  }
  
  //loop through sheets, starting with the earliest time
  timesArray = timesArray.sort(function(a, b){return a-b});
  
  for(var time in timesArray){
    var sheetName = times[timesArray[time]];
    Logger.log("Checking sheet: " + sheetName);
    var sheet = SS.getSheetByName(sheetName);

    var SETTINGS = {}
    SETTINGS["CAMPAIGN_NAME"] = sheet.getRange("A2").getValue()
    SETTINGS["MIN_QUERY_CLICKS"] = sheet.getRange("B2").getValue()
    SETTINGS["MAX_QUERY_CONVERSIONS"] = sheet.getRange("C2").getValue()
    SETTINGS["DATE_RANGE"] = sheet.getRange("D2").getValue()
    SETTINGS["NEGATIVE_MATCH_TYPE"] = sheet.getRange("E2").getValue()
    SETTINGS["CAMPAIGN_LEVEL_QUERIES"] = sheet.getRange("F2").getValue()=="Yes" ? true : false

    Logger.log("Settings: "+ JSON.stringify(SETTINGS))

    var adGroups = [];
    var adGroupMins = [];
    var adGroupColumnNumbers = [];
    var col = 2;

    var numMatchesRow = 4
    var adGroupNameRow = 5
    var firstAdGroupRow = 6
    
    //grab adGroup data from sheet,store in arrays
    while(sheet.getRange(numMatchesRow, col).getValue()){
      adGroups.push(sheet.getRange(adGroupNameRow, col).getValue());
      adGroupColumnNumbers.push(col);
      adGroupMins.push(sheet.getRange(numMatchesRow, col).getValue());
      col++;
    }
    
    //loop through the adGroups listed in the sheet
    for(var ag in adGroups){
      
      var adGroupName = adGroups[ag];
      Logger.log("Checking AdGroup: "+ adGroupName);
      var row = firstAdGroupRow;
      var keywords = [];
      //get the "positive" keywords from the sheet, for this adGroup
      while(sheet.getRange(row, adGroupColumnNumbers[ag]).getValue()){
        keywords.push(sheet.getRange(row, adGroupColumnNumbers[ag]).getValue());
        row++;
      }
      
      Logger.log("'Positive Keywords' from sheet: "+ keywords);
      //get the search queries from the campaign, we'll check against these

      var query =  "SELECT Query " +
      " FROM SEARCH_QUERY_PERFORMANCE_REPORT " 
      + " WHERE CampaignName = '"+SETTINGS["CAMPAIGN_NAME"]   + "'"

      if(SETTINGS["MIN_QUERY_CLICKS"]!=""){
        query+= " AND Clicks > " + SETTINGS["MIN_QUERY_CLICKS"]
     }


      if(SETTINGS["MAX_QUERY_CONVERSIONS"]!=""){
         query+= " AND Conversions < " + SETTINGS["MAX_QUERY_CONVERSIONS"]
      }

      if(!SETTINGS["CAMPAIGN_LEVEL_QUERIES"]){
        query+= " AND AdGroupName = '" + adGroupName + "'"
      }

      if(SETTINGS["DATE_RANGE"]!= "ALL_TIME"){

        query+= " DURING "+SETTINGS["DATE_RANGE"]

      }
      
      log("query: " + query)
      var report = AdWordsApp.report(query);
      
      var rows = report.rows();
      var negs = [];  
      //loop through this campaign's queries, add anything which doesn't contain our positive keywords to the negs array (these will be added as negatives later)
      while(rows.hasNext()){
        var nxt = rows.next();
        var q = nxt.Query;
        var matches = 0;
        var count = 0; 
        
        //loop through the positive keywords (from the sheet)
        for(var k in keywords){
          //if > min (e.g. 2) of the keywords are in the search term, then neg it
          // Logger.log("Checking against keyword: " + keywords[k]);
          // Logger.log("Checking against term:    " + q);
          count++;
          //if the keyword is in the query, we have a match. match++
          if(q.indexOf(keywords[k]) > -1){
            //Logger.log(nxt.Query + " - " + q.indexOf(keywords[k]) + " - " + keywords[k] );
            matches++;
          }
          // Logger.log("matches: " + matches);
          // Logger.log("count: " + count + " - " + keywords.length);
          
          //if we have reached the end of the positive keywords i.e. checked them all
          //and if the number of matches is leSS than the minimum number of matches for the adgroup (specified on the sheet)
          //then add the query to the negatives array
          if(matches < adGroupMins[ag] && count == keywords.length){
            // Logger.log(count + " - " + keywords.length);
            //Logger.log("adding negative: " + nxt.Query);
            negs.push(q);
            break;
          }
          
        }    
        
      }
      Logger.log("Found a total of "+negs.length+" negative keywords to add")
      Logger.log("Adding the negative keywords...")

      var adGroupFound = false;

      //we have the negs. Now add them to the adgroup...
      var iterTypes = ["Shopping", "Text"]

      for(var t in iterTypes){

        var type = iterTypes[t]

        if(type=="Shopping"){

          var adGroupIterator = AdWordsApp.shoppingAdGroups()

        }else{

          var adGroupIterator = AdWordsApp.adGroups()

        }

        adGroupIterator = adGroupIterator
        .withCondition("Name = '"+adGroupName+"'")
        .withCondition("CampaignName = '"+SETTINGS["CAMPAIGN_NAME"] +"'")
        .get();
        
          if(adGroupIterator.hasNext()){

            adGroupFound = true;

            var adGroup = adGroupIterator.next();

            for(var neg in negs){

              var neg = addMatchType(negs[neg],SETTINGS)
              adGroup.createNegativeKeyword(neg);
              
            }   

          }
     

      }//end ad types loop

      if(!adGroupFound){
        Logger.log("AdGroup '" + adGroupName + "' in Campaign '" +SETTINGS["CAMPAIGN_NAME"] + "' not found in the account. Check the AdGroup name is correct in the sheet.");
      }        
    
    }
    
    //timestamp
    var date = new Date();
    sheet.getRange(1, timeStampCol).setValue(date);
    
  }//end time array loop
  Logger.log("Finished")
  
}//end main

function addMatchType(word,SETTINGS){
  if(SETTINGS["NEGATIVE_MATCH_TYPE"].toLowerCase()=="broad"){
  word = word.trim();
  }else if(SETTINGS["NEGATIVE_MATCH_TYPE"].toLowerCase()=="bmm"){
  word = word.split(" ").map(function (x){return "+"+x}).join(" ").trim()
  }else if(SETTINGS["NEGATIVE_MATCH_TYPE"].toLowerCase()=="phrase"){
  word = '"' + word.trim() +  '"'
  }else if(SETTINGS["NEGATIVE_MATCH_TYPE"].toLowerCase()=="exact"){
  word = '[' + word.trim() +  ']'
  }else{
  throw("Error: Match type not recognised. Please provide one of Broad, BMM, Exact or Phrase")
  }
  return word;
}

function log(x){Logger.log(x)}

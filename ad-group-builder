/**
* AutomatingAdWords.com - Ad Group Builder
*
* Go to automatingadwords.com for installation instructions and advice
*
* Version: 1.4
**/
  
//your spreadsheet URL
var SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1oQ7pCFk8fAMLwv7JOgK-kmS9aNfR5Jfyn1tYLYwZ_Ms/edit#gid=0";
//your sheet (tab) name
var SHEET_NAME = "Builder";

//Cell Locations - Only to be changed if the sheet changes
var firstAdGroupRow = 9; //row of the first AdGroup, 9 is default
var urlColumn = 1;
var adGroupColumn = 2;
var keywordsColumn = 3;
var negativeKeywordColumn = 4;
var headline1Column = 5;
var headline2Column = 7;
var displayUrl1Column = 9;
var displayUrl2Column = 11;
var descriptionColumn = 13;

var ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
var sheet = ss.getSheetByName(SHEET_NAME);

//OPTIONS - update on sheet
var campaignName = sheet.getRange(2, 2).getValue();
var createGroups = sheet.getRange(3, 2).getValue();
var createAds_bool = sheet.getRange(4, 2).getValue();
var createKeywords_bool = sheet.getRange(5, 2).getValue();
var exactBid = sheet.getRange(2,4).getValue();
var phraseBid = sheet.getRange(3, 4).getValue();
var broadBid = sheet.getRange(4, 4).getValue();
var addUrlsToKeywords = sheet.getRange(5, 4).getValue();
var pauseExistingKeywords_bool = sheet.getRange(6, 2).getValue();
var adGroupLabel = sheet.getRange(2, 6).getValue();
var adLabel = sheet.getRange(3, 6).getValue();
var keywordLabel = sheet.getRange(4, 6).getValue();
var keywordGroups = [];
var adGroupsToAdd = [];



function main() {
  
  //create the labels if they don't exist
  createLabel(adGroupLabel)
  createLabel(adLabel)
  createLabel(keywordLabel)
  
  if(createGroups == true){
    // Logger.log(createGroups)
    createAdGroups();}
  if(createAds_bool == true){
    createAds();}
  createKeywords();
  if(pauseExistingKeywords_bool == true){
    pauseExistingKeywords();}
  Logger.log("All done!")
}

function createLabel(labelName){
  if(labelName==""){return}
  
        var labels = AdWordsApp.labels().get();
        var iter = 0;
        while(labels.hasNext()){
          var label = labels.next();
          if(label.getName() == labelName){
            iter++;
          }
        }
        if(iter==0){
         AdWordsApp.createLabel(labelName); 
        }
}

function pauseExistingKeywords(){
  Logger.log("Pausing existing keywords...");
  //Logger.log("keyword groups: " + keywordGroups);
  
  for(var kwGroup_i in keywordGroups){
    var keywordsToAdd = keywordGroups[kwGroup_i];
  
  //Logger.log("keywords to pause: " + keywordsToAdd);
  //Logger.log("excluding this adgroup: " + adGroupsToAdd[kwGroup_i]);
  
  var keywords = AdWordsApp.keywords()
  .withCondition("Status = ENABLED")
  .withCondition('CampaignName = "'+campaignName+'"')
  .withCondition('AdGroupName != "'+adGroupsToAdd[kwGroup_i]+'"')
  .get();
  while(keywords.hasNext()){
   var keyword = keywords.next();
    
    if(keyword.getText().indexOf("+panel")>0){
    Logger.log(keyword.getText());
    }
    if(keywordsToAdd.indexOf(keyword.getText())>-1){
      //Logger.log("Keyword to pause (after adGroup check): " + keyword.getText());
      //Logger.log(keyword.getAdGroup().getName());      
      if(adGroupsToAdd[kwGroup_i].indexOf(keyword.getAdGroup().getName())==-1){
      //Logger.log("Keyword to pause, adgroup check done: " + keyword.getText());
      keyword.pause();
      }
    }
  }
 }
}

function createAds(){
  Logger.log("Creating ads...");
   var row = firstAdGroupRow;
  while(sheet.getRange(row, adGroupColumn).getValue()){
    
    var adGroupName = sheet.getRange(row, adGroupColumn).getValue(); 
    var url = sheet.getRange(row, urlColumn).getValue();  
    var headline1 = sheet.getRange(row, headline1Column).getValue();
    var headline2 = sheet.getRange(row, headline2Column).getValue();
    var path1 = sheet.getRange(row, displayUrl1Column).getValue();
    var path2 = sheet.getRange(row, displayUrl2Column).getValue();
    var description = sheet.getRange(row, descriptionColumn).getValue();
    var fullAd = url+headline1+headline2+path1+path2+description;
    
   // Logger.log("campaign name: " + campaignName);
    var adGroupIterator = AdWordsApp.adGroups()
    .withCondition('Name = "'+adGroupName+'"')
    .withCondition('CampaignName = "'+campaignName+'"').get();
    
    if (adGroupIterator.hasNext()) {
      var adGroup = adGroupIterator.next();
      
      var currentAds = [];
    var ads = adGroup.ads().withCondition("Status = ENABLED").withCondition("Type = EXPANDED_TEXT_AD").get();
    while(ads.hasNext()){
     var ad = ads.next();

      var fullCurrentAd = ad.urls().getFinalUrl() + ad.getHeadlinePart1() + ad.getHeadlinePart2() + ad.getPath1() + ad.getPath2() +ad.getDescription();
      currentAds.push(fullCurrentAd);
     //Logger.log(fullCurrentAd);

    }
    
    if(currentAds.indexOf(fullAd)>-1){
      //Logger.log("The ad already exists in the Ad Group so will not be created again");      
    }else{
    
    var result= adGroup.newAd().expandedTextAdBuilder()
        .withHeadlinePart1(headline1)
        .withHeadlinePart2(headline2)
        .withDescription(description)
        .withPath1(path1)
        .withPath2(path2)
        .withFinalUrl(url)
        .build().getResult()
    if(adLabel!=""){
      result.applyLabel(adLabel);
    }
  }
  }
    row++;
  }  
  
}

function createKeywords(){
var ranLabelCheck = false;
  if(createKeywords_bool == true){
    Logger.log("Creating keywords...");}
    var row = firstAdGroupRow;
    var currentKeywords = [];
  while(sheet.getRange(row, adGroupColumn).getValue()){
    var keywordsToAdd = [];
    var keywords = sheet.getRange(row, keywordsColumn).getValue();
    var keywords = keywords.split(",");
    var adGroupName = sheet.getRange(row, adGroupColumn).getValue(); 
    adGroupsToAdd.push(adGroupName);
    var url = sheet.getRange(row, urlColumn).getValue();
    var adGroupIterator = AdWordsApp.adGroups()
      .withCondition('Name = "'+adGroupName+'"')
     .withCondition('CampaignName = "'+campaignName+'"')
      .get();
    
  if (adGroupIterator.hasNext()) {
    var adGroup = adGroupIterator.next();
    //add negs
    var negSplit = sheet.getRange(row, negativeKeywordColumn).getValue().split(",");
    if(createKeywords_bool && negSplit!=""){
      for(var neg_i in negSplit){
        if(negSplit[neg_i]!=""){
          adGroup.createNegativeKeyword(negSplit[neg_i])
        }
      }
    }
    currentKeywords = [];
    var currentKeywordsGet = adGroup.keywords().withCondition("Status = ENABLED").get();
    while(currentKeywordsGet.hasNext()){
     currentKeywords.push(currentKeywordsGet.next().getText()); 
    }
    function addKeyword(keyword, url, bid){
      
      
      if(createKeywords_bool==true){
        //Logger.log("adding keyword " + keyword.toLowerCase())        
        if(url!=""){
         var result = adGroup.newKeywordBuilder()
        .withText(keyword.toLowerCase())
        .withCpc(bid)                         
        .withFinalUrl(url)
        .build().getResult();
          if(keywordLabel!=""){
         result.applyLabel(keywordLabel);
          }
        }else{
         var result = adGroup.newKeywordBuilder()
        .withText(keyword.toLowerCase())
        .withCpc(bid)                         
        .build()
         .getResult();
          if(keywordLabel!=""){
         result.applyLabel(keywordLabel);
          }

        }
    }
      keywordsToAdd.push(keyword.toLowerCase());
    }
    //Logger.log("currentKeywords: " +currentKeywords);
    for(var keyword in keywords){
      if(addUrlsToKeywords == false){
        url = "";
      }        
        function myTrim(x) {
          return x.replace(/^\s+|\s+$/gm,'');
        }
        var kwToAdd = myTrim(keywords[keyword]);
        if(exactBid != 0 && exactBid != "" && exactBid > 0){
          //check if the keyword already exists in the adGroup, add it if not
          var exactKW = '['+kwToAdd+']';
          if(currentKeywords.indexOf(exactKW)>-1){
           Logger.log("The keyword "+exactKW+" already exists in the Ad Group so will not be added or modified");
          }else{
           addKeyword(exactKW, url, exactBid);
          }
        }        
        if(phraseBid != 0 && phraseBid != "" && phraseBid > 0){
          var phraseKW = '"'+kwToAdd+'"';
            //check if the keyword already exists in the adGroup, add it if not
          if(currentKeywords.indexOf(phraseKW)>-1){
           Logger.log("The keyword "+phraseKW+" already exists in the Ad Group so will not be added or modified");
          }else{
           addKeyword(phraseKW, url, phraseBid);
          }
        }
        if(broadBid != 0 && broadBid != "" && broadBid > 0){
        var broadSplit = kwToAdd.split(" ");
          var kw = "";
          for(var b in broadSplit){

            if(b == broadSplit.length-1){
              kw += "+"+broadSplit[b]+"";
            }else{
              kw += "+"+broadSplit[b]+" ";
            }
          }
          //check if the keyword already exists in the adGroup, add it if not
          if(currentKeywords.indexOf(kw)>-1){
           Logger.log("The keyword "+kw+" already exists in the Ad Group so will not be added or modified");
          }else{
           addKeyword(kw, url, broadBid); 
          }
        }
        
      
       
    }

  }
    keywordGroups[row-9] = keywordsToAdd;
    
    row++;
  }  
  
}

function createAdGroups(){
 Logger.log("Creating adgroups...");
  
  var defaultBid = sheet.getRange(2, 4).getValue();
 // Logger.log("Campaign name: " + campaignName);
  var row = firstAdGroupRow;
  while(sheet.getRange(row, adGroupColumn).getValue()){
    var adGroupName = sheet.getRange(row, adGroupColumn).getValue();
    var campaignSelector = AdWordsApp
    .campaigns()
    .withCondition('Name = "' + campaignName +'"');
    var campaignIterator = campaignSelector.get();

 while (campaignIterator.hasNext()) {
   var campaign = campaignIterator.next();
      
   if(campaign.getName() != campaignName){
     continue;
   }
   
   if(exactBid > 0 && exactBid != ""){
     var agMaxCPC = exactBid;
   }else{
    var agMaxCPC = .5; 
   }
   
 var adGroupBuilder = campaign.newAdGroupBuilder();
 var adGroupOperation = adGroupBuilder
    .withName(adGroupName)
 .withCpc(agMaxCPC)
    .build();
 var adGroup = adGroupOperation.getResult();
   if(adGroupOperation.isSuccessful() && adGroupLabel!=""){
   adGroup.applyLabel(adGroupLabel)
   }

 }
    
    row++;
     
  }  
}

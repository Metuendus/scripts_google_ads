// v34 - product buckets & flexible date range version of PMax Insights script - (c) MikeRhodes.com.au 
// copy the template sheet & then grab YOUR url & paste that below in SHEET_URL
// template is: 
// https://docs.google.com/spreadsheets/d/1-Yx2TM0hFsy4mg29ErCnYlxHL37bUPOfkCAzqM8iBr0/copy




// user config section - make your changes here


  let SHEET_URL    = 'https://docs.google.com/spreadsheets/d/17J5V0HRhUW0Bv74qmsqSUPfjHzjLXkd3rwikjFPFbzY/edit#gid=223059630'    // enter your URL here
  let clientCode   = 'Coppel eComm'    // this string will be added to the sheet name 
  
// for all settings below, you can change the Settings tab in your sheet
// if those are blank, the default values below will be used
  let fixedStartDate = '2023-11-01' // CAMBIA ESTA FECHA USANDO EL FORMATO YYYY-MM-DD
  let fixedEndDate = '2023-11-31' // CAMBIA ESTA FECHA USANDO EL FORMATO YYYY-MM-DD
  let fixedProductStart = '2023-11-01'  // CAMBIA ESTA FECHA USANDO EL FORMATO YYYY-MM-DD
  let fixedProductEnd = '2023-11-31'  // CAMBIA ESTA FECHA USANDO EL FORMATO YYYY-MM-DD
  let tCost        = 10      // the 'threshold' for low vs high cost
  let tRoas        = 4       // the 'threshold' for low vs high roas (performance)   


//
// don’t change any code below this line ——————————————————————————
//
  
  
  
  


  
  
function main() {
    
// Setup Sheet
    let accountName = AdsApp.currentAccount().getName();
    let sheetNameIdentifier = clientCode ? clientCode : accountName;
    let ss = SpreadsheetApp.openByUrl(SHEET_URL);
    ss.rename(sheetNameIdentifier + ' - PMax Insights');
    updateVariablesFromSheet(ss, 'Settings');




// define query elements. wrap with spaces for safety
    let impr       = ' metrics.impressions ';
    let clicks     = ' metrics.clicks ';
    let cost       = ' metrics.cost_micros ';
    let engage     = ' metrics.engagements ';
    let inter      = ' metrics.interactions ';
    let conv       = ' metrics.conversions ';
    let value      = ' metrics.conversions_value ';
    let allConv    = ' metrics.all_conversions ';
    let allValue   = ' metrics.all_conversions_value ';
    let views      = ' metrics.video_views ';
    let cpv        = ' metrics.average_cpv ';
    let segDate    = ' segments.date ';
    let prodTitle  = ' segments.product_title ';
    let prodID     = ' segments.product_item_id ';
    let campName   = ' campaign.name ';
    let campId     = ' campaign.id ';
    let chType     = ' campaign.advertising_channel_type ';
    let aIdAsset   = ' asset.resource_name ';
    let aId        = ' asset.id ';
    let assetSource= ' asset.source ';
    let agId       = ' asset_group.id ';
    let assetFtype = ' asset_group_asset.field_type ';
    let adPmaxPerf = ' asset_group_asset.performance_label ';
    let agStrength = ' asset_group.ad_strength ';
    let agStatus   = ' asset_group.status ';
    let asgName    = ' asset_group.name ';
    let lgType     = ' asset_group_listing_group_filter.type ';
    let aIdCamp    = ' segments.asset_interaction_target.asset ';    
    let assetName  = ' asset.name ';
    let adUrl      = ' asset.image_asset.full_size.url ';
    let ytTitle    = ' asset.youtube_video_asset.youtube_video_title ';
    let ytId       = ' asset.youtube_video_asset.youtube_video_id '; 
    let interAsset = ' segments.asset_interaction_target.interaction_on_this_asset ';   
    let pMaxOnly   = ' AND campaign.advertising_channel_type = "PERFORMANCE_MAX" '; 
    let agFilter   = ' AND asset_group_listing_group_filter.type != "SUBDIVISION" ';
    let notInter   = ' AND segments.asset_interaction_target.interaction_on_this_asset != "TRUE" ';
    let order      = ' ORDER BY campaign.name ';


    
// Date stuff --------------------------------------------
let timeZone = AdsApp.currentAccount().getTimeZone();
// Fixed start and end dates
let startDate = new Date(fixedStartDate);
let endDate = new Date(fixedEndDate);
let productStart = new Date(fixedProductStart);
let productEnd = new Date(fixedProductEnd);

// Format Dates
let formattedStart = Utilities.formatDate(startDate, timeZone, 'yyyy-MM-dd');
let formattedEnd = Utilities.formatDate(endDate, timeZone, 'yyyy-MM-dd');
let formattedProduct   = Utilities.formatDate(productStart, timeZone, 'yyyy-MM-dd');

// SQL Date Range
let mainDateRange = ` segments.date BETWEEN "${formattedStart}" AND "${formattedEnd}" `;
let prodDateRange = ` segments.date BETWEEN "${formattedProduct}" AND "${productEnd}" `;


    
// Build queries --------------------------------------------
    let assetGroupAssetColumns = [campName, asgName, agId, aIdAsset, assetFtype, campId]; //  adPmaxPerf, agStrength, agStatus, assetSource,
    let assetGroupAssetQuery = 'SELECT ' + assetGroupAssetColumns.join(',') +
      ' FROM asset_group_asset ' + 
      ' WHERE campaign.status != REMOVED ' + pMaxOnly ;


    let displayVideoColumns = [segDate, campName, aIdCamp, cost, conv, value, views, cpv, impr, clicks, chType, interAsset, campId];
    let displayVideoQuery = 'SELECT ' + displayVideoColumns.join(',') +
      ' FROM campaign ' +
      ' WHERE ' + mainDateRange + pMaxOnly + notInter + order;


    let assetGroupColumns = [segDate, campName, asgName, agStrength, agStatus, lgType, impr, clicks, cost, conv, value]; 
    let assetGroupQuery = 'SELECT ' + assetGroupColumns.join(',')  + 
      ' FROM asset_group_product_group_view ' +
      ' WHERE ' + mainDateRange + agFilter ;


    let campaignColumns = [segDate, campName, cost, conv, value, views, cpv, impr, clicks, chType, campId]; 
    let campaignQuery = 'SELECT ' + campaignColumns.join(',') + 
      ' FROM campaign ' +
      ' WHERE ' + mainDateRange + pMaxOnly  + order;


    let productColumns = [prodTitle, prodID, cost, conv, value, impr, clicks, chType]; 
    let productQuery = 'SELECT ' + productColumns.join(',')  + 
      ' FROM shopping_performance_view  ' + 
      ' WHERE metrics.impressions > 0 AND ' + prodDateRange + pMaxOnly ; 


    let assetColumns = [aIdAsset, assetSource, ytTitle, ytId, assetName] 
    let assetQuery = 'SELECT ' + assetColumns.join(',')  + 
      ' FROM asset ' ;


    


    
// Process data --------------------------------------------    


    let assetGroupAssetData = fetchData(assetGroupAssetQuery);
    let displayVideoData    = fetchData(displayVideoQuery);
    let assetGroupData      = fetchData(assetGroupQuery);
    let campaignData        = fetchData(campaignQuery);
    let assetData           = fetchData(assetQuery);
    let productData         = fetchProductData(productQuery);   


    
// Extract marketing assets & de-dupe
    let { displayAssets, videoAssets } = extractAndDeduplicateAssets(assetGroupAssetData);


    
// Filter displayVideoData based on the type
    let displayAssetData = filterDataByAssets(displayVideoData, displayAssets);
    let videoAssetData   = filterDataByAssets(displayVideoData, videoAssets);


    
// Process data for each dataset
    let processedDisplayAssetData = aggregateDataByDateAndCampaign(displayAssetData);
    let processedVideoAssetData   = aggregateDataByDateAndCampaign(videoAssetData);  
    let processedAssetGroupData   = processData(assetGroupData);
    let processedCampData         = processData(campaignData);
  
// Combine all non-search metrics, calc 'search' & process summary
    let nonSearchData = [...processedDisplayAssetData, ...processedVideoAssetData, ...processedAssetGroupData];
    let searchResults = getSearchResults(processedCampData, nonSearchData);
    let totalData     = processTotalData(processedCampData, processedAssetGroupData, processedDisplayAssetData, processedVideoAssetData, searchResults);
    let summaryData   = processSummaryData(processedCampData, processedAssetGroupData, processedDisplayAssetData, processedVideoAssetData, searchResults);    




// Aggregate the metrics for display and video assets
    let aggregatedDisplayAssetMetrics = aggregateMetricsByAsset(displayAssetData);
    let aggregatedVideoAssetMetrics   = aggregateMetricsByAsset(videoAssetData);
    let enrichedDisplayAssetDetails   = enrichAssetMetrics(aggregatedDisplayAssetMetrics, assetData, 'display');
    let enrichedVideoAssetDetails     = enrichAssetMetrics(aggregatedVideoAssetMetrics, assetData, 'video');
    let mergedDisplayData             = mergeMetricsWithDetails(aggregatedDisplayAssetMetrics, enrichedDisplayAssetDetails);
    let mergedVideoData               = mergeMetricsWithDetails(aggregatedVideoAssetMetrics, enrichedVideoAssetDetails);


    
// Output the data to respective sheets
    outputAggregatedDataToSheet(ss, 'display',  mergedDisplayData);
    outputAggregatedDataToSheet(ss, 'video',    mergedVideoData);
    outputSummaryToSheet(       ss, 'totals',   totalData);
    outputSummaryToSheet(       ss, 'summary',  summaryData);
    outputDataToSheet(          ss, 'products', productData);
    outputDataToSheet(          ss, 'group',    assetGroupData,[0,8,12,13,14]);
    
    
// get search terms for main date range    
    extractTerms(ss, mainDateRange)
    
    
} // end main








// various additional functions ---------------------------------------------








// fetch data given a query string using search (not report)
  function fetchData(queryString) {
    let data = [];
    const iterator = AdsApp.search(queryString);
    while (iterator.hasNext()) {
      const row = iterator.next();
      const rowData = flattenObject(row); // Flatten the row data
      data.push(rowData);
    }
    return data;
  }








// flatten the object data to enable more processing of that data
  function flattenObject(ob) {
    var toReturn = {};
    for (var i in ob) {
      if ((typeof ob[i]) === 'object') {
        var flatObject = flattenObject(ob[i]);
        for (var x in flatObject) {
          toReturn[i + '.' + x] = flatObject[x];
        }
      } else {
        toReturn[i] = ob[i];
      }
    }
    return toReturn;
  }








// extract search terms
  function extractTerms(ss, mainDateRange) {
      // Extract Campaign IDs with Status not 'REMOVED' and impressions > 0
      let campaignIdsQuery = AdsApp.report(`SELECT campaign.id, campaign.name, metrics.clicks, metrics.impressions, metrics.conversions, metrics.conversions_value
          FROM campaign WHERE campaign.status != 'REMOVED' AND metrics.impressions > 0 AND ${mainDateRange} ORDER BY metrics.conversions DESC `
      );


      let rows = campaignIdsQuery.rows();   
      let allSearchTerms = [['Campaign Name', 'Campaign ID', 'Category Label', 'Clicks', 'Impressions', 'Conversions', 'Conversion Value']];


      while (rows.hasNext()) {
          let row = rows.next();


          // Fetch search terms and related metrics for the campaign ordered by conversions descending
          let query = AdsApp.report(` SELECT campaign_search_term_insight.category_label, metrics.clicks, metrics.impressions, 
              metrics.conversions, metrics.conversions_value  FROM campaign_search_term_insight WHERE ${mainDateRange}
              AND campaign_search_term_insight.campaign_id = ${row['campaign.id']} ORDER BY metrics.impressions DESC `
          );


          let searchTermRows = query.rows();
          while (searchTermRows.hasNext()) {
              let searchTermRow = searchTermRows.next();
              allSearchTerms.push([row['campaign.name'], row['campaign.id'], 
                                   searchTermRow['campaign_search_term_insight.category_label'],
                                   searchTermRow['metrics.clicks'], 
                                   searchTermRow['metrics.impressions'], 
                                   searchTermRow['metrics.conversions'], 
                                   searchTermRow['metrics.conversions_value']]);
          }
      }


      // Write all search terms to the 'terms' sheet 
      let termsSheet = ss.getSheetByName('terms') ? ss.getSheetByName('terms').clear() : ss.insertSheet('terms');
      termsSheet.getRange(1, 1, allSearchTerms.length, allSearchTerms[0].length).setValues(allSearchTerms);




      // Aggregate terms and write to the 'totalTerms' sheet
      aggregateTerms(allSearchTerms, ss);
  }




// aggregate search terms
  function aggregateTerms(searchTerms, ss) {
      let aggregated = {}; // { term: { clicks: 0, impressions: 0, conversions: 0, conversionValue: 0 }, ... }


      for (let i = 1; i < searchTerms.length; i++) { // Start from 1 to skip headers
          let term = searchTerms[i][2] || 'blank'; // Use 'blank' for empty search terms


          if (!aggregated[term]) {
              aggregated[term] = { clicks: 0, impressions: 0, conversions: 0, conversionValue: 0 };
          }


          aggregated[term].clicks += Number(searchTerms[i][3]);
          aggregated[term].impressions += Number(searchTerms[i][4]);
          aggregated[term].conversions += Number(searchTerms[i][5]);
          aggregated[term].conversionValue += Number(searchTerms[i][6]);
      }


      let aggregatedArray = [['Category Label', 'Clicks', 'Impressions', 'Conversions', 'Conversion Value']];
      for (let term in aggregated) {
          aggregatedArray.push([term, aggregated[term].clicks, aggregated[term].impressions, aggregated[term].conversions, aggregated[term].conversionValue]);
      }


      let header = aggregatedArray.shift(); // Remove the header before sorting
      // Sort by impressions descending
      aggregatedArray.sort((a, b) => b[2] - a[2]);
      aggregatedArray.unshift(header); // Prepend the header back to the top


      // Write aggregated data to the 'totalTerms' sheet
      let totalTermsSheet = ss.getSheetByName('totalTerms') ? ss.getSheetByName('totalTerms').clear() : ss.insertSheet('totalTerms');
      totalTermsSheet.getRange(1, 1, aggregatedArray.length, aggregatedArray[0].length).setValues(aggregatedArray);
  }




// extract & dedupe in one step
  function extractAndDeduplicateAssets(data) {
      let displayAssetsSet = new Set();
      let videoAssetsSet = new Set();


      data.forEach(row => {
          if (row['assetGroupAsset.fieldType'] && row['assetGroupAsset.fieldType'].includes('MARKETING')) {
              displayAssetsSet.add(row['asset.resourceName']);
          }
          if (row['assetGroupAsset.fieldType'] && row['assetGroupAsset.fieldType'].includes('VIDEO')) {
              videoAssetsSet.add(row['asset.resourceName']);
          }
      });


      return {
          displayAssets: [...displayAssetsSet],
          videoAssets: [...videoAssetsSet]
      };
  }




// filter the display & video assets
  function filterDataByAssets(data, assets) {
      return data.filter(row => assets.includes(row['segments.assetInteractionTarget.asset']));
  }




// get additional data from asset for video 
  function mergeMetricsWithDetails(aggregatedVideoAssetMetrics, enrichedVideoAssetDetails) {
      return enrichedVideoAssetDetails.map(detail => {
          const metrics = aggregatedVideoAssetMetrics[detail.assetName];
          return {
              ...detail, 
              ...metrics
          };
      });
  }




// enrich assets
  function enrichAssetMetrics(aggregatedMetrics, assetData, type) {
      let assetDetailsArray = [];


      // For each asset in aggregatedMetrics, fetch details from assetData
      for (let assetName of Object.keys(aggregatedMetrics)) {
          // Find the asset in assetData
          let matchingAsset = assetData.find(asset => asset['asset.resourceName'] === assetName);


          if (matchingAsset) {
              let assetDetails = {
                  assetName: assetName,
                  assetSource: matchingAsset['asset.source'],
                  assetImage: matchingAsset['asset.name'] 
              };


              if (type === 'video') {
                  assetDetails.youtubeTitle = matchingAsset['asset.youtubeVideoAsset.youtubeVideoTitle'];
                  assetDetails.youtubeId = matchingAsset['asset.youtubeVideoAsset.youtubeVideoId'];
              }


              assetDetailsArray.push(assetDetails);
          }
      }


      return assetDetailsArray;
  }




// Aggregate display & video assets to show total metrics for each asset in their sheets
  function aggregateMetricsByAsset(data) {
      const aggregatedData = {};


      data.forEach(row => {
          const asset = row['segments.assetInteractionTarget.asset'];


          if (!aggregatedData[asset]) {
              aggregatedData[asset] = {
                  clicks: 0,
                  videoViews: 0,
                  conversionsValue: 0,
                  conversions: 0,
                  cost: 0,
                  impressions: 0
              };
          }


          aggregatedData[asset].clicks += parseInt(row['metrics.clicks']);
          aggregatedData[asset].videoViews += parseInt(row['metrics.videoViews']);
          aggregatedData[asset].conversionsValue += parseFloat(row['metrics.conversionsValue']);
          aggregatedData[asset].conversions += parseInt(row['metrics.conversions']);
          aggregatedData[asset].cost += parseInt(row['metrics.costMicros']) / 1e6; 
          aggregatedData[asset].impressions += parseInt(row['metrics.impressions']);
      });


      return aggregatedData;
  }




// function to output the display & video metrics to tabs
  function outputAggregatedDataToSheet(ss, sheetName, data) {
      // Create or access the sheet
      let sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);


      // Clear the sheet
      sheet.clear();


      // Set the header based on sheetName
      const headers = sheetName === 'video' 
          ? ['Asset Name', 'Source', 'YouTube Title', 'YouTube ID', 'Impr', 'Clicks', 'Views', 'Value', 'Conv', 'Cost']
          : ['Asset Name', 'Source', 'ImageName', 'Impr', 'Clicks', 'Views', 'Value', 'Conv', 'Cost'];
      sheet.appendRow(headers);


      // Append the aggregated data
      data.forEach(item => {
          const rowData = sheetName === 'video'
              ? [item.assetName, item.assetSource, item.youtubeTitle, item.youtubeId, item.impressions, item.clicks, item.videoViews, item.conversionsValue, item.conversions, item.cost]
              : [item.assetName, item.assetSource, item.assetImage, item.impressions, item.clicks, item.videoViews, item.conversionsValue, item.conversions, item.cost];
          sheet.appendRow(rowData);
      });
  }




// Process data: Aggregate by date, campaign name, and various metrics
  function aggregateDataByDateAndCampaign(data) {
      const aggregatedData = {};


      data.forEach(row => {
          const date = row['segments.date'];
          const campaignName = row['campaign.name'];
          const key = `${date}_${campaignName}`;


          if (!aggregatedData[key]) {
              aggregatedData[key] = {
                  'date': date,
                  'campaignName': campaignName,
                  'cost': 0,
                  'impressions': 0,
                  'clicks': 0,
                  'conversions': 0,
                  'conversionsValue': 0
              };
          }


          if (row['metrics.costMicros']) {
              aggregatedData[key].cost += row['metrics.costMicros'] / 1000000;
          }
          if (row['metrics.impressions']) {
              aggregatedData[key].impressions += parseInt(row['metrics.impressions']);
          }
          if (row['metrics.clicks']) {
              aggregatedData[key].clicks += parseInt(row['metrics.clicks']);
          }
          if (row['metrics.conversions']) {
              aggregatedData[key].conversions += parseFloat(row['metrics.conversions']);
          }
          if (row['metrics.conversionsValue']) {
              aggregatedData[key].conversionsValue += parseFloat(row['metrics.conversionsValue']);
          }
      });


      return Object.values(aggregatedData);
  }






// calc the search results
  function getSearchResults(processedCampData, nonSearchData) {
    // Initialize an object to hold 'search' metrics
    const searchMetrics = {};


    // Sum up the non-search metrics
    nonSearchData.forEach(row => {
      const key = `${row.date}_${row.campaignName}`;
      if (!searchMetrics[key]) {
        searchMetrics[key] = {
          date: row.date,
          campaignName: row.campaignName,
          cost: 0,
          impressions: 0,
          clicks: 0,
          conversions: 0,
          conversionsValue: 0
        };
      }
      for (let metric of ['cost', 'impressions', 'clicks', 'conversions', 'conversionsValue']) {
        searchMetrics[key][metric] += row[metric];
      }
    });


    // Calculate 'search' metrics
    const searchResults = [];
    processedCampData.forEach(row => {
      const rowCopy = JSON.parse(JSON.stringify(row));  // Deep copy
      const key = `${row.date}_${row.campaignName}`;
      if (searchMetrics[key]) {
        for (let metric of ['cost', 'impressions', 'clicks', 'conversions', 'conversionsValue']) {
          rowCopy[metric] -= (searchMetrics[key][metric] || 0);
        }
      }
      searchResults.push(rowCopy);
    });


    return searchResults;
  }




// process the data
  function processData(data) {
    const summedData = {};
    data.forEach(row => {
      const date = row['segments.date'];
      const campaignName = row['campaign.name'];
      const key = `${date}_${campaignName}`;
      // Initialize if the key doesn't exist
      if (!summedData[key]) {
        summedData[key] = {
          'date': date,
          'campaignName': campaignName,
          'cost': 0,
          'impressions': 0,
          'clicks': 0,
          'conversions': 0,
          'conversionsValue': 0
        };
      }
      if (row['metrics.costMicros']) {
        summedData[key].cost += row['metrics.costMicros'] / 1000000;
      }
      if (row['metrics.impressions']) {
        summedData[key].impressions += parseInt(row['metrics.impressions']);
      }
      if (row['metrics.clicks']) {
        summedData[key].clicks += parseInt(row['metrics.clicks']);
      }
      if (row['metrics.conversions']) {
        summedData[key].conversions += parseFloat(row['metrics.conversions']);
      }
      if (row['metrics.conversionsValue']) {
        summedData[key].conversionsValue += parseFloat(row['metrics.conversionsValue']);
      }
    });


    return Object.values(summedData);
  }




// output data to tab in sheet
function outputDataToSheet(spreadsheet, sheetName, data, indexesToRemove) {
    let sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
        sheet = spreadsheet.insertSheet(sheetName);
    }
    sheet.clearContents();


    // Check if data is undefined or empty, and if so, simply return after clearing the sheet
    if (!data || data.length === 0) {
        return; // Exit the function here; the sheet has been cleared, and there's no data to process
    }
  
    // If indexesToRemove is provided and it's an array, filter the data
    if (Array.isArray(indexesToRemove) && indexesToRemove.length > 0) {
        data = data.map(item => {
            if (typeof item === 'object') { // Ensure that we are dealing with an object
                let keys = Object.keys(item);
                return keys.reduce((obj, key, index) => {
                    if (!indexesToRemove.includes(index)) {
                        obj[key] = item[key]; // keep this field in the new object
                    }
                    return obj;
                }, {});
            }
            return item; // If not an object, return the item as is
        });
    }


    // Check if data is an array of objects or a simple array
    if (data.length > 0 && typeof data[0] === 'object') {
        const header = [Object.keys(data[0])];
        const rows = data.map(row => Object.values(row));
        const allRows = header.concat(rows);
        sheet.getRange(1, 1, allRows.length, allRows[0].length).setValues(allRows);
    } else {
        // Handle the case for a simple array like marketingAssets
        const rows = data.map(item => [item]);
        sheet.getRange(1, 1, rows.length, 1).setValues(rows);
    }
}




// output summary data
  function outputSummaryToSheet(ss, sheetName, summaryData) {
    // Get the sheet with the name `sheetName` from `ss`. Create it if it doesn't exist.
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    sheet.clear();
    for (let i = 0; i < summaryData.length; i++) {
      sheet.getRange(i + 1, 1, 1, summaryData[i].length).setValues([summaryData[i]]);
    }
    sheet.setFrozenRows(1);
    sheet.getRange("1:1").setFontWeight("bold");
    sheet.getRange("1:1").setWrap(true);
    sheet.setRowHeight(1, 40);
  }




// create summary on one tab
  function processSummaryData(processedCampData, processedAssetGroupData, processedDisplayData, processedDataForVideo, searchResults) {
    const header = ['Date', 'Campaign Name', 
                    'Camp Cost', 'Camp Conv', 'Camp Value',
                    'Shop Cost', 'Shop Conv', 'Shop Value',
                    'Disp Cost', 'Disp Conv', 'Disp Value',
                    'Video Cost', 'Video Conv', 'Video Value',
                    'Search Cost', 'Search Conv', 'Search Value'];


    const summaryData = {};


    // Helper function to update summaryData
    function updateSummary(row, type) {
      const date = row['date'];
      const campaignName = row['campaignName'];
      const key = `${date}_${campaignName}`;


      // Initialize if the key doesn't exist
      if (!summaryData[key]) {
        summaryData[key] = {
          'date': date,
          'campaignName': campaignName,
          'generalCost': 0, 'generalConv': 0, 'generalConvValue': 0,
          'shoppingCost': 0, 'shoppingConv': 0, 'shoppingConvValue': 0,
          'displayCost': 0, 'displayConv': 0, 'displayConvValue': 0,
          'videoCost': 0, 'videoConv': 0, 'videoConvValue': 0,
          'searchCost': 0, 'searchConv': 0, 'searchConvValue': 0
        };
      }


      // Update the metrics
      summaryData[key][`${type}Cost`] += row['cost'];
      summaryData[key][`${type}Conv`] += row['conversions'];
      summaryData[key][`${type}ConvValue`] += row['conversionsValue'];
    }


    // Process each lot of data
    processedCampData.forEach(row => updateSummary(row, 'general'));
    processedAssetGroupData.forEach(row => updateSummary(row, 'shopping'));
    processedDisplayData.forEach(row => updateSummary(row, 'display'));
    processedDataForVideo.forEach(row => updateSummary(row, 'video'));
    searchResults.forEach(row => updateSummary(row, 'search'));


    // Convert summaryData to an array format
    const summaryArray = Object.values(summaryData).map(summaryRow => {
      return [
        summaryRow.date,
        summaryRow.campaignName,
        summaryRow.generalCost, summaryRow.generalConv, summaryRow.generalConvValue,
        summaryRow.shoppingCost, summaryRow.shoppingConv, summaryRow.shoppingConvValue,
        summaryRow.displayCost, summaryRow.displayConv, summaryRow.displayConvValue,
        summaryRow.videoCost, summaryRow.videoConv, summaryRow.videoConvValue,
        summaryRow.searchCost, summaryRow.searchConv, summaryRow.searchConvValue
      ];
    });


    return [header, ...summaryArray];
  }




// Process total data for 'Totals' tab
  function processTotalData(processedCampData, processedAssetGroupData, processedDisplayData, processedDataForVideo, searchResults) {
    const header = ['Campaign Name', 
                    'Camp Cost', 'Camp Conv', 'Camp Value',
                    'Shop Cost', 'Shop Conv', 'Shop Value',
                    'Disp Cost', 'Disp Conv', 'Disp Value',
                    'Video Cost', 'Video Conv', 'Video Value',
                    'Search Cost', 'Search Conv', 'Search Value'];


    const totalData = {};


    // Helper function to update totalData
    function updateTotal(row, type) {
      const campaignName = row['campaignName'];
      const key = campaignName;


      // Initialize if the key doesn't exist
      if (!totalData[key]) {
        totalData[key] = {
          'campaignName': campaignName,
          'generalCost': 0, 'generalConv': 0, 'generalConvValue': 0,
          'shoppingCost': 0, 'shoppingConv': 0, 'shoppingConvValue': 0,
          'displayCost': 0, 'displayConv': 0, 'displayConvValue': 0,
          'videoCost': 0, 'videoConv': 0, 'videoConvValue': 0,
          'searchCost': 0, 'searchConv': 0, 'searchConvValue': 0
        };
      }


      // Update the metrics
      totalData[key][`${type}Cost`] += row['cost'];
      totalData[key][`${type}Conv`] += row['conversions'];
      totalData[key][`${type}ConvValue`] += row['conversionsValue'];
    }


    // Process each lot of data
    processedCampData.forEach(row => updateTotal(row, 'general'));
    processedAssetGroupData.forEach(row => updateTotal(row, 'shopping'));
    processedDisplayData.forEach(row => updateTotal(row, 'display'));
    processedDataForVideo.forEach(row => updateTotal(row, 'video'));
    searchResults.forEach(row => updateTotal(row, 'search'));


    // Convert totalData to an array format
    const totalArray = Object.values(totalData).map(totalRow => {
      return [
        totalRow.campaignName,
        totalRow.generalCost, totalRow.generalConv, totalRow.generalConvValue,
        totalRow.shoppingCost, totalRow.shoppingConv, totalRow.shoppingConvValue,
        totalRow.displayCost, totalRow.displayConv, totalRow.displayConvValue,
        totalRow.videoCost, totalRow.videoConv, totalRow.videoConvValue,
        totalRow.searchCost, totalRow.searchConv, totalRow.searchConvValue
      ];
    });


    return [header, ...totalArray];
  }






// Specialized fetch data function for product data
function fetchProductData(queryString) {
  let data = [];
  let aggregatedData = {};
  const iterator = AdsApp.search(queryString);


  // Threshold variables are defined globally


  while (iterator.hasNext()) {
    const row = iterator.next();
    let rowData = flattenObject(row); // Flatten the row data
    
    // Remove unwanted fields
    delete rowData['campaign.resourceName'];
    delete rowData['campaign.advertisingChannelType'];
    delete rowData['shoppingPerformanceView.resourceName'];


    // Generate a unique key using only 'segments.productTitle'
    const uniqueKey = rowData['segments.productTitle']; // Use product title as the unique key


    // Aggregate data
    if (aggregatedData.hasOwnProperty(uniqueKey)) {
      // Aggregate specific fields here
      aggregatedData[uniqueKey]['Impr'] += Number(rowData['metrics.impressions']) || 0;
      aggregatedData[uniqueKey]['Clicks'] += Number(rowData['metrics.clicks']) || 0;
      aggregatedData[uniqueKey]['Cost'] += (Number(rowData['metrics.costMicros']) / 1e6) || 0; // Convert to standard currency 
      aggregatedData[uniqueKey]['Conv'] += Number(rowData['metrics.conversions']) || 0;
      aggregatedData[uniqueKey]['Value'] += Number(rowData['metrics.conversionsValue']) || 0;
    } else {
      // If this is the first entry for this uniqueKey, initialize the values
      aggregatedData[uniqueKey] = {
        'Product Title': rowData['segments.productTitle'],
        'Impr': Number(rowData['metrics.impressions']) || 0,
        'Clicks': Number(rowData['metrics.clicks']) || 0,
        'Cost': (Number(rowData['metrics.costMicros']) / 1e6) || 0,
        'Conv': Number(rowData['metrics.conversions']) || 0,
        'Value': Number(rowData['metrics.conversionsValue']) || 0,
        'CTR': 0,
        'ROAS': 0,
        'CvR': 0
      };
    }
  }


 // Post-processing for additional fields and calculations
  for (let key in aggregatedData) {
    // Calculate CTR - thanks to MikeRhodes.com.au member Pedro Matute!
    let ctr = aggregatedData[key]['Clicks'] / aggregatedData[key]['Impr'];
    aggregatedData[key]['CTR'] = isNaN(ctr) ? 0 : ctr;  


    // Calculate ROAS and add it as a new field
    if (aggregatedData[key]['Cost'] > 0) { // Now 'Cost' is used
      let roas = aggregatedData[key]['Value'] / aggregatedData[key]['Cost'];
      aggregatedData[key]['ROAS'] = roas;
    } else {
      aggregatedData[key]['ROAS'] = 0; // Default to 0 if cost is 0 to avoid division by zero
    }


    // Calculate Conversion Rate
    let cvr = aggregatedData[key]['Conv'] / aggregatedData[key]['Clicks'];


    aggregatedData[key]['CvR'] = isNaN(cvr) ? 0 : cvr; // Check for NaN and replace with 0 if necessary


    // Calculate and assign buckets based on cost and ROAS
    if (aggregatedData[key]['Cost'] === 0) {
      aggregatedData[key]['Bucket'] = 'zombie';
    } else if (aggregatedData[key]['Conv'] === 0) {
      aggregatedData[key]['Bucket'] = 'zeroconv';
    } else {
      if (aggregatedData[key]['Cost'] < tCost) {
        if (aggregatedData[key]['ROAS'] < tRoas) {
          aggregatedData[key]['Bucket'] = 'meh';
        } else {
          aggregatedData[key]['Bucket'] = 'flukes';
        }
      } else {
        if (aggregatedData[key]['ROAS'] < tRoas) {
          aggregatedData[key]['Bucket'] = 'costly';
        } else {
          aggregatedData[key]['Bucket'] = 'profitable';
        }
      }
    }
  }


  // Convert aggregated data back to array format with the new key names, order, and 'Bucket' column
  for (let key in aggregatedData) {
    data.push(aggregatedData[key]);
  }


  return data;
}




// update global variables using values form sheet (if entered)
function updateVariablesFromSheet(ss, sheetName) {
  // Define the names of your named ranges
  const numberOfDaysRangeName = 'numberOfDays';
  const tCostRangeName = 'tCost';
  const tRoasRangeName = 'tRoas';
  const productDaysRangeName = 'productDays';


  try {
    // Access the specific sheet within the spreadsheet
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return;
    }


    // Fetch the ranges by name from the specific sheet
    const numberOfDaysRange = sheet.getRange(numberOfDaysRangeName);
    const tCostRange = sheet.getRange(tCostRangeName);
    const tRoasRange = sheet.getRange(tRoasRangeName);
    const productDaysRange = sheet.getRange(productDaysRangeName);


    // Check if the range is found and the value is not empty
    // Then update the global variables if the values are numbers
    if (numberOfDaysRange && numberOfDaysRange.getValue() !== "") {
      const value = numberOfDaysRange.getValue();
      if (!isNaN(value)) {
        numberOfDays = value;
      }
    }


    if (tCostRange && tCostRange.getValue() !== "") {
      const value = tCostRange.getValue();
      if (!isNaN(value)) {
        tCost = value;
      }
    }


    if (tRoasRange && tRoasRange.getValue() !== "") {
      const value = tRoasRange.getValue();
      if (!isNaN(value)) {
        tRoas = value;
      }
    }


    if (productDaysRange && productDaysRange.getValue() !== "") {
      const value = productDaysRange.getValue();
      if (!isNaN(value)) {
        productDays = value;
      }
    }
  } catch (e) {
    // If there's an error, we simply don't update the variables
    // Log the error for debugging purposes
    console.error("Error in updateVariablesFromSheet:", e);
  }
}
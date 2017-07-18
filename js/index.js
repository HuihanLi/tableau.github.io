//first, create the viz
title_sheet = {
        "WEEKLY TV IMPRESSIONS": "WEEKLY TV/HULU IMPRESSIONS",
        "WEEKLY PRINT CIRCULATION": "MONTHLY PRINT CIRCULATION",
        "WEEKLY RADIO IMPRESSIONS": "WEEKLY RADIO IMPRESSIONS",
        "FACEBOOK IMPRESSIONS": "WEEKLY FACEBOOK IMPRESSIONS",
        "FACEBOOK ENGAGEMENT": "WEEKLY FACEBOOK ENGAGEMENT",
        "PREROLL VIDEO IMPRESSIONS": "WEEKLY PREROLL VIDEO IMPRESSIONS",
        "PREROLL VIDEO CTR": "WEEKLY PREROLL VIDEO CTR",
        "NATIVE IMPRESSIONS": "WEEKLY NATIVE IMPRESSIONS",
        "NATIVE CTR": "WEEKLY NATIVE CTR",
        "FACEBOOK SENTIMENT": "WEEKLY FACEBOOK SENTIMENT",
        "SITE VISIT": "WEEKLY SITE VISIT"
     };

col_alias = {
  "[excel-direct.1dbzcdg1ts2uwx16oonsx1bx6wsb].[usr:Calculation_284008305897496576:qk]": "Hulu",
  "[excel-direct.1dbzcdg1ts2uwx16oonsx1bx6wsb].[sum:W25-54 Impressions:qk]": "TV",
  "[textscan.0odnxxs0zp1l2n18cnved0wpyswx].[sum:page_impressions_organic:qk]": "Page_Impressions_Organic",
  "[textscan.0odnxxs0zp1l2n18cnved0wpyswx].[sum:page_impressions_paid:qk]": "Page_Impressions_Paid",
  "[federated.0omo11h0k9xghi10nnzcl02angis].[usr:Calculation_517632548688437249:qk]": "Weekly Negative",
  "[federated.0omo11h0k9xghi10nnzcl02angis].[usr:Calculation_635289030116114432:qk]": "Weekly Neutral",
  "[federated.0omo11h0k9xghi10nnzcl02angis].[usr:Calculation_517632548688109568:qk]": "Weekly Positive",
  "[textscan.1lwe0ud158tikj17tyowb0jmmhem].[sum:pageviews:qk]": "Pageviews",
  "[textscan.1lwe0ud158tikj17tyowb0jmmhem].[sum:uniquePageviews:qk]": "uniquePageviews",
  "[textscan.0odnxxs0zp1l2n18cnved0wpyswx].[sum:page_positive_feedback_link:qk]": "Share",
  "[textscan.0odnxxs0zp1l2n18cnved0wpyswx].[sum:page_positive_feedback_comment:qk]": "Comment",
  "[textscan.0odnxxs0zp1l2n18cnved0wpyswx].[sum:page_positive_feedback_like:qk]": "Like"

};

$(document).ready(function initViz() {

  var vizDiv = document.getElementById("vizDiv"),
    url = "http://analytics.bbdo.com/t/SandersonFarms/views/SandersonFarmsSummary/SUMMARY",

    options = {
      width: '100%',
      height: '100%',
      hideToolbar: false,
      onFirstInteractive: function() {

        workbook = viz.getWorkbook();
        dash = viz.getWorkbook().getActiveSheet();
        workbook.activateSheetAsync(dash)
          .then(function(dashboard) {
            //this is the secret sauce that stores all of the sheet names on the dashboard in an array named sheetNames
            var worksheets = dashboard.getWorksheets();
            //here is the aforementioned array
            var sheetNames = [];
            //here we are looping through all the sheets on the dash, & pushing the sheet name to the sheetNames array
            for (var i = 0, len = worksheets.length; i < len; i++) {

              var sheet = worksheets[i];
              if(!sheet.getName().includes("Total")){
                sheetNames.push(title_sheet[sheet.getName()]);
                //sheetNames.push(sheet.getName());
                console.log(sheet.getName());

              }
            }

            //here is where we are creating our dropdown menu... kind of like a Tableau parameter
            //we will inject this value into the getVizData() function a little bit later
            var sel = document.getElementById('SheetList');
            var fragment = document.createDocumentFragment();
            sheetNames.forEach(function(sheetName, index) {

              var opt = document.createElement('option');
              opt.innerHTML = sheetName;
              opt.value = sheetName;
              
              fragment.appendChild(opt);
            });

            sel.appendChild(fragment);

          });
      }
    };
  viz = new tableau.Viz(vizDiv, url, options);
});

function conn(){
  window.open("https://analytics.bbdo.com");
}
function getKeyByValue(object, value) {
  return Object.keys(object).find(key => object[key] === value);
}

function getVizData(bnum) {

  button_number = bnum;
  options = {

    maxRows: 0, // Max rows to return. Use 0 to return all rows
    ignoreAliases: false,
    ignoreSelection: true,
    includeAllColumns: false
  };

  sheet = viz.getWorkbook().getActiveSheet();

  //if active tab is a worksheet, get data from that sheet
  if (sheet.getSheetType() === 'worksheet') {
    sheet.getUnderlyingDataAsync(options).then(function(t) {
      buildMenu(t);
    });

    //if active sheet is a dashboard get data from a specified sheet
  } else {
    worksheetArray = viz.getWorkbook().getActiveSheet().getWorksheets();
    for (var i = 0; i < worksheetArray.length; i++) {
      worksheet = worksheetArray[i];
      sheetName = worksheet.getName();

      //get user's selection from dropdown of sheets
      var selectedVal = getKeyByValue(title_sheet,document.getElementById("SheetList").value);

      //get the data from the selected sheet
      if (sheetName == selectedVal) {
        worksheetArray[i].getSummaryDataAsync(options).then(function(t) {
          buildMenu(t, button_number);
        });
      }
    }
  }
}

//restructure the data and build something with it
function buildMenu(table, bno) {

  //the data returned from the tableau API
  var columns = table.getColumns();
  var data = table.getData();
  // console.log(data);

  //convert to field:values convention
  function reduceToObjects(cols, data) {
    var fieldNameMap = $.map(cols, function(col) {
      return col.$impl.$fieldName;
    });
    var dataToReturn = $.map(data, function(d) {
      return d.reduce(function(memo, value, idx) {
        if (value.value == "%null%"){
          memo[fieldNameMap[idx]] = " ";
        }
        else if(value.value in col_alias){
          memo[fieldNameMap[idx]] = col_alias[value.value];
          console.log(value.value);
        }
        else{
          memo[fieldNameMap[idx]] = value.value;
        }

        return memo;
      }, {});
    });

    return dataToReturn;

  }

  var niceData = reduceToObjects(columns, data);
  //take the niceData and send it to a csv named TableauDataExport
  if(bno == 1){
      // if bno == 1, export to .csv file
      alasql("SELECT * INTO CSV('TableauDataExport.csv',{headers:true, separator:','}) FROM ?", [niceData]);
  }
  if(bno == 2){
      // if bno == 2, export to .xls file
      alasql("SELECT * INTO XLSXML('Tableau.xls', {headers:true}) FROM ?", [niceData]);
  }

}

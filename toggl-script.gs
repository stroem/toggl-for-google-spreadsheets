// Change the API key
var api_key = "b5b57917299dc9e0bf8b294f96fa9027";

function onOpen() {
    var spr = SpreadsheetApp.getActiveSpreadsheet();
    var menu_entries = [ {name: "Update times", functionName: "update"}                  ];
    spr.addMenu("Toggl", menu_entries);
}

function update() {
  
    var format = "yyyy-MM-dd HH:mm:ss";
    var entries = api("time_entries");
  
    var spr = SpreadsheetApp.getActiveSheet();
    var ct = getFirstEmptyRow();
  
    var ct_decr = 0;
    var latest_id = "";
    while(latest_id == "" && ct_decr <= ct) {
      ct_decr++;
      latest_id = spr.getRange(ct - ct_decr, 1).getValue();
    }
    
    var list = [];
    for(var i = entries.length - 1; i >= 0; i--) {
      
        if(entries[i]['id'] == latest_id)
            break;
        
        list.unshift(entries[i]);
    }
     
    for(var i = 0; i < list.length; i++) {
        var project = api("projects/" + list[i]["pid"]).data;
      
        list[i]['duration'] = list[i]['duration'] < 0 ? 0 : list[i]['duration'];
      
        spr.getRange(ct + i, 1).setValue(list[i]['id']);
        spr.getRange(ct + i, 2).setValue(project['name']);
        spr.getRange(ct + i, 3).setValue(formatDate(list[i]['start'], format));
        spr.getRange(ct + i, 4).setValue(formatDate(list[i]['stop'], format));
        spr.getRange(ct + i, 5).setValue(list[i]['duration'] / 3600);
        spr.getRange(ct + i, 6).setValue(list[i]['billable']);
        spr.getRange(ct + i, 7).setValue(formatDate(list[i]['at'], format));
        spr.getRange(ct + i, 8).setValue(list[i]['tags'].join(", "));
        spr.getRange(ct + i, 9).setValue(list[i]['description'] || '');
    }

}

function api(url) {
    var digest = Utilities.base64Encode(api_key + ":api_token");
    var digestfull = "Basic " + digest;

    var response = UrlFetchApp.fetch("https://www.toggl.com/api/v8/" + url, {
        method: "get",
        headers: {
            "Authorization": digestfull,
        }
    });
  
    return Utilities.jsonParse(response.getContentText());
}

function formatDate(dateStr, format) {
    if(dateStr == undefined)
        return "";
  
    return Utilities.formatDate(isoToDate(dateStr), "GMT+2", format)
}

function isoToDate(dateStr){
    var str = dateStr.replace(/-/,'/').replace(/-/,'/').replace(/T/,' ').replace(/\+/,' \+').replace(/Z/,' +00');
    return new Date(str);
}

function getFirstEmptyRow() {
    var spr = SpreadsheetApp.getActiveSpreadsheet();
    var column = spr.getRange('A:A');
    var values = column.getValues(); // get all data in one call
  
    var ct = 0;
    while ( values[ct][0] != "" ) {
      ct++;
    }
  
    return ct + 1;
}

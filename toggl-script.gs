// Change the API key
var api_key = "<INSERT_MAGIC_STRING_HERE>";

var date_format = "yyyy-MM-dd HH:mm:ss";
var base_url = "https://www.toggl.com/api/v8/";
var spr = SpreadsheetApp.getActiveSpreadsheet();

function onOpen() {
    var menu_entries = [
        {name: "Update times", functionName: "update"},
        {name: "Start timer", functionName: "start"},
        {name: "Stop timer", functionName: "stop"}
    ];
      
    spr.addMenu("Toggl", menu_entries);
}

function updateCell(row, entry) {
    var sps = spr.getActiveSheet();
    
    if(entry["pid"] != undefined){
        var project_name = api("GET", "projects/" + entry["pid"]).data.name;
    }else{
        var project_name = '';
    }
    
    entry['duration'] = entry['duration'] < 0 ? 0 : entry['duration'];
      
    sps.getRange(row, 1).setValue(entry['id']);
    sps.getRange(row, 2).setValue(project_name);
    sps.getRange(row, 3).setValue(formatDate(entry['start'], date_format));
    sps.getRange(row, 4).setValue(formatDate(entry['stop'], date_format));
    sps.getRange(row, 5).setValue(entry['duration'] / 3600);
    sps.getRange(row, 6).setValue(entry['billable']);
    sps.getRange(row, 7).setValue(formatDate(entry['at'], date_format));
    sps.getRange(row, 8).setValue(entry['tags'] ? entry['tags'].join(", ") : "");
    sps.getRange(row, 9).setValue(entry['description'] || ''); 
}

function stop() {
    var sps = spr.getActiveSheet();
    var ct = getFirstEmptyRow();
  
    var entry_id = sps.getRange(ct, 1);
    var entry_stop = sps.getRange(ct, 4).getValue();
    if(entry_stop == "" && entry_id != "") {
        api("PUT", "time_entries/" + entry_id + "/stop");
        var entry = api("GET", "time_entries/" + entry_id).data;
      
        updateCell(ct, entry);
    }
}
      
function start() {
      var workspaces = api("GET", "workspaces");
      
      var data = [];
      for(var i = 0; i < workspaces.length; i++) {
          data.push({
              "name": workspaces[i].name,
              "value": workspaces[i].id
          });
      }
  
      modal("Select workspace", "workspace", null, null, data);
}

function doPost(event) {
    var app = UiApp.getActiveApplication();
  
    var workspace = event.parameter.workspace;
    if(workspace != undefined) {
        var projects = api("GET", "workspaces/" + workspace + "/projects");
      
        var data = [];
        for(var i = 0; i < projects.length; i++) {
            data.push({
                "name": projects[i].name,
                "value": projects[i].id
            });
        }
      
        modal("Select project", "project", "Description:", "description", data);
    }  
  
    var project = event.parameter.project;
    if(project) {
      var description = event.parameter.description || "";
      
        var data = {
            "time_entry": {
                "description": description,
                "tags": ["excel"],
                "pid": project
            }
        };
      
        api("POST", "time_entries/start", data);
    }
  
    return app.close();
}
      
function modal(title, radiobutton_name, textbox_title, textbox_name, data) {
      var app = UiApp.createApplication();
      app.setTitle(title);
      app.setWidth(250);
      app.setHeight(300);
      
      var form_content = app.createGrid();
      form_content.resize(14, 3);
      
      for(var i = 0; i < data.length; i++) {
          var radio_button = app.createRadioButton(radiobutton_name, data[i].name);
          radio_button.setFormValue(data[i].value);
          form_content.setWidget(i, 0, radio_button);
      }
  
      if(textbox_name) {
          var label = app.createLabel(textbox_title);
          form_content.setWidget(9, 0, label);
    
          var textbox = app.createTextBox().setName(textbox_name); 
          form_content.setWidget(10, 0, textbox);
      }
  
      var button = app.createSubmitButton("Select");
      form_content.setWidget(12, 0, button);
  
      var form = app.createFormPanel().setId('form').setEncoding('multipart/form-data');
      form.add(form_content);
      
      app.add(form);
      spr.show(app);
}
      
function update() {
    var entries = api("GET", "time_entries");
  
    var sps = spr.getActiveSheet();
    var ct = getFirstEmptyRow();
  
    var ct_decr = 0;
    var latest_id = "";
    while(latest_id == "" && ct_decr <= ct) {
        ct_decr++;
        latest_id = sps.getRange(ct - ct_decr, 1).getValue();
    }
    
    var list = [];
    for(var i = entries.length - 1; i >= 0; i--) {
      
        if(entries[i]['id'] == latest_id)
            break;
        
        list.unshift(entries[i]);
    }
     
    for(var i = 0; i < list.length; i++) {
        updateCell(ct + i, list[i]);
    }
}

function api(method, url, data) {
    var digest = Utilities.base64Encode(api_key + ":api_token");
    var digestfull = "Basic " + digest;

    var options = {
        method: method,
        headers: {
            "Authorization": digestfull,
        }
    };

    if(data != undefined) {
        options["payload"] = data;
    }

    var response = UrlFetchApp.fetch(base_url + url, options);
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
    var column = spr.getRange('A:A');
    var values = column.getValues(); // get all data in one call
  
    var ct = 0;
    while ( values[ct][0] != "" ) {
      ct++;
    }
  
    return ct + 1;
}

// TODO:
// Add in the possiblity for submitting images, files, and videos to the form
// Thats about it

const ticket_spreadsheet_id = "SPREADSHEET ID";
const home_sheet = "Home";
var max = 50;

function publishTicketHandler(){
  var spread = SpreadsheetApp.openById(ticket_spreadsheet_id);
  var sheet = SpreadsheetApp.getActiveSheet();
  console.log(sheet.getName());
  var config = spread.getSheetByName("Config");
  var home = spread.getSheetByName(home_sheet);

  var config_values = config.getRange(1, 1, max, max).getValues();

  max = find_config_value("Maximum Search Length", config_values);

  var sheet_values = sheet.getRange(1, 1, max, max).getValues();
  var home_range = home.getRange(1, 1, max, max);
  var home_format = home_range.getRichTextValues();

  var home_values = home_range.getValues();

  publishTicket(sheet, config, home, config_values, sheet_values, home_format, home_values, max);
}

function publishAll(){
  var spread = SpreadsheetApp.openById(ticket_spreadsheet_id);
  var sheets = spread.getSheets();
  var config = spread.getSheetByName("Config");
  var home = spread.getSheetByName(home_sheet);

  var config_values = config.getRange(1, 1, max, max).getValues();

  max = find_config_value("Maximum Search Length", config_values);

  var home_range = home.getRange(1, 1, max, max);
  var home_format = home_range.getRichTextValues();

  var home_values = home_range.getValues();
  const non_tickets = find_sections("Non-Ticket Sheets", config_values);

  sheets.forEach((sheet) => {
    if (!non_tickets.includes(sheet.getName())){
      var sheet_values = sheet.getRange(1, 1, max, max).getValues();
      publishTicket(sheet, config, home, config_values, sheet_values, home_format, home_values, max);
    }
  });
}

function publishTicket(sheet, config, home, config_values, sheet_values, home_format, home_values, max) {

  var act_sheets = find_config_value("Current Active Tickets", config_values);
  const prioity_list = find_sections("Priority List", config_values);
  const prioity_title = find_config_value("Priority Section Title", config_values)
  const auto_publish = find_config_value("Auto Publish", config_values);
  const sections = find_sections("Form Section Names",config_values);
  const filtered_secs = find_sections("Filtered Section Names", config_values);
  const internal_sec = find_config_value("Internal Section Name", config_values);
  const total_tickets = find_config_value("Total Saved Tickets", config_values);
  const sheet_n = find_config_value("Internal Sheet Number Title", config_values);

  const sub_cell = sheet_indexer(sections[0], home_values);
  var existing = true;
  var prioity;

  const status_categories = find_sections("Internal Status Naming", config_values);
  const status = status_categories[0];
  status_categories.shift();

  const categories = [];
  const items = [];

  let int_sec_cell = sheet_indexer(internal_sec, sheet_values);

  var int_category = []
  var int_item = []
  
  for (var i = 1; i < max; i++){
    let temp_category = sheet_values[int_sec_cell[0] + i - 1][int_sec_cell[1] - 2];
    if (temp_category == ""){
      break;
    }

    int_category.push(temp_category);
    int_item.push(sheet_values[int_sec_cell[0] + i - 1][int_sec_cell[1] - 1])
  }

  if (int_category.includes(status) && int_item.includes(status_categories[0]) && auto_publish){
    sheet.getRange(int_sec_cell[0] + int_category.indexOf(status) + 1, int_sec_cell[1], 1, 1).setValue(status_categories[1]);
    sheet_values[int_sec_cell[0] + int_category.indexOf(status)][int_sec_cell[1] - 1] = status_categories[1];
  }
  if (int_category.includes(sheet_n)){
    var num = sheet_values[int_sec_cell[0] + int_category.indexOf(sheet_n)][int_sec_cell[1] - 1];
    if (num == ""){
      existing = false;
      sheet.getRange(int_sec_cell[0] + int_category.indexOf(sheet_n) + 1, int_sec_cell[1], 1, 1).setValue(total_tickets + 1);
      sheet_values[int_sec_cell[0] + int_category.indexOf(sheet_n)][int_sec_cell[1] - 1] = total_tickets + 1;
      set_active_sheets("Total Saved Tickets", config_values, total_tickets + 1, config);
    }
  }

  if (!existing){
    set_active_sheets("Current Active Tickets", config_values, act_sheets + 1, config);
  }
  else{
    var existing_row = sheet_indexer(int_item[int_category.indexOf(sheet_n)], home_values)[0];

    if (existing_row != 0){
      act_sheets = existing_row - sub_cell[0] - 1;
    }
    else{
      sheet.getRange(int_sec_cell[0] + int_category.indexOf(sheet_n) + 1, int_sec_cell[1], 1, 1).setValue(total_tickets + 1);
      sheet_values[int_sec_cell[0] + int_category.indexOf(sheet_n)][int_sec_cell[1] - 1] = total_tickets + 1;
      set_active_sheets("Total Saved Tickets", config_values, total_tickets + 1, config);
      set_active_sheets("Current Active Tickets", config_values, act_sheets + 1, config);
      existing = false;
    }
  }

  filtered_secs.forEach(sec => {
    let fil_sec_cell = sheet_indexer(sec, sheet_values);

    for (var i = 1; i < max; i++){
      let category = sheet_values[fil_sec_cell[0] + i - 1][fil_sec_cell[1] - 2];
      let item = sheet_values[fil_sec_cell[0] + i - 1][fil_sec_cell[1] - 1];

      if (category == ""){
        break;
      }
      if (category.toString().includes("Priority")){
        prioity = item;
      }

      categories.push(category);
      items.push(item);
    }
  });

  var sheet_name = SpreadsheetApp.newRichTextValue()
   .setText(sheet.getName())
   .setLinkUrl("https://docs.google.com/spreadsheets/d/" + ticket_spreadsheet_id + "/edit#gid=" + sheet.getSheetId())
   .build();

  const prioity_cell = sheet_indexer(prioity_title, home_values);
  const prioity_index = prioity_list.indexOf(prioity);

  home.getRange(sub_cell[0] + act_sheets + 1, sub_cell[1], 1, 1).setRichTextValue(sheet_name);
  home_values[sub_cell[0] + act_sheets][sub_cell[1] - 1] = sheet_name.getText();
  home_format[sub_cell[0] + act_sheets][sub_cell[1] - 1] = sheet_name;

  if (existing) {
    for (var i = 2; i < max; i++){
      for (var j = 0; j < prioity_list[0].length - 1; j++){
        if (home_format[prioity_cell[0] + i - 1][prioity_cell[1] + j - 1].getLinkUrl() == sheet_name.getLinkUrl()){
          home.getRange(prioity_cell[0] + i, prioity_cell[1] + j, 1, 1).deleteCells(SpreadsheetApp.Dimension.ROWS);
          for (var temp_row = prioity_cell[0] + i - 1; temp_row < home_values.length - 1; temp_row++){
            home_format[temp_row][prioity_cell[1] + j - 1] = home_format[temp_row + 1][prioity_cell[1] + j - 1];
            home_values[temp_row][prioity_cell[1] + j - 1] = home_values[temp_row + 1][prioity_cell[1] + j - 1];
          }
          current_prioity_tickets(j,- 1, config_values, config);
          i = max;
          j = prioity_list[0].length;
        }
      }
    }
  }
  var cur_pri_tick = current_prioity_tickets(prioity_index, 1, config_values, config);

  home.getRange(prioity_cell[0] + cur_pri_tick + 2, prioity_cell[1] + prioity_index, 1, 1).setRichTextValue(sheet_name);

  home_values[prioity_cell[0] + cur_pri_tick + 1][prioity_cell[1] + prioity_index - 1] = sheet_name.getText();
  home_format[prioity_cell[0] + cur_pri_tick + 1][prioity_cell[1] + prioity_index - 1] = sheet_name;

  var cat_row = home.getRange(sub_cell[0], 1, 1, max).getValues()[0]

  categories.forEach((category, index) =>{
    let cat_pos = row_indexer(cat_row, category);

      if (cat_pos === 0){
        cat_pos = find_num_categories(cat_row);

        home.getRange(sub_cell[0], cat_pos).setValue(category);
        home_values[sub_cell[0] - 1][cat_pos - 1] = category;
        cat_row[cat_pos - 1] = category;
      }
      home.getRange(sub_cell[0] + act_sheets + 1, cat_pos, 1, 1).setValue(items[index]);
      home_values[sub_cell[0] + act_sheets][cat_pos - 1] = items[index];
  });
}

function find_num_categories(values){
  if (!values.includes("")){
    SpreadsheetApp.getUI.alert("Maximum Search Length has been set too low in Config")
    return
  }
  return(values.indexOf("") + 1)
}

function row_indexer(row, part){
  if (row.includes(part)){
    return row.indexOf(part) + 1;
  }
  return 0;
}

function sheet_indexer(part, sheet){
  for (var i = 1; i < max; i++){
    if (sheet[i - 1].includes(part)){
      return [i, sheet[i - 1].indexOf(part) + 1];
    }
  }
  return [0, 0];
}

function find_sections(name, sheet){
  const cell = sheet_indexer(name, sheet);
  const sections = [];

  for (var i = 1; i < max; i++){
    var value = sheet[cell[0] + i][cell[1] - 1];
    if (value == ""){
      return sections;
    }
    sections.push(value);
  }

  return [0, 0];
}

function find_config_value(name, sheet){
  if (!sheet[0].includes(name)){
    SpreadsheetApp.getUi().alert("The config value:" + name + " does not exist");
    return;
  }
  const cell = row_indexer(sheet[0], name);
  return sheet[2][cell - 1];
}

function current_prioity_tickets(prioity, value, sheet, config){
  const cell = sheet_indexer("Current Priority Tickets", sheet);
  const ticket_cell = config.getRange(cell[0] + prioity + 2, cell[1], 1, 1);
  const tickets = sheet[cell[0] + prioity + 1][cell[1] - 1];
  
  ticket_cell.setValue(tickets + value);
  sheet[cell[0] + prioity + 1][cell[1] - 1] = tickets + value
  return tickets;
}

function set_active_sheets(part, sheet, value, config){
  const cell = sheet_indexer(part, sheet);
  config.getRange(cell[0] + 2, cell[1], 1, 1).setValue(value);
  sheet[cell[0] + 1][cell[1] - 1] = value;
  return;
}

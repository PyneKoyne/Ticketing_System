// TODO:
// Add in the possiblity for submitting images, files, and videos to the form
// Add data modification form delete_priority
// Create an update_ticket method

const ticket_spreadsheet_id = "Spreadsheet ID";
const spread = SpreadsheetApp.openById(ticket_spreadsheet_id);
const ss_info = spread.getSheetByName("Spreadsheet Info");
let home_sheet;
let config_sheet;
let template_sheet;
let sub_cell;

let priority_list;
let priority_title;
let auto_publish;
let sections;
let filtered_secs;
let internal_sec;
let sheet_n;
let priority_exclusion;
let status_categories;
let status;

let max = 50;

// Resets All Stored Data
function resetProps(){
  const scriptProperties = PropertiesService.getScriptProperties();
  console.log(scriptProperties.getProperties());
  scriptProperties.deleteAllProperties();
}

async function createEmptyHandler(){
  const ui = SpreadsheetApp.getUi();

  const result = ui.prompt(
      'Sheet Name: ',
      ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  const button = result.getSelectedButton();
  const text = result.getResponseText();

  if (button === ui.Button.OK) {
    await createEmpty(text);
  }
}

// Creates an empty ticket
async function createEmpty(name) {
  const template = spread.getSheetByName(template_sheet);

  let unique_name = false;
  let counter = 0;
  const sheet = template.copyTo(spread);

  while (!unique_name) {
    try {
      // If the original name is unique, it doesn't add a 0 at the end
      if (counter === 0) {
        sheet.setName(name);
      } else {
        sheet.setName(name + counter);
      }

      unique_name = true;
    } catch (err) {
      counter++;
    }
  }
  sheet.showSheet();

  await publishTicketHandler(["___SET-SUBJECT___", sheet]);
}

// Function to publish all tickets
async function publishAll(index_section) {
  await publishTicketHandler(["___ALL___", index_section]);
}

async function publishOne() {
  const ticket = spread.getActiveSheet();
  const sheet_range = ticket.getDataRange();
  const sheet_class = new SheetClass(ticket, sheet_range);
  await publishTicketHandler([sheet_class]);
}

async function publishTicketHandler(tickets) {
  const config = spread.getSheetByName(config_sheet);
  const home = spread.getSheetByName(home_sheet);

  const scriptProperties = PropertiesService.getScriptProperties();
  const raw_sheets = home_builder(scriptProperties, home, home_sheet, config);
  let setting = null;

  let home_class = raw_sheets[0];

  let config_values = raw_sheets[1];
  setConstants(config_values);

  // The Subject cell of the home page
  sub_cell = sheet_indexer(sections[0], home_class.values);

  // If called from PublishAll
  if (tickets[0] === "___ALL___") {
    setting = tickets[0];

    const priority_cell = sheet_indexer(priority_title, home_class.values);

    new Promise(function (resolve) {
      resolve(home.getRange(priority_cell[0] + 2, priority_cell[1], max, priority_list.length).clearContent());
    });

    for (let i = 1; i < home_class.values.length - priority_cell[0]; i++) {
      for (let j = 0; j < priority_list.length; j++) {
        home_class.format[priority_cell[0] + i][priority_cell[1] + j - 1] = null;
        home_class.values[priority_cell[0] + i][priority_cell[1] + j - 1] = null;
      }
    }

    set_config_value("Total Saved Tickets", config_values, config, 0);
    set_config_value("Current Active Tickets", config_values, config, 0);

    const cell = sheet_indexer("Current Priority Tickets", config_values);

    for (let i = 0; i < max; i++) {
      config_values[cell[0] + i][cell[1] - 1] = 0;
      new Promise(function(resolve){
        resolve(config.getRange(cell[0] + i + 1, cell[1]).setValue(0));
      });
    }

    findTickets(scriptProperties);

    for (let i = 0; i < tickets.length; i++) {
      let ticket = tickets[i];
      tickets[i] = new Promise(function (resolve) {
        const sheet = spread.getSheetByName(ticket);
        const sheet_range = sheet.getDataRange();
        resolve(new SheetClass(sheet, null, sheet_range.getValues(), sheet_range.getRichTextValues()));
      }).then(r => {
        publishTicket(config, config_values, r, home_class, max, setting);
      });
    }
  }

  // If called from CreateEmpty
  if (tickets[0] === "___SET-SUBJECT___") {
    const sheet_range = tickets[1].getDataRange();
    const sheet_class = new SheetClass(tickets[1], null, sheet_range.getValues(), sheet_range.getRichTextValues());
    let sheet_sub_cell = sheet_indexer(sections[0], sheet_class.values);
    sheet_class.setValues([sheet_sub_cell[0] + 1, sheet_sub_cell[1]], tickets[1].getName());

    let saved_tickets = findTickets(scriptProperties);
    saved_tickets.push(tickets[1].getName());

    new Promise(function (resolve, reject) {
      resolve(scriptProperties.setProperty(home_sheet + "_tickets", packProperties(saved_tickets)));
    }).then(r => console.log("Tickets Uploaded"));

    tickets = [sheet_class];
  }

  for (const ticket of tickets) {
    if (typeof ticket === "object" && typeof ticket.then === "function") {
      await ticket;
    }
    else {
      publishTicket(config, config_values, ticket, home_class, max, setting);
    }
  }

  await scriptProperties.setProperty(home_sheet, packProperties({
    "values": formatToJSON(home_class.format, home_class.values, false),
    "config": config_values,
  }));

  console.log("Publish Properties Set")
}

// MAIN TICKET UPDATING METHOD
function publishTicket(config, config_values, sheet_class, home_class, max, setting) {
  // Config Values
  let act_sheets = find_config_value("Current Active Tickets", config_values);
  const total_tickets = find_config_value("Total Saved Tickets", config_values);

  // A bool to check if the ticket is already existing
  let existing = true;

  let priority;
  let int_sec_cell = sheet_indexer(internal_sec, sheet_class.values);

  if (int_sec_cell[0] === 0 && int_sec_cell[1] === 0){
    throw "Internal Section Not Found";
  }

  const int_category = [];
  const int_item = [];

  // loops through all internal categories
  for (let i = 0; i < sheet_class.values.length - int_sec_cell[0]; i++) {
    let temp_category = sheet_class.values[int_sec_cell[0] + i][int_sec_cell[1] - 2];
    if (temp_category == null) {
      break;
    }

    int_category.push(temp_category);
    int_item.push(sheet_class.values[int_sec_cell[0] + i][int_sec_cell[1] - 1])
  }

  const int_status = [int_sec_cell[0] + int_category.indexOf(status) + 1, int_sec_cell[1]]

  if (int_category.includes(status) && int_item.includes(status_categories[0]) && auto_publish) {
    sheet_class.setValues(int_status, status_categories[1]);
  }

  if (int_category.includes(sheet_n)) {
    const num = sheet_class.values[int_sec_cell[0] + int_category.indexOf(sheet_n)][int_sec_cell[1] - 1];
    if (num == null || setting === "___ALL___") {
      existing = false;
      sheet_class.setValues([int_sec_cell[0] + int_category.indexOf(sheet_n) + 1, int_sec_cell[1]], total_tickets + 1);
      set_active_sheets("Total Saved Tickets", config_values, total_tickets + 1, config);
    }
  }

  if (!existing) {
    set_active_sheets("Current Active Tickets", config_values, act_sheets + 1, config);
  } else {
    let existing_row = null;

    for (let i = 0; i < home_class.values.length - sub_cell[0]; i++) {
      let link_cell = home_class.format[sub_cell[0] + i][sub_cell[1] - 1];
      if (link_cell == null || link_cell === "") {
        break;
      } else {
        for (let j = 0; j < link_cell.length; j++) {
          if (link_cell[j][2].includes(sheet_class.sheet.getSheetId().toString())) {
            existing_row = 1 + i + sub_cell[0];
            break;
          }
        }
      }
    }

    if (existing_row != null) {
      act_sheets = existing_row - sub_cell[0] - 1;
    } else {
      sheet_class.setValues([int_sec_cell[0] + int_category.indexOf(sheet_n) + 1, int_sec_cell[1]], total_tickets + 1);
      set_active_sheets("Total Saved Tickets", config_values, total_tickets + 1, config);
      set_active_sheets("Current Active Tickets", config_values, act_sheets + 1, config);
      existing = false;
    }
  }

  const cat_row = home_class.values[sub_cell[0] - 1];

  filtered_secs.forEach(sec => {
    let fil_sec_cell = sheet_indexer(sec, sheet_class.values);
    console.log(sec);

    for (let i = 0; i < sheet_class.values.length - fil_sec_cell[0]; i++) {
      let category = sheet_class.values[fil_sec_cell[0] + i][fil_sec_cell[1] - 2];
      let item = sheet_class.values[fil_sec_cell[0] + i][fil_sec_cell[1] - 1];

      if (category === "" || category == null) break;

      // Normally checks if the format and values have an incongruity, that is values has a value but format doesn't
      // This is fixed from the sanitization in SheetClass
      //
      // if (sheet_class.values[fil_sec_cell[0] + i - 1][fil_sec_cell[1] - 1] == null) {
      //     item = item.copy().setText(sheet_class.values[fil_sec_cell[0] + i - 1][fil_sec_cell[1] - 1]).build();
      // }

      if (category.includes("Priority")) {
        priority = item;
      }

      let cat_pos = row_indexer(cat_row, category);

      if (cat_pos === 0) {
        cat_pos = find_num_categories(cat_row);

        home_class.setValues([sub_cell[0], cat_pos], category);
        cat_row[cat_pos - 1] = category;
      }
      home_class.setRichValue([sub_cell[0] + act_sheets + 1, cat_pos],
          item, sheet_class.format[fil_sec_cell[0] + i][fil_sec_cell[1] - 1], true);
    }
  });

  const subject = sheet_indexer(sections[0], sheet_class.values);

  // Old Implementation of sheet_name
  //
  // const sheet_name = SpreadsheetApp.newRichTextValue()
  //     .setText(sheet_class.values[subject[0]][subject[1] - 1])
  //     .setLinkUrl("#gid=" + sheet_class.sheet.getSheetId())
  //     .build();

  const sheet_name = [
    [0, sheet_class.values[subject[0]][subject[1] - 1].length, "#gid=" + sheet_class.sheet.getSheetId(), 0, 0, 1]];

  const priority_cell = sheet_indexer(priority_title, home_class.values);
  const priority_index = priority_list.indexOf(priority);

  home_class.setRichValue([sub_cell[0] + act_sheets + 1, sub_cell[1]], sheet_class.values[subject[0]][subject[1] - 1], sheet_name, false);

  console.log("Is this ticket existing: " + existing);

  if (existing) {
    deletePriority(config, config_values, home_class, priority_list, priority_cell, "#gid=" + sheet_class.sheet.getSheetId(), max);
  }

  if (!priority_exclusion.includes(sheet_class.values[int_status[0] - 1][int_status[1] - 1])) {
    const cur_pri_tick = current_priority_tickets(priority_index, 1, config_values, config);

    home_class.setRichValue([priority_cell[0] + cur_pri_tick + 2, priority_cell[1] + priority_index], sheet_class.values[subject[0]][subject[1] - 1], sheet_name);
  }
}

function deleteTicketHandler() {
  const ui = SpreadsheetApp.getUi();

  const config = spread.getSheetByName(config_sheet);
  const home = spread.getSheetByName(home_sheet);

  const scriptProperties = PropertiesService.getScriptProperties();
  const raw_sheets = home_builder(scriptProperties, home, home_sheet, config);

  let home_class = raw_sheets[0];
  let config_values = raw_sheets[1];
  setConstants(config_values);

  const result = ui.prompt(
      sheet_n,
      ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  const button = result.getSelectedButton();
  const text = result.getResponseText();

  if (button === ui.Button.OK) {
    deleteTicket(config, spread, home_class, config_values, text);
  }
}

function deleteTicket(config, spread, home_class, config_values, text) {
  const home_n = sheet_indexer(sheet_n, home_class.values);
  const act_sheets = find_config_value("Current Active Tickets", config_values);
  const priority_cell = sheet_indexer(priority_title, home_class.values);

  for (let i = 0; i < max; i++) {
    if (home_class.values[i][home_n[1] - 1] === text) {
      let link = home_class.format[i][sub_cell[1] - 1];
      let sheet = null;
      for (let j = 0; j < link.length; j++) {
        console.log(link.split("gid="));
        if (link[j][2] !== 0) {
          sheet = spread.getSheets().filter(function (s) {
            return s.getSheetId().toString() === link[j][2].split("gid=")[1];
          })[0];

          deletePriority(config, config_values, home_class, priority_list, priority_cell, link[j][2], max)
          const cat_row = home_class.values[sub_cell[0] - 1];

          const last_cat = find_num_categories(cat_row);
          home_class.sheet.getRange(i + 1, sub_cell[1], 1, last_cat - sub_cell[1]).deleteCells(SpreadsheetApp.Dimension.ROWS);
          set_active_sheets("Current Active Tickets", config_values, act_sheets - 1, config);

          if (sheet != null) {
            spread.deleteSheet(sheet);
          }
          return;
        }
      }
    }
  }
}

function deletePriority(config, config_values, home_class, priority_list, priority_cell, link, max) {
  for (let i = 2; i < home_class.values.length - priority_cell; i++) {
    for (let j = -1; j < priority_list.length; j++) {

      let checkCell = home_class.format[priority_cell[0] + i - 1][priority_cell[1] + j - 1];
      if (checkCell != null && checkCell !== "") {

        console.log((priority_cell[0] + i - 1) + ", " +  (priority_cell[1] + j - 1) + " priority value is: " +
            home_class.value[priority_cell[0] + i - 1][priority_cell[1] + j - 1]);
        for (let k = 0; k < checkCell.length; k++) {

          if (checkCell[k][2] === link) {
            home_class.sheet.getRange(priority_cell[0] + i, priority_cell[1] + j, 1, 1).deleteCells(SpreadsheetApp.Dimension.ROWS);

            for (let temp_row = priority_cell[0] + i - 1; temp_row < home_class.values.length - 1; temp_row++) {
              home_class.format[temp_row][priority_cell[1] + j - 1] = home_class.format[temp_row + 1][priority_cell[1] + j - 1];
              home_class.values[temp_row][priority_cell[1] + j - 1] = home_class.format[temp_row][priority_cell[1] + j - 1];
            }
            current_priority_tickets(j, -1, config_values, config);
            i = max;
            j = priority_list.length;
          }
        }
      }
    }
  }
}

function find_num_categories(values) {
  if (!values.includes("")) {
    if(!values.includes(null)) {
      SpreadsheetApp.getUI.alert("Maximum Search Length has been set too low in Config")
      return;
    }
    return (values.indexOf(null) + 1);
  }
  return (values.indexOf("") + 1)
}

function row_indexer(row, part) {
  if (row.includes(part)) {
    return row.indexOf(part) + 1;
  }
  return 0;
}

function sheet_indexer(part, sheet) {
  for (let i = 0; i < sheet.length; i++) {
    if (sheet[i].includes(part)) {
      return [i + 1, sheet[i].indexOf(part) + 1];
    }
  }
  return [0, 0];
}

function find_sections(name, sheet) {
  const cell = sheet_indexer(name, sheet);
  const sections = [];

  for (let i = 1; i < sheet.length; i++) {
    const value = sheet[cell[0] + i][cell[1] - 1];
    if (value == "") {
      return sections;
    }
    sections.push(value);
  }

  return [0, 0];
}

function find_config_value(name, sheet) {
  if (!sheet[0].includes(name)) {
    SpreadsheetApp.getUi().alert("The config value:" + name + " does not exist");
    return;
  }
  const cell = row_indexer(sheet[0], name);
  return sheet[2][cell - 1];
}

function set_config_value(name, sheet, config, value) {
  if (!sheet[0].includes(name)) {
    SpreadsheetApp.getUi().alert("The config value:" + name + " does not exist");
    return;
  }
  const cell = row_indexer(sheet[0], name);
  setSSValues(config, [3, cell], value).then(r => console.log("config value: " + name + " set as : " + value));
  sheet[2][cell - 1] = value;
}

function current_priority_tickets(priority, value, sheet, config) {
  const cell = sheet_indexer("Current Priority Tickets", sheet);
  const ticket_cell = config.getRange(cell[0] + priority + 2, cell[1], 1, 1);
  const tickets = sheet[cell[0] + priority + 1][cell[1] - 1];

  ticket_cell.setValue(tickets + value);
  sheet[cell[0] + priority + 1][cell[1] - 1] = tickets + value
  return tickets;
}

function set_active_sheets(part, sheet, value, config) {
  const cell = sheet_indexer(part, sheet);
  setSSValues(config, [cell[0] + 2, cell[1]], value).then(r => console.log("active sheets set as : " + value));
  sheet[cell[0] + 1][cell[1] - 1] = value;
  return value;
}
// Created By Kenny Z
// Script 1 of 2 for the Ticketing System

// Ids of the Form and Spreadsheet
const form_id = "form_id";
const ticket_spreadsheet_id = "ticket_id";

const template_sheet = "Template";
const config_sheet = "Config"

// Run this function to set up the trigger to allow the script below to run every time a form response is recieved
function setUpTrigger(){

  ScriptApp.newTrigger('onFormSubmit')
      .forForm(form_id)
      .onFormSubmit()
      .create();
}

// This script runs everytime a form response is recieved
async function onFormSubmit(e) {

  // Opens the form to figure out the structure
  var form = FormApp.openById(form_id);
  var items = form.getItems();
  var sectionIdxs = [-1];

  // Calculates how many items there are per section
  for (var i = 0; i < items.length; i++){
    var item = items[i];
    var itemType = item.getType();
    var itemIdx = item.getIndex();

    if (itemType == "PAGE_BREAK") {
      sectionIdxs.push(itemIdx - sectionIdxs.at(-1));
    }
  }
  console.log(sectionIdxs);
  sectionIdxs.push(items.length - sectionIdxs.at(-1) - 1);
  sectionIdxs.shift();
  sectionIdxs.shift();

  // Finds the questions and answers of the form response
  var items = e.response.getItemResponses();
  var responses = [];
  var questions = [];

  items.forEach( item => responses.push(item.getResponse()));
  items.forEach( item => questions.push(item.getItem().getTitle()));

  // Opens the spreadsheet
  var spread = SpreadsheetApp.openById(ticket_spreadsheet_id);

  // Opens the template sheet
  var sheet = spread.getSheetByName(template_sheet).copyTo(spread);
  var config = spread.getSheetByName(config_sheet);

  var subject = questions.indexOf("Subject (Title of the Ticket)");

  // Tries naming the sheet
  // If a name is taken, it adds a number at the end going up to infinity
  var unique_name = false;
  var counter = 0;

  while (!unique_name){
    try{
      // If the original name is unique, it doesn't add a 0 at the end
      if (counter == 0){
        sheet.setName(responses[subject]);
      }
      else{
        sheet.setName(responses[subject] + counter);
      }

      unique_name = true;
    }
    catch(err){
      counter ++;
    }
  }

  // Finds the sections for the ticket in the config sheet
  const sections = find_sections1(config);

  // Swaps the subject and dates field so that the subject can be isolated
  let subject_cell = sheet_indexer1(sections[0], sheet);
  sheet.getRange(subject_cell[0] + 1, subject_cell[1], 1, 1).setValue(responses[subject]);

  questions[subject] = "Date";
  responses[subject] = e.response.getTimestamp();

  // Goes through the whole form and formats it onto the ticket
  counter = 0;
  sectionIdxs.forEach((section, sec_num) => {

    var cell = sheet_indexer1(sections[sec_num + 1], sheet);
    for (var i = 0; i < section; i++){
      console.log(questions[i + counter]);
      sheet.getRange(cell[0] + i + 1, cell[1] - 1, 1, 1).setValue(questions[i + counter] + ":");
      sheet.getRange(cell[0] + i + 1, cell[1], 1, 1).setValue(responses[i + counter]);
    }
    counter += section;
  })

  const email_cell = sheet_indexer1("Email", sheet);
  sheet.getRange(email_cell[0], email_cell[1] + 1, 1, 1).setValue(e.response.getRespondentEmail());

  // Unhides the ticket
  sheet.showSheet();

  const sheet_range = sheet.getDataRange();
  const sheet_class = new SheetClass(sheet, sheet_range);
  await publishTicketHandler([sheet_class]);
}

// Finds the location of a cell given part of its value and the sheet
function sheet_indexer1(part, sheet){
  for (var i = 1; i < max; i++){
    for (var j = 1; j < max; j++){
      if (sheet.getRange(i, j, 1, 1).getValue().toString().includes(part)){
        return [i, j];
      }
    }
  }
}

// Finds the sections of the ticket form the config file
function find_sections1(sheet){
  const cell = sheet_indexer1("Form Section Names", sheet);
  const sections = []
  for (var i = 2; i < max; i++){
    var value = sheet.getRange(cell[0] + i, cell[1]).getValue();
    if (value == ""){
      return sections;
    }
    sections.push(value);
  }
}
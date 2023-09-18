// Created By Kenny
// Script 1 of 2 for the Ticketing System

// Ids of the Form and Spreadsheet
const form_id = "FORM ID";
const ticket_spreadsheet_id = "SPREADSHEET ID";

// Max Search Length
var max = 50;

// Run this function to set up the trigger to allow the script below to run every time a form response is recieved
function setUpTrigger(){

  ScriptApp.newTrigger('onFormSubmit')
  .forForm(form_id)
  .onFormSubmit()
  .create();
}

// This script runs everytime a form response is recieved
function onFormSubmit(e) {

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
  sectionIdxs.push(items.length - sectionIdxs.at(-1) - 2);
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
  var sheet = spread.getSheetByName('Template');
  var config = spread.getSheetByName('Config');

  max = get_max(config);

  var subject = questions.indexOf("Subject");

  // Tries naming the sheet
  // If a name is taken, it adds a number at the end going up to infinity
  var unique_name = false;
  var counter = 0;

  while (!unique_name){
    try{
      // If the original name is unique, it doesn't add a 0 at the end
      if (counter == 0){
        sheet = sheet.copyTo(spread).setName(responses[subject]);
      }
      else{
        sheet = sheet.copyTo(spread).setName(responses[subject] + counter);
      }

      unique_name = true;
    }
    catch(err){
      counter ++;
    }
  }

  // Finds the sections for the ticket in the config sheet
  const sections = find_sections(spread.getSheetByName("Config"));

  // Swaps the subject and dates field so that the subject can be isolated
  let subject_cell = sheet_indexer(sections[0], sheet);
  sheet.getRange(subject_cell[0] + 1, subject_cell[1], 1, 1).setValue(responses[subject]);
  
  questions[subject] = "Date";
  responses[subject] = e.response.getTimestamp();

  console.log(sectionIdxs);
  console.log(responses);

  // Goes through the whole form and formats it onto the ticket
  counter = 0;
  sectionIdxs.forEach((section, sec_num) => {

    var cell = sheet_indexer(sections[sec_num + 1], sheet);

    for (var i = 0; i < section; i++){
      console.log(cell);
      sheet.getRange(cell[0] + i + 1, cell[1] - 1, 1, 1).setValue(questions[i + counter] + ":");
      sheet.getRange(cell[0] + i + 1, cell[1], 1, 1).setValue(responses[i + counter]);
    }
    counter += section;
  })

  const email_cell = sheet_indexer("Email", sheet);
  sheet.getRange(email_cell[0], email_cell[1] + 1, 1, 1).setValue(e.response.getRespondentEmail());

  // Unhides the ticket
  sheet.showSheet();
  
  GmailApp.sendEmail("kenny.zheng@student.tdsb.on.ca", responses[1], responses);
}

// Finds the location of a cell given part of its value and the sheet
function sheet_indexer(part, sheet){
  for (var i = 1; i < max; i++){
    for (var j = 1; j < max; j++){
      if (sheet.getRange(i, j, 1, 1).getValue().toString().includes(part)){
        return [i, j];
      }
    }
  }
}

function get_max(sheet){
  const cell = sheet_indexer("Maximum Search Length", sheet);
  return sheet.getRange(cell[0] + 2, cell[1], 1, 1).getValue();
}

// Finds the sections of the ticket form the config file
function find_sections(sheet){
  const cell = sheet_indexer("Form Section Names", sheet);
  const sections = []
  for (var i = 2; i < max; i++){
    var value = sheet.getRange(cell[0] + i, cell[1]).getValue();
    if (value == ""){
      return sections;
    }
    sections.push(value);
  }
}

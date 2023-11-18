class SheetClass {
    constructor(sheet, range, values, format) {
        this.sheet = sheet;

        if (range !== null) {
            this.values = range.getValues();
            this.format = range.getRichTextValues();
        } else {
            this.values = values;
            this.format = format;
        }

        this.normal_text = SpreadsheetApp.newTextStyle()
            .setFontFamily("League Spartan")
            .build();
    }

    setValues(cell, value) {
        this.checkLength(cell);
        this.values[cell[0] - 1][cell[1] - 1] = value;
        if (this.format != null) this.format[cell[0] - 1][cell[1] - 1] = [0, 0, 0, 0, 0, 1]

        setSSValues(this.sheet, cell, value).then(r => console.log("value at " + cell + " set as : " + value));
        return value;
    }

    checkLength(cell){
        if (this.values.length <= cell[0]){
            let valueRow = [];
            for(let i = 0; i < this.values[0].length; i++){
                valueRow.push(null)
            }
            this.values.push(valueRow);
            this.format.push(valueRow);
        }
    }

    setRichValue(cell, text, value, is_rtv=false) {
        this.checkLength(cell);
        let RTV;
        if (is_rtv){
            RTV = value;
            if (typeof text === 'number'){
                RTV = SheetClass.JSONtoRichValue(text.toString(), [[0, 0, 0, 0, "0", 1]]);
            }
        }
        else {
            RTV = SheetClass.JSONtoRichValue(text, value);
        }

        setRTValues(this.sheet, cell, RTV).then(r => console.log("RTValue at " + cell + " set as : " + text));

        this.values[cell[0] - 1][cell[1] - 1] = text;
        if (this.format != null && is_rtv) this.format[cell[0] - 1][cell[1] - 1] = SheetClass.richValueToJSON(text, RTV)[1];
        else if (this.format != null && !is_rtv) this.format[cell[0] - 1][cell[1] - 1] = value;

        return text;
    }

    static richValueToJSON(text, value){
        const runs = value.getRuns();
        const json = [text, []];
        runs.forEach(run => {
            let runTS = run.getTextStyle();
            let TS = 1;

            if (runTS.isBold()) TS *= 2;
            if (runTS.isItalic()) TS *= 3;
            if (runTS.isStrikethrough()) TS *= 5;
            if (runTS.isUnderline()) TS *= 7;

            let color = runTS.getForegroundColorObject() == null ? "0" : runTS.getForegroundColorObject().asRgbColor().asHexString();

            json[1].push([run.getStartIndex(),
                run.getEndIndex(),
                // If the run has a link
                run.getLinkUrl() == null ? 0 : run.getLinkUrl(),
                runTS.getFontFamily() == null ? 0 : runTS.getFontFamily(),
                runTS.getFontSize() == null ? 0 : runTS.getFontSize(),
                color,
                TS]
            );
        });
        return json;
    }

    static JSONtoRichValue(text, json){
        let richValue = SpreadsheetApp.newRichTextValue().setText(text);

        if (json != null) {
            json.forEach(run => {
                // If the run has a link
                if (run[2] !== 0) richValue = richValue.setLinkUrl(run[0], run[1], run[2]);

                // Checks if any text style is applied
                if (run[3] !== 0 && run[4] !== 0 && run[5] !== "0" && run[6] !== 1) {

                    let TS = SpreadsheetApp.newTextStyle();
                    // Font Family
                    if (run[3] !== 0) TS = TS.setFontFamily(run[3]);
                    // Font Size
                    if (run[4] !== 0) TS.setFontSize(run[4]);

                    richValue = richValue.setTextStyle(run[0], run[1], TS.setForegroundColor(run[5])
                        .setBold(run[6] % 2 === 0)
                        .setItalic(run[6] % 3 === 0)
                        .setStrikethrough(run[6] % 5 === 0)
                        .setUnderline(run[6] % 7 === 0)
                        .build());
                }
            });
        }
        return richValue.build();
    }
}

// Turns compressed array into the JSON used for the code
function JSONToFormat(compressed){
    let format = [];
    let values = [];
    compressed.forEach(row => {
        let formatRow = [];
        let valueRow = [];

        if (row[0] === '◊') {
            for (let i = 0; i < row[1]; i++) {
                for (let j = 0; j < row[2]; j++) {
                    valueRow.push([]);
                }
                format.push(valueRow);
                values.push(valueRow);

                formatRow = [];
                valueRow = [];
            }
        }
        else {
            row.forEach(cell => {
                if (Number.isInteger(cell)) {
                    for (let i = 0; i < cell; i++) {
                        formatRow.push(null);
                        valueRow.push(null);
                    }
                }
                else {
                    if (typeof cell === 'string') {
                        valueRow.push(cell);
                        formatRow.push([[0, 0, 0, 0, 0, 1]])
                    }
                    else {
                        valueRow.push(cell[0]);
                        formatRow.push(cell[1]);
                    }
                }
            });
            format.push(formatRow);
            values.push(valueRow);
        }
    });
    return [values, format];
}

// Turns the grabbed files into the JSON used to store
function formatToJSON(format, values, formatRTV = false){
    let grid = [];
    let rowN = 0;

    values.forEach((row, row_index) => {

        if (row.length === 0) row ++;
        else {
            if (rowN !== 0) {
                grid.push(["◊", rowN, row.length]);
                rowN = 0;
            }
            let gridRow = [];
            let columnN = 0;

            row.forEach((cell, column_index) => {
                if (cell === "" || cell === '' || cell == null) columnN++;
                else {
                    if (columnN !== 0) {
                        gridRow.push(columnN);
                        columnN = 0;
                    }

                    // Checks if the cell is a rich text value
                    if (formatRTV) {
                        if (format[row_index][column_index] !== null || format[row_index][column_index] !== "") {
                            gridRow.push(SheetClass.richValueToJSON(cell, format[row_index][column_index]));
                        }
                        // Else it pushes the value of the cell
                        else gridRow.push([cell.toString(), [[0, 0, 0, 0, 0, 1]]]);
                    }
                    else {
                        if (format[row_index][column_index] !== null || format[row_index][column_index] !== "") {
                            gridRow.push([cell.toString(), format[row_index][column_index]]);
                        }
                        // Else it pushes the value of the cell
                        else gridRow.push([cell.toString(), [[0, 0, 0, 0, 0, 1]]]);
                    }
                }
            });
            if (columnN !== 0){
                gridRow.push(columnN);
            }
            grid.push(gridRow);
        }
    });
    if (rowN !== 0) {
        grid.push([rowN, values[0].length]);
    }
    return grid;
}

async function setSSValues(sheet, cell, value) {
    sheet.getRange(cell[0], cell[1], 1, 1).setValue(value);
}

async function setRTValues(sheet, cell, value) {
    sheet.getRange(cell[0], cell[1], 1, 1).setRichTextValue(value);
}

function setConstants(config_values) {
    max = find_config_value("Maximum Search Length", config_values);
    priority_list = find_sections("Priority List", config_values);
    priority_title = find_config_value("Priority Section Title", config_values);
    auto_publish = find_config_value("Auto Publish", config_values);
    sections = find_sections("Form Section Names", config_values);
    filtered_secs = find_sections("Filtered Section Names", config_values);
    internal_sec = find_config_value("Internal Section Name", config_values);
    sheet_n = find_config_value("Internal Sheet Number Title", config_values);
    priority_exclusion = find_sections("Priority Exclusion", config_values);
    status_categories = find_sections("Internal Status Naming", config_values);
    status = status_categories[0];
    status_categories.shift();
}

function findTicketsFromSpread(scriptProperties, panel){
    const non_tickets = ss_info.getRange(2, 1, 1, 1).getValue().split("::");

    let sheets = spread.getSheets();
    let section_tickets = [];
    sheets.forEach((sheet) => {
        if (!non_tickets.includes(sheet.getName())) {
            // If the ticket is of the right ticketing system
            if (sheet.getRange(1, 1, 1, 1).getValue().includes(panel)) {
                section_tickets.push(sheet.getName());
            }
        }
    });
    console.log("Found Tickets are: " + section_tickets);
    new Promise(function (resolve, reject) {
        resolve(scriptProperties.setProperty(home_sheet + "_tickets", packProperties(section_tickets)));
    }).then(r => console.log("Tickets Uploaded"));

    return section_tickets;
}

function home_builder(scriptProperties, home, home_name, config) {
    let home_class;
    let config_values;

    // Try to grab config and home values from script properties
    try{
        // Get the unpacked values
        let home_values = unpackProperties(scriptProperties.getProperty(home_name));

        if (home_values === null) throw "No Home Values";
        config_values = home_values["config"];

        let script_values = JSONToFormat(home_values["values"]);
        let home_format = script_values[1];

        home_values = script_values[0];
        home_class = new SheetClass(home, null, home_values, home_format);
    } catch (err) {
        console.log(err);
        const home_range = home.getDataRange();
        config_values = config.getDataRange().getValues();
        let raw_values = formatToJSON(home_range.getRichTextValues(), home_range.getValues(), true);
        saveValues(scriptProperties, raw_values, config_values).then(r => console.log(r));
        console.log("JSON Values: ");
        console.log(JSON.stringify(raw_values));
        raw_values = JSONToFormat(raw_values);

        home_class = new SheetClass(home, null, raw_values[0], raw_values[1]);
    }

    return [home_class, config_values];
}

function findTickets(scriptProperties){
    let tickets;
    try {
        let tickets_data = unpackProperties(scriptProperties.getProperty(home_sheet + "_tickets"));
        if (tickets_data === null){
            tickets = findTicketsFromSpread(scriptProperties, tickets[1]);
        } else{
            tickets = tickets_data;
        }
    } catch (err) {
        console.log(err);
        tickets = findTicketsFromSpread(scriptProperties, tickets[1]);
    }

    return tickets;
}

async function saveValues(scriptProperties, values, config_values){
    scriptProperties.setProperty(home_sheet, packProperties({
        "values": values,
        "config": config_values,
    }));

    return ("200, Saved Values");
}

function packProperties(value){
    console.log("Unpacked Prop Size: " + JSON.stringify(value).length);
    return(Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(JSON.stringify(value), 'application/x-gzip')).getBytes()));
}

function unpackProperties(value){
    console.log("Packed Prop Size: " + value.length);
    return(JSON.parse(Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(value), 'application/x-gzip')).getDataAsString()));
}
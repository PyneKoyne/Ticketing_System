# Ticketing System
> Two Google Scripts to allow the communication between a Ticket Creation Form and Ticketing System Sheet

## Set-up
> To set-up the program, follow the instructions below for ringleader.js and ringfollower.js respectively

### How to find the ID of a Google Document
> The ids can be found by taking the URL of the Google document such as:
```
https://docs.google.com/forms/d/1KEgccOvLDEMFt-M1GofMHwljodrFbfQVDm1HKeAnrad/edit
```
> and taking the long string of gibberish out.

> In this case it is `1KEgccOvLDEMFt-M1GofMHwljodrFbfQVDm1HKeAnrac`. An example of the fields filled in is:
```js
const form_id = "1KEgccOvLDEMFt-M1GofMHwljodrFbfQVDm1HKeAnrad";
const ticket_spreadsheet_id = "1EdmlfNAdJIqXSJpA1LY9J1EjeHLKCnLRnlZoN5RcDie";
```

### Ringfollower.js
> Create a Google Sheet. This will be the sheet used to send tickets to.\
> ![image](https://github.com/PyneKoyne/Ticketing_System/assets/39810461/03553785-103b-40a6-b846-4c60e5e07b79)

> Click extensions and select **Apps Script**.\
> ![image](https://github.com/PyneKoyne/Ticketing_System/assets/39810461/64dcbeb7-67a5-4699-bbb7-2a62bb52f776)

> Copy-paste the code from **ringfollower.js** into the open file and the id of the Google Sheet into the following variable:
```js
const ticket_spreadsheet_id = "SPREADSHEET ID";
```

> You will need at least 3 seperate sheets, the home sheet, the config sheet, and the template sheet.

> The Home sheet requires one cell to be named `Subject`, which will be the top left cell of the ticket list
>
>> For the Priority List, write the name of the section in one cell of the sheet.
>
>> Starting one cell down and moving to the right, write the different priority categories, one per cell.
>
>> ![image](https://github.com/PyneKoyne/Ticketing_System/assets/39810461/c60b0cdb-c4a6-458c-a474-d0596e26b980)

> The Template sheet will be the format in which your Ticket Form responses are populated into the Spreadsheet
> You must have a cell for each of the sections listed in ***Form Section Names*** in the Config sheet.
> Each of these cells must have a 1 cell gap in between
> There must be a cell named what is written for ***Internal Section Name***
> One cell down and to the left, you must include the 3 following cells:
>> One for the status, in which it is named `Status:`
>> One for the Sheet Number, in which it is named what is written for ***Internal Sheet Number Title***
>> One which is simply named: `Email:`
> You must also include a Google Drawing and assign a script `publishTicketHandler`.
>![image](https://github.com/PyneKoyne/Ticketing_System/assets/39810461/3f4be783-4a62-4921-9fc4-8d9fb4dd07dd)


> The Config sheet must be named `Config`
> The first row of the sheet must have the following values (You may just simple copy the text below and paste them into your sheet):
```
Remember to Hide this sheet!	Priority List	Internal Section Name	Form Section Names	Filtered Section Names	Internal Status Naming	Current Active Tickets	Maximum Search Length	Auto Publish	Priority Section Title	Current Priority Tickets	Total Saved Tickets	Internal Sheet Number Title	Non-Ticket Sheets
	(From Highest Prioity to Lowest Prioity)			(Which section will be shown on the home page?)									
					Status:	0	50	FALSE		0	0		Home
								 		0			Template
										0			Config
										0			
										0			
										0			
										0			
										0			
										0			
										0			
										0			
										0			
										0			
										0			
										0			
										0			
										0			
										0			
										0			
										0			
										0			
```
> 2 cells under ***Prioity List***, list the priority options listed within the Ticketing Form\
> ![image](https://github.com/PyneKoyne/Ticketing_System/assets/39810461/8cb9ed02-8b53-4491-8e80-1b92845db2ec)

> 2 cells under ***Internal Section Name***, write the name of the internal ticket data section you want for each ticket.\
> ![image](https://github.com/PyneKoyne/Ticketing_System/assets/39810461/b5ed747f-7d5c-40b1-87e3-c32a9d2bfbb4)

> 2 cells under ***Form Section Names***, write the name of each section from the form on a new cell (The names do not have to match up with the form).\
> ![image](https://github.com/PyneKoyne/Ticketing_System/assets/39810461/9d81ebe1-06c7-4958-a469-d715e910d09b)

> 2 cells under ***Filtered Section Names***, write the name of each section from ***Form Section Names*** you want to display on the home page.\
>![image](https://github.com/PyneKoyne/Ticketing_System/assets/39810461/4efff26b-38cc-4fd3-b72d-0c919e0d8ef6)

> 2 cells under ***Internal Status Naming***, write the name you want for the status category, followed by the different possible statuses.
> ![image](https://github.com/PyneKoyne/Ticketing_System/assets/39810461/4a790b65-a04d-4f78-9c76-73f3545456db)

> 2 cells under ***Priority Section Title***, write the title of the priority section on the home page.\
> ![image](https://github.com/PyneKoyne/Ticketing_System/assets/39810461/c75bad10-a4e7-4e1c-b106-cfd901c21a9f)

> 2 cells under ***Internal Sheet Number Title***, write the name for the `Sheet Number` category on the Template Sheet.\
> ![image](https://github.com/PyneKoyne/Ticketing_System/assets/39810461/2d6fac4a-69d0-4a73-917f-01e2e2878065)

> For ***Non-Ticket Sheets***, add any sheets that are not tickets in empty cells below.\
> ![image](https://github.com/PyneKoyne/Ticketing_System/assets/39810461/cd232846-9d11-4461-846b-2d29e96efc35)


### Ringleader.js

> Create a Google Form. This will be the form that will be used to submit tickets to the ticketing system.\
> ![image](https://github.com/PyneKoyne/Ticketing_System/assets/39810461/e076b940-ddc0-4b34-a585-277256e33655)

> Click the 3 dots next to your profile icon on the top right. Select "**Script editor**"\
> ![image](https://github.com/PyneKoyne/Ticketing_System/assets/39810461/60d8a64e-2881-4df5-9d89-07f99042a82b)

> Copy-paste the code from **ringleader.js** into the open file and the ids of your Google Form and Google Sheet into the following variables:

```js
const form_id = "FORM ID";
const ticket_spreadsheet_id = "SPREADSHEET ID";
```

> Then make sure you select `setUpTrigger`\
> ![image](https://github.com/PyneKoyne/Ticketing_System/assets/39810461/cdca9c68-548c-49ee-92d1-cd8526444516)

> Now press **Run** and allow all permissions. If there is a warning, select advanced and press `Go to '' (unsafe)`.\
> ![image](https://github.com/PyneKoyne/Ticketing_System/assets/39810461/09ecb2b7-7985-4c22-9dfa-e648c587e98a)

Your Google Form Script should now be completely set up!

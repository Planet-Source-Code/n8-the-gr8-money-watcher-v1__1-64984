Money Watcher v.1

*********** Many options are configurable via options.ini *************

Money Watcher v.1 is a Visual Basic 6.0/MS Access application designed to assist with tracking withdrawals/deposits 
to your checkings and savings accounts.  Several tools are included to aid you with tracking your expedinatures
and visualize trends and utilize budgets.  

* This tool was designed to reside in the quick launch portion of MS Windows for easy access.

* The options.ini file must reside in the same directory as the file!!!!

Header Form (Main Form)

- Lists all transactions to/from your checkings/savings accounts
- Allows for real time SQL statments to access various information from the primary Access table
- Provides the date, number of records in database, available amount, posted amount, and reminders in status bar
- Buttons include insert, remove, refresh, search, and a multi-function button

* Notes

- Double click the horizontal line near the top of the form to minimize the form.
- Click and drag horizontal line near the top to move the form (applies to all forms)
- Double click the reminders indicator in the status bar to quickly access the reminders form
- SQL statments can be saved using the '+' button in the SQL entry box, use the < > buttons to 
  scroll through various saved SQL statments
- Using a "=" before your SQL statement will return a value in the x1 field (see included SQL examples),  
  ommitting the '=' will attempt to populate the flexgrid with whatever is returned from the SQL statment.
  Multiple values can be returned, just sequentially name your fields X1, X2, X3, X4, ect....
- Right click a record to show a popup menu, allowing you to copy the record or post record


Backup Database

- Simply copies database information to a backup directory controlled by the options.ini file


Bank Website

- Calls a website (to your bank) controlled by the options.ini file


Datapad

- Stores text/url information into datapad.dat, all information is automatically stored upon exiting the datapad form

* Notes

- Doubleclick urls to access them via your default web browser


Expense Pie Chart

- Displays a pie chart detailing withdrawals.  

* Notes

- Single click an item in the list box to highlight to corresponding pie piece
- Double click an item in the listbox to add/remove it from the calculations - all are checked by default


History Graph

- Allows for historical data at various intervals.  Choose a category and a timeframe click the execute button.

* Notes

- The show markers checkbox will apply markers to the coordinates
- The show 0 values checkbox  will show all days regardless of information being recorded for that day/week/month/year
- Move mouse over coordinates to display label and value, this is useful when many coordinates are displayed
  making the fntsize of the labels too small to see clearly

Budget

- Using the budgets assigned when accessing the detail records, this form details budget status and amounts

* Red colored items refer to items over budget / Green for under budget

Payments

- Shows payments for various months

* Payments are created when you create a withdrawal,  you'll be prompted if you want to use the withdrawal as a payment.
  The information is then stored in the options.ini

Reminders

- Create reminders at various repeating intervals
- Can update MS Outlook Calendar entries - controlled by options.ini 
* See calendar.gif,  the current setup reads a custom calendar named within the options.ini file (calendar_name)






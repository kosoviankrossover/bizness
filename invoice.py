# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
# example use: python code/invoice.py September 2019
# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------

# LIBRARIES

import tkinter as tk
import numpy as np
import pandas as pd
import os

from docxtpl import DocxTemplate
from docx import Document
import datetime
from docxcompose.composer import Composer
from sys import argv

# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
# COMMAND LINE INPUTS

# check command line args for month and year of invoice
if len(argv) > 2:
    MONTH = argv[1] # first argument is the month
    YEAR = argv[2] # second argument is the year
else:
    exit('Missing MONTH and/or YEAR argument.')

# month set
month_set = set(['January', 'February', 'March', 'April', 'May', 'June', 'July',
              'August', 'September', 'October', 'November', 'December'])
# check for proper month
if MONTH not in month_set:
    exit('Provide a proper month.')

# try for proper dates
try:
    CURR_DATE = datetime.datetime.strptime(MONTH+' ' + YEAR, '%B %Y')
except ValueError: # else exit
    str_exit =  'Improper command line inputs.\n'
    str_exit += 'Proper input looks like:\n'
    str_exit += 'python code/invoice.py September 2019'
    exit(str_exit)

### needed template files
# docs
invoice_template = 'templates/invoice_template.docx'
invoice_template_mult = 'templates/invoice_template_mult.docx' #goes with RPind
invoice_template_z = 'templates/invoice_template_z.docx' #goes with Zind
invoice_template_v = 'templates/invoice_template_v.docx' #goes with Vind

# data
invoice_data_template = 'templates/invoice_data_template.xlsx'

# check for needed files
if not os.path.exists(invoice_template): # word template
    exit('Missing invoice word template.')
if not os.path.exists(invoice_template_mult): # word template mult
    exit('Missing invoice word template mult.')
if not os.path.exists(invoice_template_z): # word template z
    exit('Missing invoice word template z.')
if not os.path.exists(invoice_template_v): # word template v
    exit('Missing invoice word template v.')
if not os.path.exists(invoice_data_template): # data template
    exit('Missing invoice data template.')

# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
# IMPORT and AUGMENT DATA

# column list
col_list = ['OWNER', 'FIRST', 'LAST',
            'STREET_ADDRESS', 'CITY_ADDRESS',
            'MY_NOTES',
            'MONTHLY_CHARGE',
            'EMAIL_ADDRESS',
            'COND_CHARGE', 'FILT_CHARGE',
            'COND_MONTHS', 'FILT_MONTHS',
            'RP_INDICATOR',
            'Z_INDICATOR', 'V_INDICATOR']
# data type dictionary
types = {'OWNER':str, 'FIRST':str, 'LAST':str,
        'STREET_ADDRESS':str, 'CITY_ADDRESS':str,
        'MY_NOTES':str,
        'MONTHLY_CHARGE':float,
        'EMAIL_ADDRESS':str,
        'COND_CHARGE':float, 'FILT_CHARGE':float,
        'COND_MONTHS':str, 'FILT_MONTHS':str,
        'RP_INDICATOR':int,
        'Z_INDICATOR':int, 'V_INDICATOR':int}

# read in data
tabl = pd.read_excel(invoice_data_template,
                     usecols= col_list,
                     dtype=types,
                     skiprows=1)
# fix email improper reading in of NA string values
tabl['EMAIL_ADDRESS'] = tabl['EMAIL_ADDRESS'].astype(str) #'NA' is converted to 'nan'
# fill in any missing data for string variables
# fill in with a blank ''
na_list = ['FIRST', 'LAST', 'STREET_ADDRESS', 'CITY_ADDRESS', 'MY_NOTES']
for col in na_list:
    tabl[col].fillna(value='', inplace=True)

# number of rows
N = tabl.shape[0]

# add new columns
tabl['ADD_CHARGE'] = np.zeros(N, dtype=float)
tabl['ADD_CHARGE_NOTES'] = ['']*N
tabl['CUST_REMINDER'] = ['']*N

# indicator for whether to include rows or not for the outgoing files
included = np.ones(N, dtype=int)

# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
## CREATE DIRECTORY
print("Creating directories...")

### month dict: month name -> numerical two digit code
month_dict = {'January':'01', 'February':'02', 'March':'03', 'April':'04',
                'May':'05', 'June':'06', 'July':'07', 'August':'08',
                'September':'09', 'October':'10', 'November':'11', 'December':'12'}

# make directory
# example: invoices/2019_09_September
# this allows invoices to be properly sorted numerically
# -- e.g. 2019_08_August is on top of 2019_09_September
# -- and 2018_12_December is on top of 2019_01_January
# -- and 2018_09_September is above 2019_09_September
path = 'invoices/'+YEAR+'_'+month_dict[MONTH]+'_'+MONTH
try:
    os.mkdir(path)
except OSError:
    exit("Creation of the directory %s failed" % path)

# make "emails" directory
# example: invoices/2019_09_September/emails
emails_path = path+'/emails'
try:
    os.mkdir(emails_path)
except OSError:
    exit("Creation of the directory %s failed" % emails_path)



# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
# GLOBAL VARS and CONSTANTS

##### global vars
# current row
index = 0
# indicator for whether or not the 'back button' is on in the gui
back_on = False
# indicator for restoring the main gui format upon going back to the main...
# ...program from the exit screen gui format
last_back = False
# error string to attach if we receive improper input
error_str = ''
# back string to infrom user we went back a row
back_str = ''
# indicator for whether we have finished the main invoicing program
final_end = False
# dates error indicator
dates_err_ind = False
# reminder notice indicator
remind_ind = False
# row counter, column counter
row_inc = 0
col_inc = 0
##### last tow of the gui format
# main->0; month charge->1; add charge->2; notes add char->3; notes->4;
# cust remind->5; check->6; button one->7; button two->7
last_row = 7

##### constants for tkinter display windows
font = 'Times' # font type
bold_ind = 'bold' # bold indicator
size = 18 # font size
col1 = '#14bcfe' #color-> blue-ish
col2 = '#fe9114' #color -> orange-ish
notes_color = 'black' # notes color
reminder_color = 'purple' ## '#cf2b29';'#e34240'; red-ish
amt_color = 'green' # amount color
add_color = '#0a4cbf' # #blue-ish
cust_remind_color = 'red' # cust reminder color
thick = 2 # highlighted thickness
pad = 6 # amount of padding
width = 1000 # width dim -> of the entire window
height = 800 # height dim -> of the entire window
entry_width = 80 # width dim of the user input sections
on_color = 'red' # color of checked button when on
off_color = 'grey' # color of checked button when off

# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
# WINDOW FUNCTIONS

##### center the window upon opening with a given width and height
def center_window(base, width=300, height=200):
    # get screen width and height
    screen_width = base.winfo_screenwidth()
    screen_height = base.winfo_screenheight()
    # calculate position x and y coordinates
    x = (screen_width/2) - (width/2)
    y = (screen_height/2) - (height/2)
    base.geometry('%dx%d+%d+%d' % (width, height, x, y))
    return

##### close root i.e. the main program
def close_root():
    # global vars
    global root
    global final_end

    # destroy root
    root.destroy()
    # set checkpoint
    final_end = True
    return

# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
# HELPER FUNCTIONS for navigating

# check value of checkbutton
# - if its clicked, then turn it on to appropriate color
# - else leave it deafault black color
def on_check():
    global var, chk
    if var.get() == 1:
        chk["fg"] = on_color
    else:
        chk["fg"] = off_color
    return

##### function to begin the main program
def begin():
    # global vars
    global btn_one, btn_two
    global monthly_charge_lbl, monthly_charge_entry
    global add_charge_lbl, add_charge_entry
    global add_charge_notes_lbl, add_charge_notes_entry
    global my_notes_lbl, my_notes_entry
    global cust_reminder_lbl, cust_reminder_entry
    global chk
    global row_inc, col_inc

    ##### main label
    main_lbl.grid_forget()

    ##### first button
    # change text: begin -> next
    # change command
    # if clicked, we move to the next row
    # for proper display, we forget it so that we can place it at the end later
    btn_one.grid_forget()
    btn_one.config(text= 'Next')
    btn_one.config(command= next)

    ##### second button
    # change text: close -> back
    # change command
    # if clicked, we move back to the previous row
    # for proper display, we forget it so that place it at the end later
    btn_two.grid_forget()
    btn_two.config(text= 'Back')
    btn_two.config(command= back)

    ##### pack the labels and entries
    # main label
    main_lbl.grid(row= row_inc, column= col_inc, pady= pad, columnspan=2)
    row_inc += 1
    # monthly charge
    monthly_charge_lbl.grid(row= row_inc, column= col_inc, pady= pad)
    monthly_charge_entry.grid(row= row_inc, column= col_inc+1, padx= pad)
    row_inc += 1
    # additional charge
    add_charge_lbl.grid(row= row_inc, column= col_inc, pady= pad)
    add_charge_entry.grid(row= row_inc, column= col_inc+1, padx= pad)
    row_inc += 1
    # notes on additional charges
    add_charge_notes_lbl.grid(row= row_inc, column= col_inc, pady= pad)
    add_charge_notes_entry .grid(row= row_inc, column= col_inc+1, padx= pad)
    row_inc += 1
    # my notes
    my_notes_lbl.grid(row= row_inc, column= col_inc, pady= pad)
    my_notes_entry.grid(row= row_inc, column= col_inc+1, padx= pad)
    row_inc += 1
    # customer reminders
    cust_reminder_lbl.grid(row= row_inc, column= col_inc, pady= pad)
    cust_reminder_entry.grid(row= row_inc, column= col_inc+1, padx= pad)
    row_inc += 1
    # check button
    chk.grid(row= row_inc, column= col_inc, pady= pad)
    row_inc += 1

    ##### pack back in the next button, but not the back button just yet
    # - turn the back button on at the second row
    btn_one.grid(row= row_inc, column= col_inc, pady= pad)

    # revert row increment
    row_inc = 0

    ##### update the display
    update()
    return

##### update the gui display
def update():
    # global vars
    global index
    global error_str, error_lbl
    global back_on, last_back, back_str
    global btn_one,  btn_two
    global var, chk
    global main_lbl
    global monthly_charge_lbl, monthly_charge_entry
    global add_charge_lbl, add_charge_entry
    global add_charge_notes_lbl, add_charge_notes_entry
    global my_notes_lbl, my_notes_entry
    global cust_reminder_lbl, cust_reminder_entry
    global CURR_DATE, MONTH
    global dates_err_ind, remind_ind
    global row_inc, col_inc

    ##### clear the entries
    monthly_charge_entry.delete(0, 'end')
    add_charge_entry.delete(0, 'end')
    add_charge_notes_entry.delete(0, 'end')
    my_notes_entry.delete(0, 'end')
    cust_reminder_entry.delete(0, 'end')
    ##### reset checkbox
    var.set(0) # value -> set it back to zero ie off
    chk["fg"] = off_color # color -> set back to off color, not clicked

    ##### error string and label logic
    if error_str != '': # add in the error label and message
        error_lbl.config(text= error_str)
        error_lbl.grid(row= last_row+1, column= col_inc, pady = 0, columnspan=2)
    else:
        # remove the error lbl if present before
        # if error was not present before, this does nothing
        error_lbl.grid_forget()

    ##### back button logic
    # --------------------------------------------------------------------------
    # add the 'back button' if index > 0; we don't want a back button for the
    # first row becuase there is nothing to go back to; close the window to quit
    if index == 1 and back_on == False:
        # pack the back button
        btn_two.grid(row= last_row, column= col_inc+1, padx= pad)
        # update conditional that the 'back button' is on
        back_on = True
    # --------------------------------------------------------------------------
    # remove back button if on index=0 i.e. we got back to the first row and
    # need to remove the back button because there is nothing to go back to
    if index == 0 and back_on == True:
        btn_two.grid_forget() # remove button
        back_on = False # update conditional
    # --------------------------------------------------------------------------
    # going back from the exit - last - screen
    if last_back == True:
        # we want to restore the main gui format
        # -- i.e. the labels and entries followed by the buttons

        ##### forget the buttons
        btn_one.grid_forget()
        btn_two.grid_forget()

        ##### bring back the labels and entries
        # increment row bc we do not forget the main label
        row_inc += 1
        # monthly charge
        monthly_charge_lbl.grid(row= row_inc, column= col_inc, pady= pad)
        monthly_charge_entry.grid(row= row_inc, column= col_inc+1, padx= pad)
        row_inc += 1
        # additional charge
        add_charge_lbl.grid(row= row_inc, column= col_inc, pady= pad)
        add_charge_entry.grid(row= row_inc, column= col_inc+1, padx= pad)
        row_inc += 1
        # additional charge notes
        add_charge_notes_lbl.grid(row= row_inc, column= col_inc, pady= pad)
        add_charge_notes_entry.grid(row= row_inc, column= col_inc+1, padx= pad)
        row_inc += 1
        # my notes
        my_notes_lbl.grid(row= row_inc, column= col_inc, pady= pad)
        my_notes_entry.grid(row= row_inc, column= col_inc+1, padx= pad)
        row_inc += 1
        # customer reminders
        cust_reminder_lbl.grid(row= row_inc, column= col_inc, pady= pad)
        cust_reminder_entry.grid(row= row_inc, column= col_inc+1, padx= pad)
        row_inc += 1
        # check button
        chk.grid(row= row_inc, column= col_inc, pady= pad)
        row_inc += 1

        ##### bring back the buttons
        btn_one.grid(row= row_inc, column= col_inc, pady= pad)
        btn_two.grid(row= row_inc, column= col_inc+1, padx= pad)
        # reconfig the first button i.e. the begin button at start, close at end
        btn_one.config(text= 'Next') # update text
        btn_one.config(command= next) # update command

        ###### update last_back conditional i.e. if we press the back button
        # we are no longer on the end screen, coming back from the exit screen
        last_back = False
    ##### END ----------------------------------------------- back button logic

    ##### update logic
    # check if we can update
    # if no more rows left, end
    # if we can update, replace entries with appropriate text
    if index > N-1:
        end() # end screen logic
    else:
        # grab current line
        curr_line = dict(tabl.iloc[index])

        ##### input entries
        ## main label text
        text = ''
        # inform user that we've come back to this row
        if back_str != '':
            text += back_str
            # revert back string
            back_str = ''
        # main info
        text += curr_line['FIRST']+' '+ curr_line['LAST'] + '\n\n' + \
                curr_line['STREET_ADDRESS'] + '\n' + \
                curr_line['CITY_ADDRESS']

        ### ADD TEXT - conditional
        ## cond
        if curr_line['COND_MONTHS'] != 'exclude':
            pot_cond_months = set("".join(curr_line['COND_MONTHS'].split(',')).split())
            ## - proper months
            if len(pot_cond_months - month_set) == 0:
                # check if current month is in potential cond... months
                if MONTH in pot_cond_months:
                    remind_ind = True
                    text += '\n\n' + \
                    "Remember, it's time to charge for cond...." + \
                    ' - $' + str(curr_line['COND_CHARGE'])

            ## - improper months
            else:
                dates_err_ind = True
                text += '\n\n' + "Error, improper cond... month"
        ## filt
        if curr_line['FILT_MONTHS'] != 'exclude':
            pot_filt_months = set("".join(curr_line['FILT_MONTHS'].split(',')).split())
            ## - proper months
            if len(pot_filt_months - month_set) == 0:
                # check if current month is in potential filt... months
                if MONTH in pot_filt_months:
                    ## double or single new line
                    if remind_ind or dates_err_ind:
                        text += '\n' + \
                        "Remember, it's time to charge for filt...." + \
                        ' - $' + str(curr_line['FILT_CHARGE'])
                    else:
                        text += '\n\n' + \
                        "Remember, it's time to charge for filt...." + \
                        ' - $' + str(curr_line['FILT_CHARGE'])
                    remind_ind = True
            ## - improper months
            else:
                if remind_ind or dates_err_ind:
                    text += '\n' + "Error, improper filt... month"
                else:
                    text += '\n\n' + "Error, improper filt... month"
                dates_err_ind = True

        ##### update the main label
        # color
        if dates_err_ind == True: # first indicator, red if any error
            main_lbl.config(fg= 'red')
        elif remind_ind == True: # if no error, then can set reminder
            main_lbl.config(fg= reminder_color)
        else: # revert to default, black color
            main_lbl.config(fg= 'black')
        # text
        main_lbl.config(text=text)
        ##### revert indicators
        dates_err_ind = False
        remind_ind = False

        ##### update entries
        ## monthly charge entry
        monthly_charge_entry.insert(0, curr_line['MONTHLY_CHARGE'])
        ## my notes
        my_notes_entry.insert(0, curr_line['MY_NOTES'])

        return

##### ending the program
def end():
    # global vars
    global main_lbl
    global btn_one
    global monthly_charge_lbl, monthly_charge_entry
    global add_charge_lbl, add_charge_entry
    global add_charge_notes_lbl, add_charge_notes_entry
    global my_notes_lbl, my_notes_entry
    global cust_reminder_lbl, cust_reminder_entry
    global chk
    global last_back

    ##### updating labels and buttons
    ## update the text in the main label and its color
    main_lbl.config(text= "You're done.")
    main_lbl.config(fg= 'black')

    ## update button one - originally the begin button then the next button
    btn_one.config(text= 'Close') # update the text of the button
    btn_one.config(command= close_root) # update the button's command
    # no need to change affect of the back button

    ##### forget the other labels and entries
    # monthly charge
    monthly_charge_lbl.grid_forget()
    monthly_charge_entry.grid_forget()
    # additonal charge
    add_charge_lbl.grid_forget()
    add_charge_entry.grid_forget()
    # additional charge notes
    add_charge_notes_lbl.grid_forget()
    add_charge_notes_entry.grid_forget()
    # my notes
    my_notes_lbl.grid_forget()
    my_notes_entry.grid_forget()
    # customer reminder
    cust_reminder_lbl.grid_forget()
    cust_reminder_entry.grid_forget()
    # check button
    chk.grid_forget()

    # going back from the end
    last_back = True # update the indicator to restore main gui format
    return

##### next - go to the next row
def next():
    # global vars
    global index, error_str, var
    global monthly_charge_entry
    global add_charge_entry, add_charge_notes_entry
    global my_notes_entry, cust_reminder_entry

    # grab and try reading in input
    try:
        # monthly charge input
        monthly_in_amt_str = monthly_charge_entry.get().strip()
        if monthly_in_amt_str == '':
            monthly_in_amt_str = '0'
        monthly_in_amt = float(eval(monthly_in_amt_str))
        # additional charge input
        add_in_amt_str = add_charge_entry.get().strip()
        if add_in_amt_str == '':
            add_in_amt_str = '0'
        add_in_amt = float(eval(add_in_amt_str))
        # additional charge notes input
        in_add_charge_notes = str(add_charge_notes_entry.get().strip())
        # my_notes
        in_my_notes = str(my_notes_entry.get().strip())
        # cust_reminder
        in_cust_reminder = str(cust_reminder_entry.get().strip())

        # update data
        tabl.at[index, 'MONTHLY_CHARGE'] = monthly_in_amt
        tabl.at[index, 'ADD_CHARGE'] = add_in_amt
        tabl.at[index, 'ADD_CHARGE_NOTES'] = in_add_charge_notes
        tabl.at[index, 'MY_NOTES'] = in_my_notes
        tabl.at[index, 'CUST_REMINDER'] = in_cust_reminder

        ##### checkbox - included update
        if var.get() == 1:
            included[index] = 0

        ##### clear the error string
        # ie we have not run into an error at this point
        error_str = ''

        ##### update index -- bc we have no error
        index += 1

    # reached an error
    except NameError:
        # we have an error, so update error string
        # and do not move forward with index
        error_str = '\n\n' + \
        'Previous error. Make sure amounts are valid real numbers.'

    ##### update
    # -- if we had no error, we update to the next row
    # -- if we had an error, we re-update the current row with an error message
    update()
    return

##### next - go to the next row
def back():
    # global vars
    global index, back_str, error_str

    # revert index
    index -= 1

    # make sure we revert the included ariable indicator just in case
    included[index] = 1

    # update back string
    back_str = 'WENT BACK - '

    # update error string - in case we were at an error
    error_str = ''

    # we went back, so now update the main gui display
    update()
    return

# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
# SET-UP the gui

##### root for gui
root = tk.Tk()
root.title('Invoice: '+MONTH+' '+YEAR)
#root.configure(background='grey')

##### center the window
center_window(root, width, height)

###### main label where we will show input
main_lbl = tk.Label(root, text= 'Click to begin.')
main_lbl.config(font= (font, size, bold_ind))
main_lbl.config(fg= 'black')

main_lbl.grid(row=0, column=0, pady=pad, columnspan=2)

##### error label
error_lbl = tk.Label(root, text= error_str)
error_lbl.config(font= (font, size, bold_ind))
error_lbl.config(fg= 'red')

##### the button to begin -> next -> end program
btn_one = tk.Button(root, text= 'Begin', command= begin,
                highlightbackground= col1,
                highlightthickness= thick)
btn_one.config(font= (font, size, bold_ind))
btn_one.grid(row=1, column=0, pady=pad)

##### the button to close -> back
btn_two = tk.Button(root, text= 'Close', command= close_root,
                    highlightbackground= col2,
                    highlightthickness= thick)
btn_two.config(font= (font, size, bold_ind))
btn_two.grid(row=1, column=1, pady=pad)

##### monthly charge
## label
monthly_charge_lbl = tk.Label(root, text= 'Monthly Charge:')
monthly_charge_lbl.config(font= (font, size, bold_ind))
monthly_charge_lbl.config(fg= amt_color)
## entry
monthly_charge_entry = tk.Entry(root, width= entry_width)
monthly_charge_entry.config(font= (font, size))

##### additional charge
## label
add_charge_lbl = tk.Label(root, text= 'Additional Charge:')
add_charge_lbl.config(font= (font, size, bold_ind))
add_charge_lbl.config(fg= add_color)
## entry
add_charge_entry = tk.Entry(root, width= entry_width)
add_charge_entry.config(font= (font, size))

##### additional charge notes
## label
add_charge_notes_lbl = tk.Label(root, text= 'Notes on additional charge:')
add_charge_notes_lbl.config(font= (font, size, bold_ind))
add_charge_notes_lbl.config(fg= add_color)
## entry
add_charge_notes_entry = tk.Entry(root, width= entry_width)
add_charge_notes_entry.config(font= (font, size))

##### my notes
## label
my_notes_lbl = tk.Label(root, text= 'My notes:')
my_notes_lbl.config(font= (font, size, bold_ind))
my_notes_lbl.config(fg= notes_color)
## entry
my_notes_entry = tk.Entry(root, width= entry_width)
my_notes_entry.config(font= (font, size))

##### reminders
## label
cust_reminder_lbl = tk.Label(root, text= 'Set Customer Reminder:')
cust_reminder_lbl.config(font= (font, size, bold_ind))
cust_reminder_lbl.config(fg= cust_remind_color)
## entry
cust_reminder_entry = tk.Entry(root, width= entry_width)
cust_reminder_entry.config(font= (font, size))

##### include check button
var = tk.IntVar()
chk = tk.Checkbutton(root, text='Do Not Include?', variable=var,
            selectcolor= 'blue', command= on_check, fg= off_color)
chk.config(font= (font, size, bold_ind))
# var.get()

##### main loop
root.mainloop()

# check if we completed the main program all the way through or not
# - if we did not, then exit with an error
if final_end == False:
    exit('Failed to complete full program.')

# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
# INITIALIZE DOC WRITING

# ADD DATA i.e. total column, and current date for output doc
# create total column
tabl['TOTAL'] = tabl['MONTHLY_CHARGE'] + tabl['ADD_CHARGE']
# add included
tabl['INCLUDED'] = included

# current date
DATE = datetime.datetime.today().strftime('%B %d, %Y')

# FIND FIRST ROW TO START
where_filt = np.where(included==1)[0] # first non-skip row
if where_filt.size == 0:
    exit('Error in finding beginning row.')
start = where_filt[0]

# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
# HELPER FUNCTION to grab appropriate data

# grab a row from the table and process it for data to place in doc
def grab_data(row):
    # global data
    global data

    # convert to dict
    data = dict(tabl.iloc[row])
    # add dates
    data['DATE'] = DATE
    data['MONTH'] = MONTH
    data['YEAR'] = YEAR
    # if no additional charge then show nothing, else format float
    if data['ADD_CHARGE'] == 0:
        data['ADD_CHARGE'] = ''
    else:
        data['ADD_CHARGE'] = '${:,.2f}'.format(data['ADD_CHARGE'])
    # format floats
    data['MONTHLY_CHARGE'] = '${:,.2f}'.format(data['MONTHLY_CHARGE'])
    data['TOTAL'] = '${:,.2f}'.format(data['TOTAL'])
    return

# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
print("Writing files....")

### first file
# grab data and process it
data = None # declare global var first
grab_data(start) # grab data from first doc

# create first doc
doc = DocxTemplate(invoice_template)
doc.render(data)

# create invoice file and save
invoice_doc = path+'/invoice_'+MONTH+'_'+YEAR+'.docx'
doc.save(invoice_doc)

# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
# CONCATENATE DOC through loop

##### 'rp' indicator and values --- specific use
rp_inc = 0 # increment
rp_ind = False # rp process completion indicator
rp_months = []
rp_adds = []
rp_notes = []
#####

# loop
for i in range(start+1,N):
    # progress print
    print(i)

    # do not include the current row
    if included[i] == 0:
        continue # skip it

    ##### create temp dox

    # grab data and process it
    grab_data(i)

    ### temp doc
    # rp conditionals
    if tabl['RP_INDICATOR'][i] == 1:
        # increment
        rp_inc += 1
        # have we seen three?
        if rp_inc == 3:
            # first append
            rp_months.append(data['MONTHLY_CHARGE'])
            rp_adds.append(data['ADD_CHARGE'])
            rp_notes.append(data['ADD_CHARGE_NOTES'])
            ### new data
            rm = data['CUST_REMINDER'] # cust reminder
            # save data for processing
            data = {}
            data['DATE'] = DATE
            data['MONTH'] = MONTH
            data['YEAR'] = YEAR
            data['M1'] = rp_months[0]
            data['M2'] = rp_months[1]
            data['M3'] = rp_months[2]
            data['A1'] = rp_adds[0]
            data['A2'] = rp_adds[1]
            data['A3'] = rp_adds[2]
            data['A1_NOTES'] = rp_notes[0]
            data['A2_NOTES'] = rp_notes[1]
            data['A3_NOTES'] = rp_notes[2]
            data['CUST_REMINDER'] = rm
            ### loop for total sum
            total_sum = 0
            # month charge
            for ele in rp_months:
                if ele != '':
                    total_sum += float(ele.strip('$'))
            # add charge
            for ele in rp_adds:
                if ele != '':
                    total_sum += float(ele.strip('$'))
            data['TOTAL'] = '$'+str(total_sum)
            rp_ind = True # finished the 'rp' process, note it
        else:
            # nothing to do yet but to append the data
            rp_months.append(data['MONTHLY_CHARGE'])
            rp_adds.append(data['ADD_CHARGE'])
            rp_notes.append(data['ADD_CHARGE_NOTES'])
            continue

    ### temp doc
    if rp_ind: # rp stuff
        temp = DocxTemplate(invoice_template_mult)
        rp_ind = False # revert rp process completion indicator for next round
    elif tabl['Z_INDICATOR'][i] == 1: # Z stuff
        temp = DocxTemplate(invoice_template_z)
    elif tabl['V_INDICATOR'][i] == 1: # v stuff
        temp = DocxTemplate(invoice_template_v)
    else: # default
        temp = DocxTemplate(invoice_template)
    # render the doc
    temp.render(data)

    # if no email indicator, merge to paper invoice
    if tabl['EMAIL_ADDRESS'][i].strip() == 'nan':
        ### merge docs
        # load master doc
        doc = Document(invoice_doc)
        # add page break
        doc.add_page_break()
        # compose it
        composer = Composer(doc)
        # merge it with the temp doc
        composer.append(temp)

        # save the final output
        composer.save(invoice_doc)
    else:
        # path
        temp_emails_path = emails_path+'/'+tabl['EMAIL_ADDRESS'][i].strip()
        # make directory
        os.mkdir(temp_emails_path)
        # temp email paths word doc file name
        temp_emails_path_doc = temp_emails_path+'/invoice_'+MONTH+'_'
        temp_emails_path_doc += '_'.join(tabl['FIRST'][i].strip().lower().split())
        temp_emails_path_doc += '_'
        temp_emails_path_doc += tabl['LAST'][i].strip().lower()
        temp_emails_path_doc += '.docx'
        # save doc
        temp.save(temp_emails_path_doc)

# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
# CREATE EXCEL SHEET for output
# drop columns not needed
tabl.drop(['OWNER', 'EMAIL_ADDRESS',
            'COND_CHARGE', 'FILT_CHARGE',
            'COND_MONTHS', 'FILT_MONTHS',
            'RP_INDICATOR',
            'Z_INDICATOR', 'V_INDICATOR'], axis=1, inplace=True)

# taken from:
# https://xlsxwriter.readthedocs.io/example_pandas_column_formats.html
# Create a Pandas Excel writer using XlsxWriter as the engine
invoice_data = path+'/data_'+MONTH+'_'+YEAR+'.xlsx' # file path
writer = pd.ExcelWriter(invoice_data, engine='xlsxwriter')
# Convert the dataframe to an XlsxWriter Excel object.
tabl.to_excel(writer, sheet_name='Sheet1', index= False)
# Get the xlsxwriter workbook and worksheet objects.
workbook  = writer.book
worksheet = writer.sheets['Sheet1']
# Set the column width and format.
worksheet.set_column('A:B', 18, None) # first, last
worksheet.set_column('C:D', 18*1.5, None) # street, city
worksheet.set_column('E:E', 18*2.5, None) # my notes
worksheet.set_column('F:F', 18, None) # monthly charge
worksheet.set_column('G:G', 18, None) # additional charge
worksheet.set_column('H:H', 18*2.5, None) # add charge notes
worksheet.set_column('I:I', 18*1.5, None) # cust reminder
worksheet.set_column('J:K', 18, None) # total, included
# Close the Pandas Excel writer and output the Excel file.
writer.save()

# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
# CREATE EXCEL SHEET for shorter output
# 'short' path
short_path = path+'/short_data_'+MONTH+'_'+YEAR+'.xlsx'
# create shorter data
short_tabl = tabl[['LAST','STREET_ADDRESS','ADD_CHARGE_NOTES']].copy()
# add month service column
short_tabl[MONTH+' '+YEAR] = tabl['MONTHLY_CHARGE'].copy()
# add recrods column
short_tabl['Records'] = ['']*N

# create excel output
short_writer = pd.ExcelWriter(short_path, engine='xlsxwriter')
# Convert the dataframe to an XlsxWriter Excel object.
short_tabl.to_excel(short_writer, sheet_name='Sheet1', index= False)
# Get the xlsxwriter workbook and worksheet objects.
short_workbook  = short_writer.book
short_workbook = short_writer.sheets['Sheet1']
# Set the column width and format.
short_workbook.set_column('A:A', 15, None) # last
short_workbook.set_column('B:B', 25, None) # street
short_workbook.set_column('C:C', 40, None) # add charge notes
short_workbook.set_column('D:D', 15, None) # monthly service
short_workbook.set_column('E:E', 15, None) # records
# MARGINS
short_workbook.set_margins(left=0.1, right=0.1, top=0.1, bottom=0.1)
# Close the Pandas Excel writer and output the Excel file.
short_writer.save()

# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
## main program done
print('Main program done.')

# total total sum
print('Total: ${:.2f}'.format(tabl['TOTAL'].values.sum()))

# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
# ## RUN AUX PROGRAM
# string_exe = 'python code/aux.py ' + path + ' ' + 'data_'+MONTH+'_'+YEAR+'.xlsx'
# os.system(string_exe)

# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------

print('Complete!!!!!')

# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------

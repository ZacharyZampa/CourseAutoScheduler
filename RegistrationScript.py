import datetime
import openpyxl
import mechanize
import os

# Process the CRN codes from a supplied worksheet
os.chdir("C:\\Users\\zackz\\Documents\\School")  # change working directory to where spreadsheet is
wb = openpyxl.load_workbook('ClassPlan.xlsx')  # load in the workbook
sheet = wb['Original']  # choose the worksheet

# Access the cells holding the CRNs
crnTuple = tuple(sheet['E34':'E37'])
uid = sheet['B40'].value
password = sheet['B41'].value


crnList = []
# Get the CRNs that are to be entered and store into list for easy access
for rowOfCellObjects in crnTuple:
    for cellObj in rowOfCellObjects:
        crnList.append(cellObj.value)

# print codes acquired to user
print("The courses with the following CRN codes will be added")
for code in crnList:
    print("\t" + str(code))

wb.close()  # close the workbook


desiredTime = "20:00:00"  # Time to begin registration at

# Loop until time to begin
while True:
    now = datetime.datetime.now()  # get the current system time
    desTime = datetime.datetime.strptime(desiredTime, "%H:%M:%S")
    desTime = now.replace(hour=desTime.time().hour, minute=desTime.time().minute,
                          second=desTime.time().second, microsecond=0)
    if now > desTime:
        break


# Begin registration process now that it is the specified time


# Fill out the form for the user and submit
url = "https://bannerweb.miamioh.edu/ban8ssb/bwskfreg.P_AltPin"
br = mechanize.Browser()
br.set_handle_robots(False)  # ignore robots
br.set_handle_redirect(True)
br.set_handle_refresh(True)
br.open(url)  # open the url

# login the user
br.select_form(nr=0)

br.form['sid'] = uid
br.form['PIN'] = password
br.submit()

print("Signed in as: ", uid, "\n")

# navigate to the registration page
br.open('https://bannerweb.miamioh.edu/ban8ssb/bwskfreg.P_AltPin')

br.select_form(nr=1)

br.form['term_in'] = ['202010']  # enter in the term number here ---------------------
br.submit()

br.select_form(nr=1)

# # add info to the slots
# cid = 1
# for code in crnList:
#     cid_int = int(cid)  # convert ID tag to int
#     # ensure it is a valid id tag
#     if 1 <= cid_int <= 10:
#         print("Adding: ", crnList[(cid_int-1)])
#         to_add = br.form.find_control(name='CRN_IN', id='crn_id' + str(cid_int))
#         to_add.value = str(crnList[(cid_int-1)])
#     cid = cid + 1
#
#
# br.submit()  # submit the form
#
# print("submitted")


forms = list(br.forms())  # convert all forms to a list

course = ""  # empty string to print with
print("\n----Displaying added courses----")
print("\tCRN\t\tSUBJ\tCRSE\tSEC")
for key, value in forms[1]._pairs():
    if value == "DUMMY":
        # ignore the dummy variables
        continue
    if key == "CRN_IN":
        course += value
    if key == "SUBJ":
        course += "\t" + value
    if key == 'CRSE':
        course += "\t\t" + value
    if key == 'SEC':
        course += "\t\t" + value
        print("\t" + course)
        course = ""


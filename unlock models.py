import os
import win32com.client
import easygui

print("Program is initiating. \nAll existing Excel processes will be shut down.\nPlease take the opportunity to save "
      "and close existing Excel files.\nYou can stop the program by exiting the windows console.\n")
password = easygui.passwordbox("enter your password to unlock", title='Unlocking Models')
i = 0
t = '_unlocked'


def unlockExcel(filename1, filename2):
    xcl = win32com.client.Dispatch("Excel.Application")
    wb = xcl.Workbooks.Open(filename1, False, False, None, password, password, True)
    xcl.DisplayAlerts = False
    wb.SaveAs(filename2, None, '', '')
    xcl.Quit()


for entry in os.scandir(os.getcwd()):

    if (entry.path.endswith(".xlsx") or entry.path.endswith(".xlsb")) and entry.is_file():
        try:
            unlockExcel(entry.path, entry.path[:-5] + t + entry.path[-5:])
        except:
            easygui.msgbox("Password is incorrect or unknown error is encountered. Try again, before contacting Ray.",
                           title='Error Alert')
            os.sys.exit(1)
        print(entry.name + " is unlocked now!\n")
        i += 1


easygui.msgbox(
    "Job is completed with " + str(i) + " file(s) unlocked.\n" + "Feedbacks are welcomed to starstream521@gmail.com",
    title='Job Finished!')


os.sys.exit(0)

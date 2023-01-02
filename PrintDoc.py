"""
print all files in the designated folder and its subfolder
author: Ruru Dai
"""

import os
import win32api
import win32print

"""
print 
@path: String filepath
@file_extension: String extension for the file to print
@subfolder: boolean to indicate access subfolder or not 
"""
def get_files(path, file_extension):
    filelist = []
    for root, dirs, files in os.walk(path):
        for file in files:
            if file_extension.lower() == 'all':
                filelist.append(os.path.join(root,file))
            elif file.endswith("." + file_extension.lower()):
                filelist.append(os.path.join(root,file))
    return filelist

def main():
    # Ask for user input
    while True:
        filepath = input("Please provide your file path\n")
        file_extenstion = input("Please enter the file extension you want to print, eg. xls, pdf, doc\n"
                                "or enter 'all' if you want to print all files in the filepath\n")

        # Get the list of files user want to print
        fileList = get_files(filepath,file_extenstion)

        # Print out the name of files for user to confirm
        if len(fileList) == 0:
            print ("no file in the extension provided")
        else:
            print ("Here are the documents you want to print")
            for file in fileList:
                print(file)

        # ask for user input for the next step
        printfiles = input("\nPlease enter\n"
                           "'y' to start printing\n"
                           "'r' to restart the program\n"
                           "'q' to exit the program\n")
        printfiles = printfiles.lower()

        # print documents
        if printfiles == "y":
            for file in fileList:
                win32api.ShellExecute(0,"print",file,None,".",0)
            print("Print completed. Thanks for using the mini printing automation!")
            continue
        # restart the program
        elif printfiles == "r":
            print("Program will restart")
            continue

        # exit the program
        elif printfiles == "q":
            print("Exit. Thanks for using the mini printing automation!")
            return

        # restart the program for invalid input
        else:
            print("Please provide a valid input. Program will restart")
            continue
if __name__ == '__main__':
    main()
"""
print all files in the designated folder and its subfolder
author: HR_Coding
"""

import os

import pywintypes
import win32api

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
            elif file.endswith(file_extension.lower()):
                filelist.append(os.path.join(root,file))
    return filelist

def main():
    print("\n欢迎使用HR自动化工具箱 —— 一键式打印")
    while True:
        filepath = input("\n请提供需要打印的文件夹的真实路径，例如：C:\Documents\Resume\FoldertoPrint\n")
        file_extenstion = input("\n请输入你需要打印的文件格式,例如，“docx”， “xls”, “pdf”\n"
                                "或者输入“all”打印文件夹内所有文件\n")

        # Get the list of files user want to print
        fileList = get_files(filepath,file_extenstion)

        # Print out the name of files for user to confirm
        if len(fileList) == 0:
            print ("没有找到所输入的文件格式")
            continue
        else:
            print ("以下是所有符合条件的文件")
            for file in fileList:
                print(file)

        # ask for user input for the next step
        printfiles = input("\n请输入\n"
                           "'y' 开始打印\n"
                           "'r' 重启一键式打印\n"
                           "'q' 退出\n")
        printfiles = printfiles.lower()

        if printfiles == "y":
            try:
                number_print = int(input("请输入所需打印份数"))
                count = 0
                while count <= number_print:
                    for file in fileList:
                        win32api.ShellExecute(0, "print", file, None, ".", 0)
                        count += 1
                print("\n打印完成，感谢您使用HR自动化工具箱——一键式打印。\n关注微信公众号，获得更多工具，提高工作效率")
            except pywintypes.error:
                print("\n打印机未连接")
                continue
            except TypeError:
                print("请输入有效数字")
            except ValueError:
                print("请输入有效数字")
            except Exception as e:
                print("\n出错啦！请将错误信息截图发送给客服:\n" + e)
                continue

        # restart the program
        elif printfiles == "r":
            print("\n重启一键式打印")
            continue

        # exit the program
        elif printfiles == "q":
            print("\n退出，感谢您使用HR自动化工具箱——一键式打印。\n关注微信公众号，获得更多工具，提高工作效率")
            break

        # restart the program for invalid input
        else:
            print("\n请输入有效字母")
            continue


if __name__ == '__main__':
    main()

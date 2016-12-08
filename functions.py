import os, shutil, pprint
from openpyxl import load_workbook

fileNames = {}
alts = {}
scriptNames = {}
missingFiles = {}
lineTotal = 0


def getFileNames():
    '''
    Search within cwd, adding all .wav files to fileNames
    :return:
    '''
    for i in os.listdir('.'):
        if not str(i).endswith('.wav'):
            continue
        if not in fileNames:
            fileNames[i] = True
        else:
            alts[i] += 1


def checkMissing(script, file_names):
    '''
    compare dicts script and files, returns a with missing values
    :param script:
    dictionary
    :param file_names:
    dictionary
    :return:
    '''
    missing_files = {}
    for i in script:
        if i not in file_names:
            missing_files[i] = True
    return missing_files


def getWorkbook(directory):
    '''
    iterate a directory to find .xlsx files and adds them to a list.
    :return:list
    '''
    workbook_names = []
    for i in os.listdir(directory):
        if not str(i).endswith('.xlsx'):
            continue
        workbook_names.append(str(i))
    if not len(workbook_names):
        return False
    elif len(workbook_names) == 1:
        return str(workbook_names[0])
    else:
        return workbook_names


def choose_workbook(directory):
    '''
    looks in directory and adds found xlsx files
    to a list which is returned only if there is more than one entry.
    otherwise exit()
    :param directory:
    :return: list or exit()
    '''

    workbook_names = getWorkbook(directory)
    if not workbook_names:
        print('No .xlsx files found. Exiting...')
        exit()
    elif len(workbook_names) == 1:
        return str(workbook_names[0])
    else:
        # TODO - allow user to choose a file from this list and return it
        print("There are multiple .xlsx files in the directory:")
        pprint(workbook_names)
        exit()


def get_workbook(workbook_path):
    wb = load_workbook(filename=str(workbook_path))
    return wb


def choose_wb_sheet(wb):
    '''
    prints names of sheets in a workbook object along with their index in a list.
    user input gives index of chosen sheet
    returns string of the name of sheet chosen by user
    :param wb: workbook object
    :return: string of chosen sheet from wb
    '''
    ws_list = wb.get_sheet_names()
    for i in ws_list:
        print('{}. {}'.format(ws_list.index(i), str(i)))
    while True:
        try:
            sheet_number = int(input("Which # sheet do you want to open?"))
        except ValueError:
            pass
        if (sheet_number > -1) and (sheet_number <= len(ws_list) - 1):
            break
    return str(ws_list[sheet_number])


def find_ws_column(ws, keyword):
    '''
    iterate rows in ws, searching for value 'KEYWORD'
    :param ws: ws object
    :param keyword: string to search for in column
    :return: string of chosen sheet from s
    '''
    column_list = []
    for col in ws.columns[0]:
        column_list.append(col.value)
    if str(keyword) in column_list:


def get_scriptNames(wb):

import win32com.client as win32
import os


def os_walk(root):
    list_filepath = []
    list_files = os.walk(root)
    for path, dir_list, file_list in list_files:
        # print(file_list)
        for file_name in file_list:
            # print(os.path.join(path, file_name))
            list_filepath.append(os.path.join(path, file_name))
    return list_filepath


def xls_to_xlsx(filename_list):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    for filename in filename_list:
        wb = excel.Workbooks.Open(filename)
        wb.SaveAs(filename+"x", FileFormat=51)
        wb.Close()
    excel.Application.Quit()


def remove_xls(filename_list):
    for file_name in filename_list:
        if file_name.split(".")[-1] == 'xls':
            os.remove(file_name)


def extract_file(root):
    file_path_list_raw = os_walk(root)
    file_path_list_xlsx = []
    file_path_list_xls = []
    list_result = []
    for file_name in file_path_list_raw:
        if file_name.split(".")[-1] == 'xlsx':
            file_path_list_xlsx.append(file_name)
        elif file_name.split(".")[-1] == 'xls':
            file_path_list_xls.append(file_name)
        else:
            0
    file_path_list_xls_abs_path = ['C:\\Users\\closeet\\PycharmProjects\\AssetPenetration\\' + file_path for file_path
                                   in file_path_list_xls if file_path.split(".")[0] not in [
                                       files.split(".")[0] for files in file_path_list_xlsx
                                   ]]
    # print(file_path_list_xls_abs_path)
    xls_to_xlsx(file_path_list_xls_abs_path)
    remove_xls(file_path_list_xls_abs_path)
    for file_name in os_walk(root):
        if file_name.split(".")[-1] == 'xlsx':
            list_result.append(file_name)
    list_result = [result.replace('\\', '/') for result in list_result]
    return list_result




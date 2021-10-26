from RPA.Excel.Files import Files
from RPA.FileSystem import FileSystem


def main():
    lib = Files()
    lib_2 = FileSystem()
    dir_path = 'output'
    lib_2.create_directory(dir_path)
    lib.create_workbook(dir_path, 'xlsx')
    lib.rename_worksheet('Sheet', 'Renamed sheet')
    lib.save_workbook(f'{dir_path}/test.xlsx')
    return


if __name__ == '__main__':
    main()

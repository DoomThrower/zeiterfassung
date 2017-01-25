import openpyxl
import collections

path_to_workhours_folder = 'C:/Users/DooM/PycharmProjects/zeiterfassung-git/zeiterfassung/'

wb = openpyxl.load_workbook(filename='employee-list.xlsx')
ws = wb.get_sheet_by_name('Lista pracownikow')


def format_name(name):
    """ Konwertuje z 'Imie Nazwisko' na 'imie.nazwisko' """
    return name.replace(' ', '.').lower()


def format_path(name):
    """ Generuje sciezke do pliku z excelem pracownika """
    return path_to_workhours_folder + name + '/' + name + '-zeiterfassung.xlsx'


def process_employee(name):
    """ Przetwarza godziny pracownika """
    print name
    name = format_name(name)
    employee_wb = openpyxl.load_workbook(filename=format_path(name), data_only=True)
    projects_dict = {}
    # dla kazdego arkusza w pliku
    for employee_ws in employee_wb.worksheets:
        # dla rzedow z projektami/godzinami
        for row_idx in range(4, employee_ws.max_row + 1):
            project_name = employee_ws.cell(row=row_idx, column=1).value
            hours = employee_ws.cell(row=row_idx, column=20).value # nie wiem czy zawsze w 20 kolumnie bedzie suma godzin
            add_project_hours(projects_dict, project_name, hours)
    # sortujemy slownik
    workhours = collections.OrderedDict(sorted(projects_dict.items()))
    for item in workhours.items():
        print '\t', item[0] + ': ' + str(item[1])


def add_project_hours(dict, name, hours):
    if name not in dict:
        dict[name] = hours
    else:
        dict[name] = dict[name] + hours


def main():
    # zbieramy liste pracownikow
    for row_idx in range(2, ws.max_column - 1):
        name = ws.cell(row=row_idx, column=1).value
        process_employee(name)


if __name__ == "__main__":
    main()

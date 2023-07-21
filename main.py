from operations import MyExcelTable


def main():
    path = 'C:/Users/Admin/OneDrive/Заявки Евгений/2023 год/5 Май'

    my_table = MyExcelTable(name='Реестр по АКТам 2023.xlsx', path=path)

    my_table.save_to_doc()
    # workbook = openpyxl.load_workbook('Реестр по АКТам 2023.xlsx', read_only=False)
    # sheet = workbook.active

    # my_table = MyExcelTable(path)

    # names = my_table.get_names_in_dir()
    # print(names)
    # cell_value = sheet['A5'].value.date().strftime('%d.%m.%Y')

    # print(cell_value)
    # row_values = [cell.value for cell in sheet[5]]
    # row_values[2] = 'Сочи'
    # print(row_values)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    main()

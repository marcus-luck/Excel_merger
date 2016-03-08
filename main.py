from openpyxl import load_workbook
import os


def main():
    folder = 'input\\'
    out_folder = 'output\\'

    files = os.listdir(folder)

    # Open old file
    out_name = 'Purchase.xlsx'
    output_name = out_folder + out_name
    print(output_name)
    wb = load_workbook(output_name)


    # Read each file as a matrix and append the content to a list
    for file in files:
        if file[0] is '~':
            pass
        else:
            wb_in = load_workbook(folder + file)
            ws = wb.create_sheet()
            ws.title = file[0:30]
            ws_in = wb_in.active

            for row in ws_in.iter_rows():
                sb = []
                for cell in row:
                    if cell.value != None:
                        sb.append(str(cell.value))
                    else:
                        sb.append('')

                ws.append( sb )

    output_name = out_folder + out_name
    print()

    try:
        wb.save(output_name)
    except:
        print("Please close Excel file")



if __name__ == '__main__':
    main()


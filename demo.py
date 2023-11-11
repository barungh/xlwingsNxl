from pathlib import Path
import xlwings as xw

file_path = Path(__file__).parent / 'hello.txt'

def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    if sheet["A1"].value == "Hello xlwings!":
        sheet["A1"].value = "Bye xlwings!"
    else:
        sheet["A1"].value = "Hello xlwings!"

def ahoy():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    excel_value = sheet.range('A1').value
    with open(file_path, 'w') as f:
        f.write(f'Called from ahoy:{excel_value}')


@xw.func
def hello(name):
    return f"Hello {name}!"


if __name__ == "__main__":
    xw.Book("demo.xlsm").set_mock_caller()
    main()

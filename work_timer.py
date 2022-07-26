from calendar import month_name, different_locale, monthrange
import datetime

import csv
from tempfile import NamedTemporaryFile

import openpyxl

from fastapi import UploadFile


class ParsedCsv:
    date_to_times: dict
    date_to_comment: dict
    current_year: int
    current_month: int

    @property
    def current_month_name(self) -> str:
        with different_locale("de_DE"):
            return month_name[self.current_month]


async def convert(file: UploadFile):
    parsed_csv: ParsedCsv = await parse_csv(file)
    workbook = openpyxl.load_workbook("Arbeitszeitnachweis Vorlage.xlsx")
    fill_workbook(workbook, parsed_csv)

    result_file = f"Arbeitszeiten_Asib_Kamalsada_{parsed_csv.current_year}_{parsed_csv.current_month:02d}_{parsed_csv.current_month_name}.xlsx"

    with NamedTemporaryFile() as tmp:
        workbook.save(tmp.name)
        tmp.seek(0)
        binary_result = tmp.read()

    print(f"csv read from '{file.filename}'")
    workbook.close()
    return result_file, binary_result


async def parse_csv(file: UploadFile):
    parsed_csv = ParsedCsv()
    parsed_csv.date_to_comment = dict()
    parsed_csv.date_to_times = {}

    contents = (await file.read()).decode("utf-8").splitlines()

    spam_reader = csv.DictReader([x.strip() for x in contents[:-3]])
    parsed_csv.current_month = None
    parsed_csv.current_year = None
    for row in spam_reader:
        start: str = row["Von"]
        end: str = row["Bis"]
        if datetime.date.fromisoformat(start.split()[0]) == datetime.date.fromisoformat(end.split()[0]):
            current_date = datetime.date.fromisoformat(start.split()[0])
            if parsed_csv.current_month:
                if parsed_csv.current_month != current_date.month:
                    raise Exception("not the same months")
            else:
                parsed_csv.current_month = current_date.month
                parsed_csv.current_year = current_date.year
        else:
            raise Exception("working over the end of a day")
        start_time = datetime.datetime.fromisoformat(start)
        end_time = datetime.datetime.fromisoformat(end)
        if not parsed_csv.date_to_times.get(current_date.day, None):
            parsed_csv.date_to_times[current_date.day] = list()
        parsed_csv.date_to_times[current_date.day].append((start_time, end_time))
        if row.get("Kommentar", "") != "":
            parsed_csv.date_to_comment[current_date.day] = row["Kommentar"]
        del start_time, end_time, start, end
    del spam_reader
    return parsed_csv


def fill_workbook(workbook, parsed_csv):
    sheet = workbook.active

    sheet["D4"] = parsed_csv.current_month_name

    for month_day in range(1, monthrange(parsed_csv.current_year, parsed_csv.current_month)[1] + 1):
        sheet[f"B{6 + month_day}"] = f"{month_day}."

    for day, times in parsed_csv.date_to_times.items():
        times.sort(key=lambda x: x[0])
        start: datetime.datetime = times[0][0]
        end: datetime.datetime = times[-1][-1]

        sheet[f"C{6 + day}"] = start.time()  # .isoformat(timespec="seconds")
        sheet[f"C{6 + day}"].number_format = "h:mm"
        sheet[f"D{6 + day}"] = end.time()  # .isoformat(timespec="seconds")
        sheet[f"D{6 + day}"].number_format = "h:mm"

        if len(times) > 1:
            pause = datetime.timedelta()
            for i in range(len(times) - 1):
                pause += times[i + 1][0] - times[i][1]
            sheet[f"E{6 + day}"] = pause
            sheet[f"E{6 + day}"].number_format = "h:mm"

        if parsed_csv.date_to_comment.get(day):
            sheet[f"i{6 + day}"] = parsed_csv.date_to_comment[day]

#
# o = win32com.client.Dispatch("Excel.Application")
#
# o.Visible = False
#
# path_to_pdf = os.path.abspath(os.path.join(os.path.dirname(sys.argv[1]), "pdfs", f"{file_name_without_extension}.pdf"))
# if os.path.exists(path_to_pdf):
#     os.remove(path_to_pdf)
#
# print_area = 'A1:I37'
#
# try:
#     wb = o.Workbooks.Open(result_path)
#     ws = wb.Worksheets[0]
#     ws.PageSetup.PaperSize = 9  # A4 laut https://docs.microsoft.com/de-de/office/vba/api/excel.xlpapersize
#     ws.ExportAsFixedFormat(0, path_to_pdf)
#     print(f"printed as pdf to {path_to_pdf}")
# except pywintypes.com_error as err:
#     print(err)
# finally:
#     wb.Close(False)

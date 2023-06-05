import os
import dataclasses
from calendar import monthrange
import datetime

import csv
from tempfile import NamedTemporaryFile

import openpyxl

from fastapi import UploadFile


ENCODING = "utf-8"


@dataclasses.dataclass
class Timed:
    start: datetime
    end: datetime
    comment: str


@dataclasses.dataclass
class WorkTimes:
    date_to_times: dict[int, list[Timed]]
    date_to_vacation_duration: dict[int, datetime.timedelta]
    holidays: list[int]
    sick_days: list[int]
    current_year: int = None
    current_month: int = None

    @property
    def current_month_name(self) -> str:
        return [
            "Januar",
            "Februar",
            "März",
            "April",
            "Mai",
            "Juni",
            "Juli",
            "August",
            "September",
            "Oktober",
            "November",
            "Dezember"
        ][self.current_month - 1]


async def convert(file: UploadFile):
    work_times: WorkTimes = await parse_csv(file)
    workbook = openpyxl.load_workbook("Arbeitszeitnachweis Vorlage.xlsx")
    fill_workbook(workbook, work_times)

    result_file = f"Arbeitszeiten_Asib_Kamalsada_{work_times.current_year}" \
                  f"_{work_times.current_month:02d}_{work_times.current_month_name}.xlsx"

    with NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        workbook.save(tmp.name)

    print(f"csv read from '{file.filename}'")
    workbook.close()
    return result_file, tmp.name


async def parse_csv(file: UploadFile):
    work_times = WorkTimes(date_to_times={}, date_to_vacation_duration={}, holidays=[], sick_days=[])

    contents = [x.strip() for x in (await file.read()).decode(ENCODING).splitlines()[:-7]]

    uploaded_csv: csv.DictReader = csv.DictReader(contents)
    for row in uploaded_csv:
        activity: str = row["Aktivitätstyp"]
        start: datetime = datetime.datetime.fromisoformat(row["Von"])
        end: datetime = datetime.datetime.fromisoformat(row["Bis"])
        comment: str = row.get("Kommentar", None)

        if work_times.current_month is None:
            work_times.current_month = start.month
            work_times.current_year = start.year

        if start.date().replace(day=1) \
                != end.date().replace(day=1)  \
                != datetime.date(year=work_times.current_year, month=work_times.current_month, day=1):
            raise Exception("not the same months in the same years")

        match activity:
            case "Clocked":
                pass
            case "Homeoffice":
                if start.day != end.day:
                    raise Exception("working over the end of a day")

                if work_times.date_to_times.get(start.day, None) is None:
                    work_times.date_to_times[start.day] = []
                work_times.date_to_times[start.day].append(Timed(start, end, comment))
            case "Urlaub":
                pass
            case "Feiertag":
                pass
            case "Krank":
                pass
            case _:
                raise Exception(f"unknown activity type: {activity}")

    return work_times


def fill_workbook(workbook, work_times: WorkTimes):
    sheet = workbook.active

    sheet["D4"] = work_times.current_month_name

    for month_day in range(1, monthrange(work_times.current_year, work_times.current_month)[1] + 1):
        sheet[f"B{6 + month_day}"] = f"{month_day}."

    for day, times in work_times.date_to_times.items():

        times = [time for time in times if time.comment.lower() != "clocked"]

        if not times:
            continue

        comment = ", ".join({time.comment for time in times if time.comment.strip() != ""})
        if comment and comment.strip() != "":
            sheet[f"i{6 + day}"] = comment

        times.sort(key=lambda x: x.start)
        start: datetime.datetime = times[0].start
        end: datetime.datetime = times[-1].end

        sheet[f"C{6 + day}"] = start.time()  # .isoformat(timespec="seconds")
        sheet[f"C{6 + day}"].number_format = "h:mm"
        sheet[f"D{6 + day}"] = end.time()  # .isoformat(timespec="seconds")
        sheet[f"D{6 + day}"].number_format = "h:mm"

        if len(times) > 1:
            pause = datetime.timedelta()
            for i in range(len(times) - 1):
                pause += times[i + 1].start - times[i].end
            sheet[f"E{6 + day}"] = pause
            sheet[f"E{6 + day}"].number_format = "h:mm"


def cleanup(file_name):
    os.remove(file_name)


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

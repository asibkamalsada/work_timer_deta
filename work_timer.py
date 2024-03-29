import os
import dataclasses
from calendar import monthrange
import datetime

import csv
from tempfile import NamedTemporaryFile
from typing import Dict, List, Set

import openpyxl

from fastapi import UploadFile
from strenum import StrEnum

ENCODING = "utf-8"


class Activities(StrEnum):
    WORK = "Clocked"
    HOMEOFFICE = "Homeoffice"
    VACATION = "Urlaub"
    HOLIDAY = "Feiertag"
    SICK = "Krank"


@dataclasses.dataclass
class Timed:
    start: datetime
    end: datetime
    comment: str


@dataclasses.dataclass
class WorkTimes:
    work: Dict[int, List[Timed]]
    vacations: Dict[int, datetime.timedelta]
    holidays: Set[int]
    sick_days: Set[int]
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
            "Dezember",
        ][self.current_month - 1]


async def convert(file: UploadFile, current_month_year: str):
    month_year_selection = datetime.datetime.strptime(current_month_year, "%Y-%m")
    work_times: WorkTimes = await parse_csv(
        file, month_year_selection.year, month_year_selection.month
    )
    workbook = openpyxl.load_workbook("Arbeitszeitnachweis Vorlage.xlsx")
    fill_workbook(workbook, work_times)

    result_file = (
        f"Arbeitszeiten_Asib_Kamalsada_{work_times.current_year}"
        f"_{work_times.current_month:02d}_{work_times.current_month_name}.xlsx"
    )

    with NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        workbook.save(tmp.name)

    print(f"csv read from '{file.filename}'")
    workbook.close()
    return result_file, tmp.name


async def parse_csv(file: UploadFile, current_year: int, current_month: int):
    work_times = WorkTimes(
        work={},
        vacations={},
        holidays=set(),
        sick_days=set(),
        current_year=current_year,
        current_month=current_month,
    )

    contents = [x.strip() for x in (await file.read()).decode(ENCODING).splitlines()]

    uploaded_csv: csv.DictReader = csv.DictReader(contents)
    for row in uploaded_csv:
        try:
            activity: Activities = Activities(row["Aktivitätstyp"])
            start: datetime = datetime.datetime.fromisoformat(row["Von"])
            end: datetime = datetime.datetime.fromisoformat(row["Bis"])
            t = datetime.datetime.strptime(row["Dauer"], "%H:%M")
            duration: datetime.timedelta = datetime.timedelta(
                hours=t.hour, minutes=t.minute
            )
            comment: str = (row.get("Kommentar") or "").strip()
        except Exception:
            continue

        if start.day != end.day:
            if end.day - start.day > 1:
                raise Exception(f'worked at least one day straight (24h) in "{row}"')

            # assuming it is just one night between start and end

            if (
                start.year == work_times.current_year
                and start.month == work_times.current_month
            ):
                new_end = datetime.datetime.combine(start.date(), datetime.time.max)
                new_end = new_end.replace(second=0)
                await parse_activity(work_times, activity, start, new_end, comment)

            new_start = datetime.datetime.combine(end.date(), datetime.time.min)
            new_start = new_start.replace(second=0)
            if (
                new_start.year == work_times.current_year
                and new_start.month == work_times.current_month
            ):
                await parse_activity(work_times, activity, new_start, end, comment)
        else:
            await parse_activity(work_times, activity, start, end, comment)

    return work_times


async def parse_activity(
    work_times: WorkTimes,
    activity: Activities,
    start: datetime.datetime,
    end: datetime.datetime,
    comment: str,
):
    current_day = start.day

    if activity == Activities.WORK:
        pass
    elif activity == Activities.HOMEOFFICE:
        if current_day not in work_times.work:
            work_times.work[current_day] = []
        work_times.work[current_day].append(Timed(start, end, comment))
    elif activity == Activities.VACATION:
        work_times.vacations[current_day] = end - start
    elif activity == Activities.HOLIDAY:
        work_times.holidays.add(current_day)
    elif activity == Activities.SICK:
        work_times.sick_days.add(current_day)
    else:
        raise Exception(f"unknown activity type: {activity}")


def fill_workbook(workbook, work_times: WorkTimes):
    if not work_times.vacations.keys().isdisjoint(work_times.holidays):
        raise Exception(
            "took vacation on holidays: "
            f"{work_times.vacations.keys() & work_times.holidays}"
        )

    sheet = workbook.active

    sheet["D4"] = work_times.current_month_name

    for month_day in range(
        1, monthrange(work_times.current_year, work_times.current_month)[1] + 1
    ):
        sheet[f"B{6 + month_day}"] = f"{month_day}."

    # loop over homeoffice (or other commented not clocked work)
    for day, times in work_times.work.items():
        row = 6 + day

        if not times:
            continue

        comment = ", ".join({time.comment for time in times if time.comment != ""})
        if comment == "":
            sheet[f"I{row}"] = Activities.HOMEOFFICE
        else:
            sheet[f"I{row}"] = comment

        times.sort(key=lambda x: x.start)
        start: datetime.datetime = times[0].start
        end: datetime.datetime = times[-1].end

        sheet[f"C{row}"] = start.time()  # .isoformat(timespec="seconds")
        sheet[f"C{row}"].number_format = "h:mm"
        sheet[f"D{row}"] = end.time()  # .isoformat(timespec="seconds")
        sheet[f"D{row}"].number_format = "h:mm"

        if len(times) > 1:
            pause = datetime.timedelta()
            for i in range(len(times) - 1):
                pause += times[i + 1].start - times[i].end
            sheet[f"E{row}"] = pause
            sheet[f"E{row}"].number_format = "h:mm"

    for day, duration in work_times.vacations.items():
        row = 6 + day

        sheet[f"H{row}"] = Activities.VACATION
        sheet[f"G{row}"] = duration
        sheet[f"G{row}"].number_format = "h"

    for day in work_times.holidays:
        # don't write holidays into the Arbeitszeitnachweis
        pass

    for day in work_times.sick_days:
        row = 6 + day

        sheet[f"H{row}"] = Activities.SICK
        sheet[f"G{row}"] = 8


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

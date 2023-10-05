import os
import asynczipstream
import typing as t
from dataclasses import dataclass
from xml.sax.saxutils import escape
import csv
import io


TEMPLATE_DIR = f"{os.path.dirname(__file__)}/xlsx_template"

sheet_head = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr xr2 xr3" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" xr:uid="{8C2E347F-C8DB-3F42-B23A-5A89E855980D}"><dimension ref="A1:C2"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="D8" sqref="D8"/></sheetView></sheetViews><sheetFormatPr baseColWidth="10" defaultRowHeight="16" x14ac:dyDescent="0.2"/><sheetData>"""
sheet_foot = """ </sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>"""
start_row_template = """<row r="{row}" spans="1:3" x14ac:dyDescent="0.2">"""
end_row_template = """</row>"""
cell_template = """<c r="{column}{row}" t="inlineStr"><is><t>{value}</t></is></c>"""


@dataclass
class TemplateFile:
    filename: str
    content: bytes


def get_filepaths_in_dir(dirname):
    def _get_filepaths_recursive(dir_path, paths):
        for f in os.listdir(dir_path):
            fp = os.path.join(dir_path, f)
            if os.path.isfile(fp):
                paths.append(dir_path + "/" + os.path.basename(fp))
            else:
                _get_filepaths_recursive(
                    dir_path + "/" + str(os.path.basename(fp)), paths
                )

    paths = []
    _get_filepaths_recursive(paths=paths, dir_path=dirname)
    return paths


template_files: t.List[TemplateFile] = []
for file_path in get_filepaths_in_dir(TEMPLATE_DIR):
    with open(file_path, "rb") as template_file:
        template_files.append(
            TemplateFile(
                filename=file_path[len(TEMPLATE_DIR) :], content=template_file.read()
            )
    )


def column_number_to_name(n):
    d, m = divmod(n, 26)
    return column_number_to_name(d - 1) + chr(m + 65) if d else chr(m + 65)


class XlsxFile:
    def __init__(self):
        self.zip_stream = asynczipstream.ZipFile(
            mode="w", compression=asynczipstream.ZIP_STORED, allowZip64=True
        )
        self.is_sheet_header_in_writed = False
        for template_file in template_files:
            self.zip_stream.write_iter(
                arcname=template_file.filename,
                iterable=self._async_generator_wrapper(template_file.content),
                compress_type=asynczipstream.ZIP_STORED,
            )

    async def __aiter__(self):
        async for data in self.zip_stream:
            yield data

    async def __sheet_data_generator(self, data):
        yield sheet_head.encode("utf-8")

        row_number = 1
        async for row in data:
            yield start_row_template.format(row=row_number).encode("utf-8")
            column_number = 0
            async for cell in row:
                cell = escape(str(cell)) if cell is not None else ""
                yield cell_template.format(
                    column=column_number_to_name(column_number),
                    row=row_number,
                    value=cell,
                ).encode("utf-8")
                column_number += 1
            yield end_row_template.encode("utf-8")
            row_number += 1

        yield sheet_foot.encode("utf-8")

    async def _async_generator_wrapper(self, data):
        yield data

    def write_sheet(self, data):
        self.zip_stream.write_iter(
            arcname="xl/worksheets/sheet1.xml",
            iterable=self.__sheet_data_generator(data),
            compress_type=asynczipstream.ZIP_STORED,
        )


class CsvFile:
    def __init__(self, export_type: t.Literal["csv", "tsv"]):
        csv_dialect = {"csv": csv.excel, "tsv": csv.excel_tab}[export_type]
        self._csvfile_buffer = io.StringIO()
        self._csv_writer = csv.writer(self._csvfile_buffer, dialect=csv_dialect)

    async def __aiter__(self):
        async for row_generator in self._data:
            row = []
            async for value in row_generator:
                row.append(value)
            self._csvfile_buffer.seek(0)
            written_bytes_count = self._csv_writer.writerow(row)
            self._csvfile_buffer.seek(0)
            csv_row = self._csvfile_buffer.read(written_bytes_count).encode("utf-8")
            yield csv_row

    def write_sheet(self, rows_data):
        self._data = rows_data


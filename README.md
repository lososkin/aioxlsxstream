
# aioxlsxstream

Generate the xlsx file on fly without storing entire file in memory.
Generated excel workbook contain only one sheet. All data writes as strings.

## Installation

```
pip install aioxlsxstream
```

## Requirements

  * Python 3.8+

### Simple exapmle

```python
import aioxlsxstream
import asyncio

async def main():
    async def rows_generaor():
        async def row_cells_generator(row):
            for col in range(5):
                yield row*5 + col
        for row in range(10):
            yield row_cells_generator(row)

    xlsx_file = aioxlsxstream.XlsxFile()
    xlsx_file.write_sheet(rows_generaor())

    with open("example.xlsx", "wb") as f:
        async for data in xlsx_file:
            f.write(data)

asyncio.run(main())
```

### aiohttp Example

```python
from aiohttp import web, hdrs
import aioxlsxstream

async def rows_generaor():
    async def row_cells_generator(row):
        for col in range(5):
            yield row*5 + col
    for row in range(10):
        yield row_cells_generator(row)

async def handle(request):
    filename = "example.xlsx"
    response = web.StreamResponse(
        status=200,
        headers={
            hdrs.CONTENT_TYPE: "application/octet-stream",
            hdrs.CONTENT_DISPOSITION: (f"attachment; " f'filename="{filename}"; '),
        },
    )
    await response.prepare(request)

    xlsx_file = aioxlsxstream.XlsxFile()
    xlsx_file.write_sheet(rows_generaor())

    async for data in xlsx_file:
        await response.write(data)
    return response

app = web.Application()
app.add_routes([web.get('/', handle)])

if __name__ == '__main__':
    web.run_app(app)
```

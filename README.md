# Power Query snippets

Power Query (M) snippets I reuse when preparing data for Power BI and Fabric. Paste one into a blank query in the Advanced Editor and change the names in the first lines.

The paths, tables and column names are placeholders. Nothing here comes from a client or employer.

## Contents

1. [Date table](#1-date-table)
2. [Trim and clean every text column](#2-trim-and-clean-every-text-column)
3. [Unpivot month columns](#3-unpivot-month-columns)
4. [Combine Excel files from a folder](#4-combine-excel-files-from-a-folder)
5. [Keep only valid numbers](#5-keep-only-valid-numbers)

## 1. Date table

Generates one row per day. Change the start and end dates, add columns as needed, then mark the table as a date table in the model.

```m
let
    StartDate = #date(2020, 1, 1),
    EndDate   = #date(2030, 12, 31),
    Days      = List.Dates(StartDate, Duration.Days(EndDate - StartDate) + 1, #duration(1, 0, 0, 0)),
    AsTable   = Table.FromList(Days, Splitter.SplitByNothing(), {"Date"}, null, ExtraValues.Error),
    Typed     = Table.TransformColumnTypes(AsTable, {{"Date", type date}}),
    AddYear   = Table.AddColumn(Typed, "Year", each Date.Year([Date]), Int64.Type),
    AddMonthNo = Table.AddColumn(AddYear, "Month No", each Date.Month([Date]), Int64.Type),
    AddMonth  = Table.AddColumn(AddMonthNo, "Month", each Date.ToText([Date], "MMM"), type text)
in
    AddMonth
```

Tip: sort the `Month` column by `Month No` in the model so months do not sort alphabetically.

## 2. Trim and clean every text column

Removes leading and trailing spaces and non-printing characters from all text columns in one step, instead of one step per column.

```m
let
    Source   = Excel.CurrentWorkbook(){[Name = "RawData"]}[Content],
    TextCols = Table.ColumnsOfType(Source, {type text}),
    Cleaned  = Table.TransformColumns(
        Source,
        List.Transform(TextCols, each {_, each Text.Trim(Text.Clean(_)), type text})
    )
in
    Cleaned
```

`Table.ColumnsOfType` only finds columns that already have a text type, so set the types first.

## 3. Unpivot month columns

Turns a report with one column per month into one row per account and month, which is the shape a data model needs.

```m
let
    Source    = Excel.CurrentWorkbook(){[Name = "WideReport"]}[Content],
    Unpivoted = Table.UnpivotOtherColumns(Source, {"Account"}, "Month", "Amount"),
    Typed     = Table.TransformColumnTypes(Unpivoted, {{"Amount", type number}})
in
    Typed
```

Everything except the columns you list (here `Account`) is unpivoted, so new month columns are picked up automatically.

## 4. Combine Excel files from a folder

Reads the first sheet named `Sheet1` from every `.xlsx` file in a folder and stacks them. It skips temporary lock files that Excel creates while a workbook is open.

```m
let
    Source = Folder.Files("C:\Data\Monthly"),
    Files  = Table.SelectRows(
        Source,
        each Text.EndsWith([Name], ".xlsx") and not Text.StartsWith([Name], "~$")
    ),
    Parsed = Table.AddColumn(
        Files,
        "Data",
        each Excel.Workbook([Content], true){[Item = "Sheet1", Kind = "Sheet"]}[Data]
    ),
    Combined = Table.Combine(Parsed[Data])
in
    Combined
```

Every file must have the same column names. Put the folder path in a query parameter if the report will be refreshed from the service.

## 5. Keep only valid numbers

Filters out rows where an amount is text, an error or blank, so one bad cell does not break the refresh. Better still, send the removed rows to a separate query and review them.

```m
let
    Source = Excel.CurrentWorkbook(){[Name = "Entries"]}[Content],
    Valid  = Table.SelectRows(Source, each [Amount] is number),
    Errors = Table.SelectRows(Source, each not ([Amount] is number))
in
    Valid
```

Load `Errors` as its own table to get an exception list.

## Licence and use

Free to use and adapt. No warranty: run each snippet on a copy of your data first.

---

Nadeera Madushanka. Power BI, Microsoft Fabric and web apps. [nadeera.officesolutionscube.com](https://nadeera.officesolutionscube.com)

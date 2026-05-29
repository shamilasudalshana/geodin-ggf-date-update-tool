# GeoDIN GGF Date Tool - User Manual

## Purpose

This tool updates sampling dates inside a prepared GeoDIN `.GGF` graphics/layout file.

It is designed for workflows where the same GeoDIN layout is reused for another well, but the sampling dates need to be updated from Excel.

## Required files

You need:

1. A prepared GeoDIN `.GGF` template file
2. An Excel `.xlsx` file containing sampling dates and `Ionenbilanz`

## Important requirement for the GGF template

The `.GGF` template must contain 60 filled date slots.

Each `$SMPDATE$` value should contain an 8-digit date.

Example:

```text
$SMPDATE$='19000101'
```

Do not use empty date slots like:

```text
$SMPDATE$=''
```

Empty slots are shorter in the binary file and may break replacement logic.

## Excel requirements

The Excel file should contain:

- a date column, usually called `Datum`
- a filter column, usually called `Ionenbilanz`

The default filter is:

```text
-5 <= Ionenbilanz <= 5
```

Only dates where `Ionenbilanz` is within this range are written into the GGF file.

## Dummy date

Unused slots are filled with:

```text
19000101
```

This keeps the GeoDIN file structure stable.

## How to use

1. Open the tool.
2. Select the `.GGF` template file.
3. Select the Excel file.
4. Confirm the date column.
5. Confirm the filter column.
6. Confirm or change the filter range.
7. Choose the output `.GGF` file.
8. Click **Update GGF Dates**.
9. Check the generated log/report.

## Duplicate dates

If duplicate dates are found, the tool shows a warning.

You can choose whether to keep duplicate dates or remove duplicates before writing.

## Output

The tool creates:

1. Updated `.GGF` file
2. Text report/log file

## Recommended safety workflow

Always keep the original `.GGF` template unchanged.

Use a new output filename, for example:

```text
Br_3_temp_updated.GGF
```

Then open the updated file in GeoDIN and verify the result.

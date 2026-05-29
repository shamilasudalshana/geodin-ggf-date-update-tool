# Developer Notes

## Core principle

The `.GGF` file is treated as a binary file.

The tool only replaces existing fixed-length UTF-16-LE date blocks.

It does not insert or delete bytes.

## Why dummy dates are used

Filled slots such as:

```text
$SMPDATE$='20241113'
```

and empty slots such as:

```text
$SMPDATE$=''
```

have different byte lengths.

Changing the binary length can make GeoDIN unable to read the file.

Therefore, the tool keeps all 60 date slots filled with 8-digit dates.

Unused positions are filled with:

```text
19000101
```

## Regex

The backend searches for UTF-16-LE encoded blocks matching:

```text
$SMPDATE$='########'
```

where `########` is an 8-digit date.

## Future improvements

- better drag-and-drop support
- column matching with fuzzy search
- Excel preview table
- PDF report
- PyInstaller build script
- digital signing for internal company distribution

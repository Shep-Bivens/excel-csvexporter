## CHANGELOG: CSV Exporter Excel VBA Add-In

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](http://keepachangelog.com/en/1.0.0/)
and this project adheres to [Semantic Versioning](http://semver.org/spec/v2.0.0.html).


### [1.3.0] - 2026-02-01

#### Added
- **Option to use each output cell’s number format**  
  Each cell is output using the same number format that is used to
  display the value in the Excel worksheet.

- **Option to quote only values that require quoting**  
  A cell value is quoted only if it contains a quote, a
  separator, a carriage return, or a line feed.  
  This behavior matches Excel’s CSV output.

- **Option to use tab or blank as the separator**

- **Choice of output encoding**  
  ANSI, UTF-8, or UTF-8 with BOM.

- **File overwrite warning**  
  If append-to-file is not selected and an export would overwrite an
  existing file, a warning is displayed allowing the user to overwrite
  the file or cancel the export.

- **Warning for non-ANSI output with ANSI encoding selected**  
  If an export requires a non-ANSI character while ANSI encoding is
  selected, a warning message is displayed and no file is written or
  appended.

- **Settings persistence**  
  Current settings (folder, filename, format selection, etc.) are saved
  each time **Export** or **Close** is clicked and are restored when the
  exporter is loaded.  
  When closing with **Ctrl-Close**, the user is given the option to
  remove saved settings and close, save settings and close, or return to the
  exporter.

- **Context-sensitive help (up to two lines)**  
  When a Shift key is held down and the cursor is moved over an item in
  the exporter UI, help text for that item is displayed at the bottom of
  the exporter.

#### Changed
- **Separator warnings**  
  A warning is issued only if a separator appears in an output value
  *and* the value is not quoted.

#### Fixed
- **Extra output columns**  
  Eliminated extra output columns when exporting a full-sheet selection
  or a full-row selection.


### [1.2.2] - 2022-03-10

#### Fixed

- Fixed #36, truncation of cell contents to 255 characters by `Format`,
  by skipping the call to `Format(...)` when `Len(cel.Value)` is greater
  than 250. (Content lengths this long ***should*** only occur when the
  cell contains free text for which `Format()` is a no-op anyways.)


### [1.2.1] - 2021-07-28

#### Internal

- Renamed the `Sub` in `Exporter.bas` that launches the CSV Exporter
  `UserForm` to the less-anonymous `showCSVExporterForm()`.


### [1.2.0] - 2020-05-29

#### Added

- Implemented option for exporting data values surrounded by user-definable quotes.
  User can select whether to quote all exported values, or just
  non-numeric values (as determined by VBA's `IsNumeric` built-in).


### [1.2.0.dev2] - 2020-02-07

#### Fixed

- Fixed RTE 91 raised when the header option is enabled and 
  entire row(s)/column(s) are selected that do not intersect `.UsedRange`
  (#25).
- Due to the way Excel handles `.UsedRange` on an empty sheet (returns
  `Range("$A$1")` instead of `Nothing`), it was necessary to add explicit
  check for this case, in order to correctly report invalid selection
  status when entire row(s)/column(s) are selected on an empty sheet.


### [1.2.0.dev1] - 2020-02-03

Development release issued, to facilitate user testing before considering
the below new features as "final".

#### Added

- Hidden rows and columns now are **NOT** exported by default; checkboxes
  to enable export of hidden cells (per row and/or per-column) are
  now provided
- An option is now provided to export the cells from one or more rows on
  the active sheet above/below the exported data block as "header row(s)"


### [1.1.0] - 2019-01-08

#### Added

- New information box on the form indicates the sheet and range of
  cells currently set to be exported
- New warning added, if the separator appears in the data to be exported;
  this should minimize accidental generation of files that cannot be
  used subsequently, due to the excess separator characters

#### Changed

- UserForm now reappears in its prior location when closed
  and re-opened, instead of always reappearing in the center
  of the Excel window
- Selection of multiple areas now results in an "<invalid selection>"
  message in the new information box; and, greying out of the 'Export'
  button instead of a warning message after clicking 'Export'
- Selection of entire rows/columns now sets for export the intersection
  of the selection and the UsedRange of the worksheet; selection of an
  entire row/column outside the UsedRange of the worksheet gives an
  "<invalid selection>" message in the new information box and disables
  the 'Export' button

#### Fixed

- Userform now disappears when a chart-sheet is selected, and reappears
  when a worksheet is re-selected; Userform will silently refuse to open
  if triggered when a chart-sheet is active
- Error handling added around folder selection and output file opening
  for write/append


### [1.0.0] - 2016-01-30

*Initial release*

#### Features
- Folder selection works
- Name, number format, and separator entry work
- Append vs overwrite works
- Modeless form retains folder/filename/format/separator/etc. within a given Excel instance

#### Limitations
- Exports only a single contiguous range at a time

#### Internals
- Modest validity checking implemented for filename
  - Red text and disabled `Export` button on invalid filename
- No validity checking implemented for number format
- Disabled `Export` button if number format or separator are empty

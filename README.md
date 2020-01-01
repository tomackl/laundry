# Laundry: Folding spreadsheets into neat shapes

Spreadsheets are easy to record and manipulate data with but do a poor job of displaying the data in an easy to read way. Laundry provides a simple way of converting spreadsheet data into Word files by by specifying the output document's format within the spreadsheet itself.

Laundry provides two modes in which to generate the Word files:

* Single file mode, where details of a single output file are defined and passed at the CLI, and
* Multi file mode. where multiple output files can be defined and batched from a worksheet within the original Excel spreadsheet.

While spreadsheets are, in many aspects, similar to database tables, Laundry is only a formatting tool and not a database.

### Overview

A general overview of how Laundry operates is:

1. Having a spreadsheet containing some organised data stored in the `data_worksheet`. There is no restriction on how the data stored is the `data_worksheet` other than it must be a flat table.

2. Defining the output structure within a separate worksheet contained within the spreadsheet. This is known as the `structure_worksheet`.

3. The `structure_worksheet` references data to be formatted by using the the data's column names, the styles contained within the `template_file`, and how the data will be arranged, for example, as a table, paragraph or photo. The `structure_worksheet` does have a specific format that must be used to ensure that the conversion process takes place. 

4. A `template_file` Word document (`.docx` file) is used as the basis of the output file. This file will contain the styles that will be used to format the output file.

5. If a batch process is desired a `batch_worksheet` will be required to define the above requirements.

6. The app is run using the command line interface (CLI). This is as easy as:

    - `laundry single -t <template_file> <input_file> <output_file>`, or
    - `laundry multi <input_file>`

## Running the App

Laundry is operated from the command line, so if you're not familiar with command line then you'll need to familarise yourself with it first.

There are two ways of running Laundry, single mode and and multi mode. 

Single mode will convert a single `data_worksheet` into a `output_file` based on the `template_file` using the structure defined in the `structure_worksheet`. You may need to define the `data-header` if the data is not in the first row of the `data_worksheet`.

Multi mode allows one or more `output_files` based on information defined in the `format_worksheet`. This mode still requires the same information as single mode but it is stored within a worksheet in the `input_file`.

## Manual

### Input File

The `input_file` is the spreadsheet that contains the data to be formatted and exported to the `output_file`.

The `input_file` __must__ contain worksheets that contain:
 
* The structure or output formatting requirements of the output file: this is the `structure_worksheet`.
* When the `multi` sub-command a `format_worksheet` is required to allow the batch operation to take place.

The `input_file`'s data is stored in the `data_worksheet`. This worksheet can be given any name provided it doesn't clash with the `structure_worksheet` name.

### Structure Worksheet

A `structure worksheet` must be defined for each output file that will be produced. In `single` mode the worksheet is assumed to be associated with the `data worksheet` that is passed to the app. In `multi` mode a single `structure worksheet` can be used for multiple output files, however it must be explicitly referenced for each output file. 

The following column names __must__ implemented in the `structure_worksheet`. If they are not used the app will not work.

#### `section_type`
This defines the type of section that will be inserted with the output document. The top-to-bottom order will determine the order of the types in the `output_file`. The options that are currently supported are:

- `para`: This will insert a paragraph with the column header being the paragraph title styled using the style defined in `section_style`.

- `table`: Will insert a table that contains the data contained in `section_contains` cell with the table column containing the same column heading as the original data (with any underscores removed).

- `heading`: Will use the data contained within the cell as a section heading made up of a title (column heading) and a paragraph (cell contents).

- `photo`: Inserts one or more photos (images) into the paragraph. Multiple images can be added by including them in the `section_contains` column. The `photo` column heading is used to provide a file path to the directory containing the files. Including the photo's file extension is not required. The app will sequence through a number of popular formats before providing an error message to the standard output and adding the error message to the output document.

#### `section_contains`
This contains one or more of the column names used in the `data` worksheet. Multiple column names can be used with `table` using  `section_type` provided the are separated by `new lines` (alt + enter) in the cell, or commas.  `new lines` are the preferred method.

#### `section_style`
This Word *style* contained within the `template_file`'s `.docx` file. If this style is not in the `template_file` then Word's default styles will be used. *This is a limitation of Word*.

#### `title_style`
The Word formatting style to be used for titles within the document. This column does not apply to tables.

#### `section_break`
Must be either `True` or `False`. If `True` it will insert an empty paragraph provide a visual break between it and the following paragraph.

#### `page_break`
Must be either `True` or `False`. If `True` it will insert a page break (that is, start a new page) after the paragraph. This is useful for clearly separating information that is related to different `data_worksheet` rows.

#### `path`
This defines the relative file path to the directory (folder) containing the photos. All photos referenced in the cell must be stored in the same directory.

### `section_type` Limits

Each of the `section_type` types have some limit regarding their operation.

- Where `heading` is used it should be a single column heading per paragraph.

- Where `paragraph` is used it should be a single column heading per paragraph.

- `table` can use any number of columns headings. Each column heading should be separated by a `newline` (`\n`) (preferred option) or by a comma (`,`).

- `section_style` _can_ only be a single value. This is a string (text) that can contain spaces. Make sure that the spelling and capitalisation is correct.

- `title_style` _can_ only be a single value. This is a string (text) that can contain spaces. Make sure that the spelling and capitalisation is correct.

- `page_break` is a `True`/`False` value.

## Arranging Spreadsheet Data

The following applies to the both `single` and `multi` sub-commands.

### `data_worksheet`

The `data_worksheet` contains the data that will be formatted into the `output_file`.

1. The column headers should be in row 0 (that is, the first row) in the worksheet. If not, 
    
    * When using `single` sub-command the `--data-head` option must to be used, *and* the row number specified.  For example if the column headers start on row 5:

    `laundry single -df=5 -t <template-file> <input-file> <output-file>`
    
    *  When using `multi` mode the row must be recorded in the `header row` column.
    
    `laundry multi -f <format-worksheet> <input-file>`

2. Do not use numbers for column header names. This will cause problems.

3. Column header names *should* avoid spaces, and either use underscores (`_`) or use camel case (`ThisIsAnExampleOfCamelCase`). Stick numbers at the end of the names if you need numbers in the name.

4. If you want multiple paragraphs or bullet points, or more than one image, then use alt-enter in an Excel cell to allow this to take place.

5. Include the file extension (for example `.jpg`, `.png`) when entering including images in the cell. File extensions are part of the file name and the app expects them to be included.

### `structure_worksheet`

1. The column headers must be in row 0 (that is, the first row) in the worksheet. If not, the app will not work.

2. The `section_type` must be one of the following:

    - `photo`
    - `para`
    - `table`
    - `heading`

3. `table` sections do not require a value for `title_style`.

4. `photo` sections do not require a value for `title_style` or `section_style`, however they do require `path` to be completed.

5. `section_break` and `page_break` must be `TRUE` or `FALSE`.

### `batch_worksheet`

The `batch_worksheet` defines the requirements for each of the output files. The spreadsheet will require a row for each file that is to be output.

#### `data_worksheet` requirements

This is the name of the `data_worksheet` that contains the data to be exported as described above. 

#### `structure_worksheet` requirements

The`structure_worksheet`defines how the output file will be structured. This is the same as described above.

#### `header_row` requirements

The number of the row that defines the header row in the `data_worksheet`. This functionally identical to the `-dh` flag used with the `single` sub-command as described above.

#### `drop_empty_columns` requirements

Not currently used.

#### `template_file` requirements

The name of the `.docx`file being used as the template for the output file. Different template files can be used provided that the are all stored in the directory that the `laundry` command is run from.

#### `filter_rows`

This allows rows to be filtered from the output file by only including the rows that meet the criteria defined in the filter. More than one row can be specified to filter by.

The row requires a particular format `<row_heading>:<criteria1,...,criteriaN>`. The heading row must be separated from the filter criteria by a colon (`:`). The filters criteria *must* be separated by commas (`,`), and each row must be separated by a newline (`alt-enter` in Excel).

The filter criteria must be _identical_ to the cell values. If the criteria can take different forms, that either:
- All forms must be included in the filter, or
- Correct all the values so that they take common form (probably the better option). 

#### `output_file`

The name of the file that is produced by the app. The file will be saved in the directory that the app is run from. The file name can be a relative or absolute file path.

## FAQs

The following is a list of commonly experienced issues.

### I get something containing `KeyError` and what looks like a column header

The likely culprit is an incorrectly spelt column header in the `structure_worksheet` or you haven't allowed for the conversion of underscores (`_`) replacing spaces within the column headers.


### Something like `No sheet named <[some_worksheet_name]>`

Check the `data_worksheet` and `structure_worksheet` names that you've used when running the app from the CLI. Check for spelling mistakes.

### `Error: Invalid value for "input_file"` appears

Check the spelling of the `input_file`'s name and that the file path is correct. The app will check that the `input_file` exists before attempting to run the app.

### `AttributeError: 'int' object has no attribute 'lower'`

The likely issue here is the column headers are not contained in the first row of the `data_worksheet`. Also check that none of the column headers are numbers.

### `Error: Invalid value for "--template" / "-t": Path "[some_path_&_filename].docx" does not exist.`

The `template_file` filename is incorrect. This could also be the path to the file. Check both.

### I have a heap of formatted empty pages at the bottom of my `converted_file`.

For each row of your `data_worksheet` that contains some data the app will produce a formatted section. By removing all the rows that you don't want at the bottom of the `data_worksheet` you can prevent this from occurring. There is a way to drop rows that are missing data, however this requires some work to enable this feature.

### `KeyError: "no style with name '[some_Word_style]'"`

The Word formatting style is not present in the `template_file`. Check that the style name matches exactly its name in Word.

Some Word styles appear to be concatenations of other styles within Word, e.g. 'Heading 3, List'. These styles don't work as expected and the reason for this has not been determined.

### How do I use a specific Word style?

Due to limitations with Word user specific styles need to be saved to  the `template_file` for them to be available. If the specified name is not present in the `template_file` then the app will not function.

### I want to create a single table from the data but every row has the header details: how do I remove the header details without using Word?

This is a known problem, currently without a solution. A solution is being considered.
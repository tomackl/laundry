# Templar: Beating spreadsheets into shape

Spreadsheets are easy to record and manipulate data with but do a poor job of displaying the data in an easy to read way. Templar provides a simple way of converting spread data into word by by specifying the output document's format within the spreadsheet itself.

While spreadsheets are, in manner aspects, similar to database tables, Templar is only a formatting tool and not a database.

### Overview

Templar operates by:

1. Having a spreadsheet containing some organised data stored in the `data_worksheet`.

2. Defining the output structure within a separate worksheet contained within the spreadsheet. This is known as the `structure_worksheet`.

3. The `structure_worksheet` references data to be formatted by using the the data's column names, the styles contained within the `template_file`, and how the data will be arranged, for example, as a table, paragraph or photo. The `structure_worksheet` does have a specific format that must be used to ensure that the conversion process takes place.

4. A `template_file` Word document (`.docx` file) is used as the basis of the output file. This file will contain the styles that will be used to format the output file.

5. The app is run using the command line interface (cli).


## Manual

### Input File

The `input_file` is the spreadsheet that contains the data to be formatted and exported to the `output_file`.

The `input_file` __must__ contain a worksheet that contains the structure or output formatting requirements of the output file: this is the `structure_worksheet`.

The `input_file`'s data is stored in the `data_worksheet`. This worksheet can be given any name provided it doesn't clash with the `structure_worksheet` name.

### Structure Worksheet

The following column names __must__ implemented in the `structure_worksheet`. If they are not used the app will not work.

#### `sectiontype`
This defines the type of section that will be inserted with the output document. The top-to-bottom order will determine the order of the types in the `output_file`. The options that are currently supported are:

- `para`: This will insert a paragraph with the column header being the paragraph title styled using the style defined in `sectionstyle`.

- `table`: Will insert a table that contains the data contained in `sectioncontains` cell with the table column containing the same column heading as the original data (with any underscores removed).

- `heading`: Will use the data contained within the cell as a section heading made up of a title (column heading) and a paragraph (cell contents).

- `photo`: Inserts one or more photos (images) into the paragraph. Multiple images can be added by including them in the `sectioncontains` column. The `photo` column heading is used to provide a file path to the directory containing the files.

#### `sectioncontains`
This contains one or more of the column names used in the `data` worksheet. Multiple column names can be used with `table` using  `sectiontypes` provided the are separated by `new lines` (alt + enter) in the cell, or commas.  `new lines` are the preferred method.

#### `sectionstyle`
This Word *style* contained within the `template_file`'s `.docx` file. If this style is not in the `template_file` then Word's default styles will be used. *This is a limitation of Word*.

#### `titlestyle`
The Word formatting style to be used for titles within the document. This column does not apply to tables.

#### `sectionbreak`
Must be either `True` or `False`. If `True` it will insert an empty paragraph provide a visual break between it and the following paragraph.

#### `pagebreak`
Must be either `True` or `False`. If `True` it will insert a page break (that is, start a new page) after the paragraph. This is useful for clearly separating information that is related to different `data_worksheet` rows.

##### `path`
This defines the relative file path to the directory (folder) containing the photos. All photos referenced in the cell must be stored in the same directory.

### `sectiontype` Limits

Each of the `sectiontypes` have some limit regarding their operation.

- Where `heading` is used it should be a single column heading per paragraph.

- Where `paragraph` is used it should be a single column heading per paragraph.

- `table` can use any number of columns headings. Each column heading should be separated by a `newline` (preferred option) or by a comma (`,`).

- `sectionstyles` _can_ only be a single value.

- `titlestyles` _can_ only be a single value.

- `pagebreak` is a `True`/`False` value.

## Data Setout

[add something about the need to set the worksheets with the column headers in the first row of the worksheet otherwise bad things happen.]

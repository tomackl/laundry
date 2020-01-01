2020.1.1
========

Features
--------

* The readme has been updated to include changes made in 2020.1.0.
* The changelog has been updated. :)

Bug Fixes
---------

Other
-----

2020.1.0
========

Features
--------

The app has been completely rewritten in an effort to make it faster and more verbose to communicate what is going on the user. For those interested dictionaries have been replaced by greater uses of Pandas DataFrames, and avoided repeated imports from disk.

Bug Fixes
---------

* It's been completely rewritten. There will be new bugs. ;)

Other
-----

2019.0.7
========

Features
--------

* Added a check to ensure that any referenced directories are checked to exist before proceeding.

Bug Fixes
---------

Other
-----

Deprecations
------------

2019.0.5
========

Features
--------

* The ability to batch multiple output files at the same time has been.
* The app has two sub-commands "single" and "multi". "single" allows a single output file to be produced using the standard flags, "multi" requires reference to a batch worksheet within the spreadsheet.
* A number of error checks that allow for missing worksheets to be detected and addressed.

Bug Fixes
---------

Other
-----

* Dedicated function to control exporting of the document,

Deprecations
------------

* None


2019.0.4
========

Features
--------

* Added '-sw' to structure worksheet CLI flag
* CLI help updated to be explicit regarding the 'Master List' as default name for the data worksheet name.
* Failure to include a file extension to a photo will not raise an exception. The app will cycle through '.jpg', '.jpeg', '.png', and '.tiff' file extension to check for the existence of the the file before outputing an error message and inserting a 'PHOTO NOT FOUND' message into the output file.

Bug Fixes
---------

* The path component has been modified to allow either Windows or POSIX paths to be used. This allows a standard template to be used without concern regarding the OS it is run on.
* Rows that do not have at least two columns populated in in a row will be dropped.

Other
-----
* None

Deprecations
------------

* None


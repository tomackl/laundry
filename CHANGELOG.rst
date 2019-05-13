2019.0.4
========

NOT RELEASED

Features
--------

* Added '-sw' to structure worksheet CLI flag
* CLI help updated to be explicit regarding the 'Master List' as default name for the data worksheet name.
* Failure to include a file extension to a photo will not raise an exception. The app will cycle through '.jpg', '.jpeg', '.png', and '.tiff' file extension to check for the existence of the the file before outputing an error message and inserting a 'PHOTO NOT FOUND' message into the output file.

Bug Fixes
---------

* Empty data worksheet columns will be removed if one or more columns is empty in the worksheet.
* The path component has been modified to allow either Windows or POSIX paths to be used. This allows a standard template to be used without concern regarding the OS it is run on.
* Rows that do not have at least two columns populated in in a row will be dropped.

Other
-----
* None

Deprecations
------------

* None


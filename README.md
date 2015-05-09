# xlrd_diff
This is a simple script for "diffing" Excel files.
It is based on xlrd, and should support .xls and .xlsx files.

Usage:

``exceldiff [--names] file1.xls file2.xls``

The ``--names`` option tells the script to match sheets based on their names,
while the default is to match them based on their positions (first sheet of
``file1.xls`` with first sheet of ``file2.xls``, and so on).

Notice the [ExcelCompare](https://github.com/na-ka-na/ExcelCompare)
project has many more options, a better ouput, and support for more types of files.

(The motivation for xlrd_diff was just that ExcelCompare failed on some files)

Also see [this StackOverflow question](http://stackoverflow.com/questions/114698/how-do-i-create-a-proper-diff-of-two-spreadsheets-when-i-perform-a-version-contr).

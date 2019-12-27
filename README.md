# utl-import-ms-word-docx-table-to-sas-table
Import ms word docx table to sas table
    Import ms word docx table to sas table

    Python is the best language for this?

    github
    https://tinyurl.com/wkaubvg
    https://github.com/rogerjdeangelis/utl-import-ms-word-docx-table-to-sas-table

    Not tested in WPS. No time to hook up my autocall library without the WPS CLI.

    SAS Forum
    https://tinyurl.com/v5c7e98
    https://communities.sas.com/t5/SAS-Programming/To-read-width-of-colums-of-a-table/m-p/613869

    pyreadstat is experimental?

    There are SAS macros to import rtf tables, but I don't believe there is macro for
    DOCX files.

    *_                   _
    (_)_ __  _ __  _   _| |_
    | | '_ \| '_ \| | | | __|
    | | | | | |_) | |_| | |_
    |_|_| |_| .__/ \__,_|\__|
            |_|
    ;

    /*
    You need to open the ods output rtf table in word and
    then save as a d;/doc/class.docx.;
    If you don't have word then run the code on the end of this message.
    The ouput is a MS word docx file.
    */

    ods rtf file="d:/rtf/class.rtf" style=minimal;
    proc print data=sashelp.classfit(keep=name age sex obs=5);
    run;quit;
    ods rtf close;

      d;/doc/class.docx

       +--------+------------+------------+------------+
       | OBS    | NAME       |   SEX      |    AGE     |
       +--------+------------+------------+------------+
       |  1     |Alice       |    F       |    15      |
       +--------+------------+------------+------------+
       |  2     |Al          |    M       |    12      |
       +--------+------------+------------+------------+
       |  3     |Mary        |    F       |    16      |
       +--------+------------+------------+------------+
       |  4     |Joan        |    F       |    13      |
       +--------+------------+------------+------------+
       |  5     |Jake        |    M       |    14      |
       +--------+------------+------------+------------+

    *            _               _
      ___  _   _| |_ _ __  _   _| |_
     / _ \| | | | __| '_ \| | | | __|
    | (_) | |_| | |_| |_) | |_| | |_
     \___/ \__,_|\__| .__/ \__,_|\__|
                    |_|
    ;

    WORK.WANT total obs=5

    Obs    AGE     NAME      SEX

     1     11     Joyce       F
     2     12     Louise      F
     3     13     Alice       F
     4     12     James       M
     5     11     Thomas      M

    *          _       _   _
     ___  ___ | |_   _| |_(_) ___  _ __
    / __|/ _ \| | | | | __| |/ _ \| '_ \
    \__ \ (_) | | |_| | |_| | (_) | | | |
    |___/\___/|_|\__,_|\__|_|\___/|_| |_|

    ;

    %utlfkil(d:/xpt/want.xpt);

    %utl_submit_py64_37("
    import pandas as pd;
    from docx.api import Document;
    import pyreadstat;
    document = Document('d:/doc/class.docx');
    table = document.tables[0];
    data = [];
    keys = None;
    for i, row in enumerate(table.rows):;
    .   text = (cell.text for cell in row.cells);
    .   if i == 0:;
    .       keys = tuple(text);
    .       continue;
    .   row_data = dict(zip(keys, text));
    .   data.append(row_data);
    .   print (data);
    want = pd.DataFrame(data);
    print(want);
    pyreadstat.write_xport(want, 'd:/xpt/want.xpt',table_name='want');
    ");

    libname xpt xport "d:/xpt/want.xpt";
    data want;
      set xpt.want;
    run;quit;
    libname xpt clear;

    *_
    | | ___   __ _
    | |/ _ \ / _` |
    | | (_) | (_| |
    |_|\___/ \__, |
             |___/
    ;

    [{'Obs': '1', 'NAME': 'Joyce', 'SEX': 'F', 'AGE': '11'}]
    [{'Obs': '1', 'NAME': 'Joyce', 'SEX': 'F', 'AGE': '11'},
    [{'Obs': '1', 'NAME': 'Joyce', 'SEX': 'F', 'AGE': '11'},
    [{'Obs': '1', 'NAME': 'Joyce', 'SEX': 'F', 'AGE': '11'},
    'Obs': '4', 'NAME': 'James', 'SEX': 'M', 'AGE': '12'}]
    [{'Obs': '1', 'NAME': 'Joyce', 'SEX': 'F', 'AGE': '11'},
    'Obs': '4', 'NAME': 'James', 'SEX': 'M', 'AGE': '12'}, {
      AGE    NAME Obs SEX
    0  11   Joyce   1   F
    1  12  Louise   2   F
    2  13   Alice   3   F
    3  12   James   4   M
    4  11  Thomas   5   M

    4045  libname xpt xport "d:/xpt/want.xpt";
    NOTE: Libref XPT was successfully assigned as follows:
          Engine:        XPORT
          Physical Name: d:\xpt\want.xpt
    4046  data want;
    4047    set xpt.want;
    4048  run;

    NOTE: There were 5 observations read from the data set XPT.WANT.
    NOTE: The data set WORK.WANT has 5 observations and 4 variables.
    NOTE: DATA statement used (Total process time):
          real time           0.00 seconds
          user cpu time       0.00 seconds

    4048!     quit;
    4049  libname xpt clear;
    NOTE: Libref XPT has been deassigned.


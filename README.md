SpecimenInfo
============
Fetch and format plant specimen informations from data file and web, save
outcome to xlsx file and SQLite3 db file.

Screen Shot (GUI Version)
-------------------------
![Screen Shot](./data/img.PNG)


Preparation
-----------
You need to prepare two xlsx files to run this program.

- Query file (Include specimen query information)
- Data file (Include informations about specimen collection and identification)

Please download sample file for more details.


Usage
-----
1. For quick use, you can download `specimen_info_gui.py` and double click.
   You will get a graphical user interface.

   - Select valid query xlsx file (default: query.xlsx);
   - Select valid data xlsx file (default: data.xlsx);
   - Change output name if you want;
   - Click **Start Query** button to start.

   After execution, an .xlsx file and an SQLite3 db file which contains the
   detailed specimen infomations will be generated.

2. For user who are familiar with console, you can download `specimen_info.py`.
   At console or command line, type this:

        python specimen_info.py -i query.xlsx -d data.xlsx -o outfile.xlsx

   If you changed your query file and data file to default name:

   - query file: query.xlsx
   - data file: data.xlsx

   Then you can just type:

        python specimen_info.py

   After execution, an .xlsx file and an SQLite3 db file which contains the
   detailed specimen infomations will be generated.

3. For extented use: If you just want to get the output tuple and want to save
   output information to other places (for example, MySQL), do this:

        from specimen_info import (Query, write_to_xlsx_file,
                                  write_to_sqlite3)

        q = Query(query_file=query_filename, offline_data_file=data_filename)
        out_tuple_list = q.do_multi_query()

        # If you want to save output to xlsx file
        write_to_xlsx_file(out_tuple_list, xlsx_outfile_name="specimen.xlsx")

        # If you want to save output to SQLite3 db file
        write_to_sqlite3(out_tuple_list, sqlite3_file="specimen.sqlite")

        # If you want to save to other places,
        # Just write your extension code.


# trmxlx2nmea
it's a python script for convert Trimble xlx format file into NMEA (GGA) format file

some customer claim that they only have a evaluation tools on NMEA format file, 
for our Trimlbe autonomy B5000 unit, it can only output xlx file, this script may help, 
currently it just a draft version, not tested by those tools yet.

0. it's a windows base script
1. install python
2. install openpyxl (pip install openpyxl)
3. install pynmea2 (pip install pynmea2)
4. save xlx as xlsx (openpyxl can only parse xlsx format)
5. edit xlx2nmea.py to your test path and xlsx file name
6. run it (xlx2nmea.py)

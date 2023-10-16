# trmxlx2nmea
it's a python script for convert Trimble xlx format file into NMEA (GGA/RMC/GSV/GSA) format file

some customer claim that they only have a evaluation tools on NMEA format file, 
for our Trimlbe autonomy B5000 unit, it can only output xlx file, this script may help.

. it's a windows base script

. install python

. install openpyxl (pip install openpyxl)

. save xlx as xlsx (openpyxl can only parse xlsx format)

. run it (directly run xlx2nmea.py, or, you may use "pyinstaller --onefile" to generate EXE file, the EXE file was too big to be put on GitHub)

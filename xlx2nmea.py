"""
NMEA-0183 GGA (https://receiverhelp.trimble.com/alloy-gnss/en-us/NMEA-0183messages_GGA.html)

Field	Meaning
0	    Message ID $GPGGA

1	    UTC of position fix

2	    Latitude

3	    Direction of latitude:
        N: North
        S: South

4	    Longitude

5	    Direction of longitude:
        E: East
        W: West

6	    GPS Quality indicator:
        0: Fix not valid
        1: GPS fix
        2: Differential GPS fix (DGNSS), SBAS, OmniSTAR VBS, Beacon, RTX in GVBS mode
        3: Not applicable
        4: RTK Fixed, xFill
        5: RTK Float, OmniSTAR XP/HP, Location RTK, RTX
        6: INS Dead reckoning

7	    Number of SVs in use, range from 00 through to 24+

8	    HDOP

9	    Orthometric height (MSL reference)

10	    M: unit of measure for orthometric height is meters

11	    Geoid separation

12	    M: geoid separation measured in meters

13	    Age of differential GPS data record, Type 1 or Type 9. Null field when DGPS is not used.

14	    Reference station ID, range 0000 to 4095. 
        A null field when any reference station ID is selected and no corrections are received. 
        Reference Station ID	Service
        0002	                CenterPoint or ViewPoint RTX
        0005	                RangePoint RTX
        0006	                FieldPoint RTX
        0100	                VBS
        1000	                HP
        1001	                HP/XP (Orbits)
        1002	                HP/G2 (Orbits)
        1008	                XP (GPS)
        1012	                G2 (GPS)
        1013	                G2 (GPS/GLONASS)
        1014	                G2 (GLONASS)
        1016	                HP/XP (GPS)
        1020	                HP/G2 (GPS)
        1021	                HP/G2 (GPS/GLONASS) 

15	    The checksum data, always begins with *
"""

import os
import openpyxl
import pynmea2

xlx_path = 'C:\\Work\\Tools\\XLX2NMEA\\'
xlx_name = '22020831_sv_example'
xlx_tail = '.xlsx'
nmea_tail = '.log'

xlx_wb = openpyxl.load_workbook(xlx_path+xlx_name+xlx_tail)
xlx_sht = xlx_wb[xlx_wb.sheetnames[0]]

# Open a file for exclusive creation. If the file already exists, the operation fails.
with open(xlx_path+xlx_name+nmea_tail, 'x') as nmea_log:
    for i in range(2, xlx_sht.max_row):
        msg = pynmea2.GGA('GP', 'GGA', 
                          (str(xlx_sht.cell(row=i, column=4).value),        # Time of Day (sec UTC)
                           str(xlx_sht.cell(row=i, column=10).value), 'N',  # Latitude (deg)
                           str(xlx_sht.cell(row=i, column=9).value), 'E',   # Longitude (deg)
                           str(xlx_sht.cell(row=i, column=18).value),       # Fix Type
                           str(xlx_sht.cell(row=i, column=41).value),       # SVs (used) 
                           str(xlx_sht.cell(row=i, column=34).value),       # HDOP  
                           str(xlx_sht.cell(row=i, column=11).value), 'M',  # Altitude (m MSL)
                           '0', 'M',
                           str(xlx_sht.cell(row=i, column=20).value),       # DGPS
                           #xlx_sht.cell(row=i, column=36).value,           # Age of Corrections (sec) 
                           '0002'))
        nmea_log.write(str(msg))
        nmea_log.write('\n')

print('Successful convert XLX into NMEA log (',xlx_wb.sheetnames[0],')')

"""
NMEA-0183 GGA (https://receiverhelp.trimble.com/alloy-gnss/en-us/NMEA-0183messages_GGA.html)
               https://orolia.com/manuals/VSP/Content/NC_and_SS/Com/Topics/APPENDIX/NMEA_GGAmess.htm
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
xver = '0.2'
import os
import openpyxl
import pynmea2

xlx_path = 'C:\\Work\\Tools\\XLX2NMEA\\'
#xlx_path = '/Users/Hawk/Downloads/'
xlx_name = 'small_example'
xlx_tail = '.xlsx'
nmea_tail = '.txt'

xlx_wb = openpyxl.load_workbook(xlx_path+xlx_name+xlx_tail)
xlx_sht = xlx_wb[xlx_wb.sheetnames[0]]

nmea_UTC    = 4     #Time of Day (sec UTC)
nmea_La     = 10    #Latitude (deg)
nmea_Lg     = 9     #Longitude (deg)
nmea_GQI    = 18    #Fix Type
nmea_SV     = 18    #SVs (used) 
nmea_HDOP   = 34    #HDOP 
nmea_OH     = 11    #Altitude (m MSL)
nmea_ADGPS  = 20    #DGPS
#nmea_ADGPS  = 36    #Age of Corrections (sec)

# Open a file for exclusive creation. If the file already exists, the operation fails.
with open(xlx_path+xlx_name+nmea_tail, 'x') as nmea_log:
    for i in range(2, xlx_sht.max_row):
        
        La_data = xlx_sht.cell(i, nmea_La).value
        if La_data > 0:
            La_dir = 'N'
        else:
            La_dir = 'S'
        La_data = int(La_data)*100+round((La_data - int(La_data))*60,3)

        Lg_data = xlx_sht.cell(i, nmea_Lg).value
        if  Lg_data > 0:
            Lg_dir = 'E'
        else:
            Lg_dir = 'W'
            Lg_data *= -1
        Lg_data = int(Lg_data)*100+round((Lg_data - int(Lg_data))*60,3)

        msg = pynmea2.GGA('GP', 'GGA',
                          (str(xlx_sht.cell(i, nmea_UTC).value),
                           str(La_data).zfill(4), La_dir,
                           str(Lg_data).zfill(5), Lg_dir,
                           str(xlx_sht.cell(i, nmea_GQI).value),
                           str(xlx_sht.cell(i, nmea_SV).value).zfill(2),
                           str(round(xlx_sht.cell(i, nmea_HDOP).value,1)),
                           str(round(xlx_sht.cell(i, nmea_OH).value,1)), 'M',
                           '0.0', 'M',
                           str(xlx_sht.cell(i, nmea_ADGPS).value),
                           '0002'))
        nmea_log.write(str(msg))
        nmea_log.write('\n')

print('Version:',xver)
print('Successful convert XLX (',xlx_sht.max_row-1,'lines) into NMEA log:')
print(xlx_path+xlx_name+nmea_tail)

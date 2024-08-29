#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      ak34
#
# Created:     16/03/2022
# Copyright:   (c) ak34 2022
# Licence:     <your licence>
# Py Version    Python 3.7.3
#-------------------------------------------------------------------------------

import pyodbc
import os.path
import xlwt
import datetime
import database_conn as dbconn
#from dateutil.relativedelta import relativedelta

def writeData(fo, path, filename, linestring, header):
    if fo is None:
        if os.path.isfile(path + "/" + filename):
            fo = open(path + "/" + filename, 'a')
        else:
            fo = open(path + "/" + filename, 'w')
            #write header of output file
            fo.write(header)

    fo.write(linestring)
    fo.close()
    return True

def checkNone(string_value, fieldname = '', throwError = False):
    if string_value == 'None':
        if throwError:
            print('***ERROR -- VALUE MISSING in field {}'.format(fieldname))
            return ''
        else:
            return ''
    else:
        return string_value


def getdatetime():
    d = datetime.datetime.now()
    return d.strftime("%y") + d.strftime("%m") + d.strftime("%d") + "_" + d.strftime("%H") + d.strftime("%M")

def writeToExcel(row, fieldValues, style, vba_sheet = True):

    j = 0

    for part in fieldValues:
        if vba_sheet == True:
            sheet.write(row, j, part, style)
        else:
            sheet2.write(row, j, part, style)
        j += 1

    return True

##=========================================================================================================================
##=========================================================================================================================
##=========================================================================================================================
##=========================================================================================================================
##=========================================================================================================================
##=========================================================================================================================
##=========================================================================================================================
##=========================================================================================================================
##=========================================================================================================================


##.........................................................................................................................
##.........................................................................................................................
##.............................. USER DATA ................................................................................
##.........................................................................................................................
##.........................................................................................................................

outpath = dbconn.outpath

sql = dbconn.sql

conn = dbconn.conn

##.........................................................................................................................
##.........................................................................................................................
##.........................................................................................................................

fo = None
sheet = None
sheet2 = None
sheet3 = None

cursor = conn.cursor()
cursor.execute(sql)
input_count = 0


##### TO FIND DRIVER NAMES print(pyodbc.drivers()) ###########

##header = 'VBA Project ID,VBA Permit ID,VBA Reference Document ID,VBA Login Name(s) (comma separated),Site ID,DO NOT USE,DO NOT USE,Drainage Division,Riverbasin No.,Waterbody ID,Tributary of,Site Name,Location description,Survey Name,Start date,End date,Type of coordinate system,GPS used (y, n),Datum (g, a, w),X-coordinate (easting or longitude),Y-coordinate (northing or latitude),Mapsheet number (1:100k),MGA Zone (54 or 55),Positional accuracy (m),Altitude (m),Users own reference number,Method Code,Method details code,Sampling details - Date,Sampling details - Numeric Value,Sampling details - Text value,Taxon Name,Taxon Common Name,VBA Taxon ID,Type of Observation,Count,Count Qualifier (comma separated),Weight range (g),Length range (mm),Reproductive condition,Disease; parasites; deformities,Origin of Taxon at site,Specimen sent to,Specimen reference number,Max Depth (m) (of sample area),Ave. Depth (m) (of sample area),Depth of measurement (m),EC measured at Water Temp,EC measured at 25 °C,Water temperature °C,Dissolved Oxygen (mg/L),pH,Turbidity ntu,Secchi depth (m),Hardness (CaCO3 ppm),Comments on Sampling site,Observer\n'
vbaAquaticFields = ["VBA Project ID","VBA Permit ID","VBA Reference Document ID","VBA Login Name(s) (comma separated)","Site ID","DO NOT USE","DO NOT USE","Drainage Division","Riverbasin No.","Waterbody ID","Tributary of","Site Name","Location description","Survey Name","Start date","End date","Type of coordinate system","GPS used (y, n)","Datum (g, a, w)","X-coordinate (easting or longitude)","Y-coordinate (northing or latitude)","Mapsheet number (1:100k)","MGA Zone (54 or 55)","Positional accuracy (m)","Altitude (m)","Users own reference number","Method Code","Method details code","Sampling details - Date","Sampling details - Numeric Value","Sampling details - Text value","Taxon Name","Taxon Common Name","VBA Taxon ID","Type of Observation","Count","Count Qualifier (comma separated)","Weight range (g)","Length range (mm)","Reproductive condition","Disease; parasites; deformities","Origin of Taxon at site","Specimen sent to","Specimen reference number","Max Depth (m) (of sample area)","Ave. Depth (m) (of sample area)","Depth of measurement (m)","EC measured at Water Temp","EC measured at 25 °C","Water temperature °C","Dissolved Oxygen (mg/L)","pH","Turbidity ntu","Secchi depth (m)","Hardness (CaCO3 ppm)","Comments on Sampling site","Observer"]

vbaMarineFields =["VBA Project ID","VBA Permit ID","VBA Reference Document ID","VBA Observer ID (comma seperated)","Site ID","Restricted site","Reason for restriction","Site Name","Location description","Survey Name","Start date","End date","Type of co-ordinate system","GPS used (Y, N)","Datum","X-coordinate","Y-coordinate","Mapsheet number","MGA Zone","Positional accuracy (m)","Users own reference number","Low Tide - Time (HHMM)","Low Tide - Height","High Tide - Time (HHMM)","High Tide - Height","Swell Height","Surge","AME Dive Number","Sea Scale","Wind Speed (km/hr)","Method","Method details","Sampling details - Date","Sampling details - Numeric Value","Sampling details - Text value","Taxon Name","Taxon Common Name","Taxon Code","Type of Observation","Count","Cover abundence","Count Qualifyer (comma seperated)","Origin of Taxon at site","Specimen sent to","Specimen reference number","Observer"]


##writeData(fo, outpath, file, outline, header)


ts = getdatetime()
wb = None

prev_site_name = -1
prev_discipline = ''
outputfile = ''


columns = [column[0] for column in cursor.description]
last_survey_id = -1

for row in cursor:
    project_id = row[0]
    site_name  = '{0} | {1}'.format(row[10], row[11])
    discipline = (row[45]).lower()
    input_count += 1


    if prev_discipline != discipline:
        if outputfile != '':
            wb.save(outpath + outputfile)

        wb = None
        sheet = None
        rw1 = 1
        rw2 = 1
        outputfile = "vbaProject_" + str(project_id)  + "_import_" + discipline.replace(' ', '_') + "_" + ts + ".xls"
        print('***' + outputfile)
        styleNorm = xlwt.easyxf()
        encoding = 'latin1'
        wb = xlwt.Workbook(encoding=encoding)
        sheet=xlwt.Workbook()
##        sheetname = 'VBA_BATCH_READY_01'
##        sheetname2 = 'AAE_DB_DATA'
        sheet = wb.add_sheet('VBA_BATCH_READY_01')
        sheet2 = wb.add_sheet('AAE_DB_DATA')
        sheet3 = wb.add_sheet('AAE_DB_SQL')
        sheet3.write(0, 0, '{0} WHERE lower(primary_discipline) LIKE \'{1}\''.format(sql, discipline), styleNorm)

        #write header line for worksheet
        ##                   print str(6)

        if discipline == 'marine':
            writeToExcel(0,vbaMarineFields, styleNorm)
        else:
            writeToExcel(0,vbaAquaticFields, styleNorm)

        writeToExcel(0,columns, styleNorm, False)

        wb.save(outpath + outputfile)
        prev_discipline = discipline


    if (site_name != prev_site_name) or (last_survey_id != row[49]): ## and row[28] != 175: #new site only with sample info (175 = No Fish)

        print('{0} | {1}'.format(row[10], row[11]))
        last_method_idx = -1
        last_survey_id = row[49]


        for i in range(21, 26):
            if str(row[i]) != 'None':
                last_method_idx = i

##        print('method {0}: last_method {1}'.format(method_idx,last_method_idx))

        if last_method_idx > 0: #and method_idx != last_method_idx: #populate method info

            first_method_detail = -1
            for i in range(20, 26):
##                if i > method_idx:
                if str(row[i]) != 'None':
                    sample_num = '0'
                    if i == 20:
                        method_detail = '30'
##                        sample_num = str(row[i])
                        sample_date = ''
                    elif i == 21:
                        method_detail = '1'
                        sample_num = ''
                        sample_date = str(row[i])
                    elif i == 22:
                        method_detail = '2'
                        sample_num = ''
                        sample_date = str(row[i])
                    elif i == 23:
                        method_detail = '3'
##                        sample_num = str(row[i])
                        sample_date = ''
                    elif i == 24:
                        method_detail = '15'
##                        sample_num = str(row[i])
                        sample_date = ''
                    elif i == 25:
                        method_detail = '4'
                        sample_num = str(row[i])
                        sample_date = ''
                    elif i == 26:
                        method_detail = '5'
                        sample_num = str(row[i])
                        sample_date = ''

##                    print(i)
                    if first_method_detail < 1:
                        writeToExcel(rw2, row, styleNorm, False)
                        rw2 += 1
##                        print(row[29])

                    if first_method_detail < 0 and i != last_method_idx:


                        if discipline == 'marine':
                            outline = [str(row[0]),"","",str(row[3]),checkNone(str(row[4])),"","",str(row[9]),str(row[10]),str(row[11]),str(row[12]),str(row[52]),str(row[13]),str(row[14]),str(row[15]),str(row[16]),str(row[17]),"","",str(row[18]),"","","","","","","","","","",str(row[19]),method_detail,sample_date,sample_num,"","","","","","","","","","","",""]

                        else:

                            outline = [str(row[0]),"","",row[3],"","","",str(row[5]),str(row[6]),checkNone(str(row[7]),'vba_waterbody_id', True),row[8],row[9],row[10],row[11],str(row[12]),str(row[52]),row[13],row[14],row[15],str(row[16]),str(row[17]),"","",str(row[18]),"","",str(row[19]),method_detail,sample_date,sample_num,"","","","","","","","","","","","","","",checkNone(str(row[33])),checkNone(str(row[34])),checkNone(str(row[35])),checkNone(str(row[36])),checkNone(str(row[37])),checkNone(str(row[38])),checkNone(str(row[39])),checkNone(str(row[40])),checkNone(str(row[41])),checkNone(str(row[42])),checkNone(str(row[43])),"",""]

                        first_method_detail = 1

                    elif first_method_detail < 0 and i == last_method_idx:

                        if discipline == 'marine':
                             outline = [str(row[0]),"","",str(row[3]),checkNone(str(row[4])),"","",str(row[9]),str(row[10]),str(row[11]),str(row[12]),str(row[52]),str(row[13]),str(row[14]),str(row[15]),str(row[16]),str(row[17]),"","",str(row[18]),"","","","","","","","","","",str(row[19]),method_detail,sample_date,sample_num,"",str(row[27]),str(row[27]),str(row[28]),str(row[29]),str(row[30]),"","","","","",str(row[44])]

                        else:

                            outline = [str(row[0]),"","",row[3],"","","",str(row[5]),str(row[6]),checkNone(str(row[7]),'vba_waterbody_id', True),row[8],row[9],row[10],row[11],str(row[12]),str(row[52]),row[13],row[14],row[15],str(row[16]),str(row[17]),"","",str(row[18]),"","",str(row[19]),method_detail,sample_date,sample_num,"",row[27],row[27],str(row[28]),row[29],str(row[30]),"",checkNone(str(row[31])),checkNone(str(row[32])),"","","","","",checkNone(str(row[33])),checkNone(str(row[34])),checkNone(str(row[35])),checkNone(str(row[36])),checkNone(str(row[37])),checkNone(str(row[38])),checkNone(str(row[39])),checkNone(str(row[40])),checkNone(str(row[41])),checkNone(str(row[42])),checkNone(str(row[43])),"",row[44]]

                        first_method_detail = 1
                    else:
                        if i != last_method_idx: #write method without site or sample info
                            if discipline == 'marine':
                                outline = ["","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","",method_detail,sample_date,sample_num,"","","","","","","","","","","",""]
                            else:

                                outline = ["","","","","","","","","","","","","","","","","","","","","","","","","","","",method_detail,sample_date,sample_num,"","","","","","","","","","","","","","","","","","","","","","","","","","",""]

##                                writeToExcel(rw, outline, styleNorm)
##                                writeToExcel(rw, row, styleNorm, False)
##                                rw += 1


                        else: #last method code written also has first sample record
                            if discipline == 'marine':

                                outline = ["","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","",method_detail,sample_date,sample_num,"",str(row[27]),str(row[27]),str(row[28]),str(row[29]),str(row[30]),"","","","","",str(row[44])]

                            else:
                                outline = ["","","","","","","","","","","","","","","","","","","","","","","","","","","",method_detail,sample_date,sample_num,"",row[27],row[27],str(row[28]),row[29],str(row[30]),"",checkNone(str(row[31])),checkNone(str(row[32])),"","","","","","","","","","","","","","","","","",row[44]]

##                                writeToExcel(rw, outline, styleNorm)
##                                writeToExcel(rw, row, styleNorm, False)
##                                rw += 1

                    writeToExcel(rw1, outline, styleNorm)
                    rw1 += 1
    ##                    else:
##                        print('skipped None method {0}'.format(i))


        else: #no method info so start with first sample

            if discipline == 'marine':
                outline = [str(row[0]),"","",str(row[3]),checkNone(str(row[4])),"","",str(row[9]),str(row[10]),str(row[11]),str(row[12]),str(row[52]),str(row[13]),str(row[14]),str(row[15]),str(row[16]),str(row[17]),"","",str(row[18]),"","","","","","","","","","",str(row[19]),"","","","",str(row[27]),str(row[27]),str(row[28]),str(row[29]),str(row[30]),"","","","","",str(row[44])]

            else:
                outline = [str(row[0]),"","",row[3],"","","",str(row[5]),str(row[6]),str(row[7]),row[8],row[9],row[10],row[11],str(row[12]),str(row[52]),row[13],row[14],row[15],str(row[16]),str(row[17]),"","",str(row[18]),"","",str(row[19]),"","","","",row[27],row[27],str(row[28]),row[29],str(row[30]),"",checkNone(str(row[31])),checkNone(str(row[32])),"","","","","","","","","","","","","","","","","",row[44]]

            writeToExcel(rw1, outline, styleNorm)
            writeToExcel(rw2, row, styleNorm, False)
            rw1 += 1
            rw2 += 1

        prev_site_name = site_name

    else: #subsequent samples after first
##        if row[28] != 175:
        if discipline == 'marine':
            outline = ["","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","",str(row[27]),str(row[27]),str(row[28]),str(row[29]),str(row[30]),"","","","","",str(row[44])]
        else:
            outline = ["","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","",row[27],row[27],str(row[28]),row[29],str(row[30]),"",checkNone(str(row[31])),checkNone(str(row[32])),"","","","","","","","","","","","","","","","","",row[44]]

        writeToExcel(rw1, outline, styleNorm)
        writeToExcel(rw2, row, styleNorm, False)
        rw1 += 1
        rw2 += 1

print('*** {0} rows processed'.format(str(input_count)))
print('')
print('ENSURE surveys updated to SUBMITTED level in vba_submission column (VBA_SUBMITTED_RECORDS_UPDATE.sql)')

wb.save(outpath + outputfile)
cursor.close()
conn.close()


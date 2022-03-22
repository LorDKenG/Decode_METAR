from ast import If
from cgi import test
import os
import re 
import latlon
import datetime
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm, Inches, Mm, Emu
from datetime import date, timedelta


# TIME for default (last month)
lastMonth = date.today().replace(day=1) - timedelta(days=1)

class decode:
    
    
    # input variable
    def __init__(self, code = 'VTBS', mon = lastMonth.strftime("%m") , year = lastMonth.strftime("%Y")):
        self.code = code
        self.mon = mon
        self.year = year
        self.filename_in = mon+"_"+code
        self.CSV_path_file = "Result_CSV_"+code+"_"+mon+"_"+year
        self.Climate_path_file = "Climatological_"+code+"_"+mon+"_"+year

    # decode METAR
    def filerw(self):
        # Read file
        os.makedirs(self.CSV_path_file, exist_ok=True)
        rawmetar = open(self.filename_in, encoding="utf8", errors='ignore')
        splite_file = open(self.CSV_path_file+'/'+self.filename_in+"_decode.csv", "w")
        splite_file.writelines("AIRPOT,DAY,TIME(HHMM) UTC,WIND DIRECTION,WIND SPEED,VARY to,VARY,GUST,VISIBILITY MINIMUM,VISIBILITY MINIMUM in DIRECTION,STATUSRVR,RVRVARY,RUNWAY,VISRUNWAY1,VISRUNWAY2,RVRFUTURE,CLOUD1,HIGH1,CLOUD2,HIGH2,CLOUD3,HIGH3,CLOUD4,HIGH4,WEATHER,TEMPURATURE,DEWPOINT,PRESSURE,WEATHER"+'\n')
        outputfileA = open(self.CSV_path_file+'/'+self.filename_in+"_model-A.csv", "w")
        outputfileA.writelines("TIME,VIS,HIGHofCLD\n")
        outputfileB = open(self.CSV_path_file+'/'+self.filename_in+"_model-B.csv", "w")
        outputfileB.writelines("TIME,VIS\n")
        outputfileC = open(self.CSV_path_file+'/'+self.filename_in+"_model-C.csv", "w")
        outputfileC.writelines("TIME,CLD(ft/100),TPYE\n")
        outputfileD = open(self.CSV_path_file+'/'+self.filename_in+"_model-D.csv", "w")
        outputfileD.writelines("WIND_ANG,WIND_SP\n")
        outputfileE = open(self.CSV_path_file+'/'+self.filename_in+"_model-E.csv", "w")
        outputfileE.writelines("TIME,TEMP\n")
        outputfileF = open(self.CSV_path_file+'/'+self.filename_in+"_model-F.csv", "w")
        outputfileF.writelines("DATE,GUST_SP\n")
        outputfileG = open(self.CSV_path_file+'/'+self.filename_in+"_model-G.csv", "w")
        outputfileG.writelines("TIME,VIS_RUNWAY,HIGHofCLD\n")

        # skip line
        datametar = rawmetar.readlines()[3:]
        for lines in datametar:
            # airport and time
            AIRP = re.findall("VT\w{2}", lines)
            DTZ = re.findall("\d{6}Z", lines)
            DD = DTZ[0][0:2]
            HH = DTZ[0][2:4]
            MM = DTZ[0][4:6]

            # wind
            if re.findall("\d{5}KT", lines):
                WIND = re.findall("\d{5}KT", lines)
                GUSTW = "No"
                GUSTWV = "No"
                WVARY = "No"
                WD = WIND[0][0:3]
                WS = WIND[0][3:5]
                WV1 = "none"
                WV2 = "none"
            elif re.findall("VRB\d{2}KT", lines):
                WIND = re.findall("VRB\d{2}KT", lines)
                GUSTW = "No"
                GUSTWV = "No"
                WVARY = "Yes"
                WD = "VRB"
                WS = WIND[0][3:5]
            elif re.findall("\d{5}G\d{2}KT", lines):
                WIND = re.findall("\d{5}G\d{2}KT", lines)
                GUSTW = "Yes"
                GUSTWV = WIND[0][6:8]
                WVARY = "No"
                WD = WIND[0][0:3]
                WS = WIND[0][3:5]
                WV1 = "none"
                WV2 = "none"
            elif re.findall("VRB\d{2}G\d{2}KT", lines):
                WIND = re.findall("VRB\d{2}G\d{2}KT", lines)
                GUSTW = "YES"
                GUSTWV = WIND[0][6:8]
                WVARY = "YES"
                WD = "VRB"
                WS = WIND[0][3:5]
            else:
                GUSTW = "YES"
                GUSTWV = "No"
                WVARY = "YES"
                WD = "-"
                WS = "-"           

            # vary wind    
            if  re.findall("\d{3}V\d{3}", lines):
                WV = re.findall("\d{3}V\d{3}", lines)
                WV1 = WV[0][0:3]
                WV2 = WV[0][4:7]
            else:
                WV1 = "none"
                WV2 = "none"

            # visibility
            if  re.findall("9999", lines):
                VIS = re.findall("\s\d{4}\s", lines)
                VISV1 = VIS[0][1:5]
                VISD = "none"

            elif re.findall("(?:CAVOK)", lines):
                VISV1 = "9999"
                VISD = "none"
        
            elif  re.findall("\s\d{4}(?:NE|NW|SW|SE|N|S|E|W)", lines):
                VIS = re.findall("\s\d{4}(?:NE|NW|SW|SE|N|S|E|W)", lines)
                VISV1 = VIS[0][1:5]
                VISD = VIS[0][5:]
            
            elif re.findall("\s\d{3}0\s", lines):
                VIS = re.findall("\s\d{3}0\s", lines)
                VISV1 = VIS[0][1:5]
                VISD = "none"

            # RVR
            if re.findall("R\d{2}/\d{4}[U,D,N,\s]", lines):
                RVR = re.findall("R\d{2}/\d{4}[U,D,N,\s]", lines)
                STATUSRVR = "Yes"
                RVRVARY = "No"
                RUNWAY = RVR[0][1:3]
                VISRUNWAY1 = RVR[0][4:8]
                VISRUNWAY2 = RVR[0][4:8]
                RVRFUTURE = RVR[0][8:9]
        
            elif re.findall("R\d{2}/\d{4}V\d{4}[U,D,N,\s]", lines):
                RVR = re.findall("R\d{2}/\d{4}V\d{4}[U,D,N,\s]", lines)
                STATUSRVR = "Yes"
                RVRVARY = "Yes"
                RUNWAY = RVR[0][1:3]
                VISRUNWAY1 = RVR[0][4:8]
                VISRUNWAY2 = RVR[0][9:13]
                RVRFUTURE = RVR[0][13:14]
        
            else:
                STATUSRVR = "No"
                RVRVARY = "No"
                RUNWAY = "No"
                VISRUNWAY1 = "No"
                VISRUNWAY2 = "No"
                RVRFUTURE = "No"
    
            # cloud
            CLOUD = re.findall("(?:FEW|SCT|BKN|OVC)\d{3}(?:TCU|CB|\s)", lines)
                
            if len(CLOUD) == 4 :
                CL1 = CLOUD[0][0:3]
                CLD1 = CLOUD[0][3:6]
                CLS1 = CLOUD[0][6:]
                CL2 = CLOUD[1][0:3]
                CLD2 = CLOUD[1][3:6]
                CLS2 = CLOUD[1][6:]
                CL3 = CLOUD[2][0:3]
                CLD3 = CLOUD[2][3:6]
                CLS3 = CLOUD[2][6:]
                CL4 = CLOUD[3][0:3]
                CLD4 = CLOUD[3][3:6]
                CLS4 = CLOUD[3][6:]
            elif len(CLOUD) == 3 :
                CL1 = CLOUD[0][0:3]
                CLD1 = CLOUD[0][3:6]
                CLS1 = CLOUD[0][6:]
                CL2 = CLOUD[1][0:3]
                CLD2 = CLOUD[1][3:6]
                CLS2 = CLOUD[1][6:]
                CL3 = CLOUD[2][0:3]
                CLD3 = CLOUD[2][3:6]
                CLS3 = CLOUD[2][6:]
                CL4 = "none"
                CLD4 = "none"
                CLS4 = "none"
            elif len(CLOUD) == 2 :
                CL1 = CLOUD[0][0:3]
                CLD1 = CLOUD[0][3:6]
                CLS1 = CLOUD[0][6:]
                CL2 = CLOUD[1][0:3]
                CLD2 = CLOUD[1][3:6]
                CLS2 = CLOUD[1][6:]
                CL3 = "none"
                CLD3 = "none"
                CLS3 = "none"
                CL4 = "none"
                CLD4 = "none"
                CLS4 = "none"
            elif len(CLOUD) == 1 :
                CL1 = CLOUD[0][0:3]
                CLD1 = CLOUD[0][3:6]
                CLS1 = CLOUD[0][6:]
                CL2 = "none"
                CLD2 = "none"
                CLS2 = "none"
                CL3 = "none"
                CLD3 = "none"
                CLS3 = "none"
                CL4 = "none"
                CLD4 = "none"
                CLS4 = "none"
            elif len(CLOUD) == 0 :
                CL1 = "none"
                CLD1 = "none"
                CLS1 = "none"
                CL2 = "none"
                CLD2 = "none"
                CLS2 = "none"
                CL3 = "none"
                CLD3 = "none"
                CLS3 = "none"
                CL4 = "none"
                CLD4 = "none"
                CLS4 = "none"


            if re.findall("(?:SKC|CAVOK|N.C)", lines):
                VC = re.findall("(?:SKC|CAVOK|N.C)", lines)
                VCC = VC[0][0:]
            else:
                VCC = "none"

            # TEMPURATUE
            TEMDEW = re.findall("\s\d{2}/\d{2}\s", lines)
            TEM = TEMDEW[0][1:3]
            DEW = TEMDEW[0][4:6]

            if re.findall("(?:BKN|OVC)\d{3}", lines):
                modelc = re.findall("(?:BKN|OVC)\d{3}", lines)
                modelc2 = modelc[0][0:3]
                modelc1 = modelc[0][3:6]
            else:
                modelc = '000000'
                modelc2 = 'non'
                modelc1 = 'non'

            # PRESSURE
            if  re.findall("Q\d{4}", lines):
                PP = re.findall("Q\d{4}", lines)
                PP1 = PP[0][1:5]
            else:
                PP1 = ""

            # WEATHER
            if re.findall('\w+\W+\w+\W+\w+\W+\w+\W+\w+\W+\w+\W+\w+\W+Q}', lines):
                CUTDATAF = re.findall('\w+\W+\w+\W+\w+\W+\w+\W+\w+\W+\w+\W+\w+\W+Q', lines)
            else:
                CUTDATAF = ['0']

            if re.findall("\s\D{1}TSRA\s|\sTSRA\s", CUTDATAF[0]):
                WH = re.findall("\s\D{1}TSRA\s|\sTSRA\s", CUTDATAF[0])
            elif re.findall("\s\D{3}\s|\s\D{2}\s", CUTDATAF[0]):
                WH = re.findall("\s\D{3}\s|\s\D{2}\s", CUTDATAF[0])
            else:
                WH = [' - ']
   

            splite_file.writelines("{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{}\n".format(AIRP[0],DD,HH+MM,WD,WS,WV1,WV2,GUSTWV,VISV1,VISD,STATUSRVR,RVRVARY,RUNWAY,VISRUNWAY1,VISRUNWAY2,RVRFUTURE,CL1,CLD1,CL2,CLD2,CL3,CLD3,CL4,CLD4,VCC,TEM,DEW,PP1,WH[0][1:-1]))
            outputfileA.writelines(HH+MM+","+VISV1+","+CLD1+'\n')
            outputfileB.writelines(HH+MM+","+VISV1+'\n')
            outputfileC.writelines(HH+MM+","+modelc1+","+modelc2+'\n')
            outputfileD.writelines(WD[0:2]+","+WS+'\n')
            outputfileE.writelines(HH+MM+","+TEM+'\n')
            outputfileF.writelines(DD+","+GUSTWV+'\n')
            outputfileG.writelines(HH+MM+","+VISRUNWAY1+","+modelc1+'\n')
            lines = rawmetar.readline()

        rawmetar.close()
        splite_file.close()
        outputfileA.close()
        outputfileB.close()
        outputfileC.close()
        outputfileD.close()
        outputfileE.close()
        outputfileF.close()
        outputfileG.close()

    # Frequencies (percent) of the Occurrence of Visibility (in meters) and|or the Height of the Base of 
    # the Lowest Cloud Layer (in feet) Extent below specified values at specified times
    def model_a(self):

        os.makedirs(self.Climate_path_file, exist_ok=True)
        a = "0000 0030 0100 0130 0200 0230 0300 0330 0400 0430 0500 0530 0600 0630 0700 0730 0800 0830 0900 0930 1000 1030 1100 1130 1200 1230 1300 1330 1400 1430 1500 1530 1600 1630 1700 1730 1800 1830 1900 1930 2000 2030 2100 2130 2200 2230 2300 2330"
        aa = a.split()
        outputfileA = open(self.CSV_path_file+'/'+self.filename_in+"_model-A-result.csv", "w")

        for i in aa:
            inputfileA = open(self.CSV_path_file+'/'+self.filename_in+"_model-A.csv", "r")
            linesAA = inputfileA.readline()
            linesAAA = linesAA.replace("none","0")
            t1=t2=t3=t4=t5=t6=t7=t8=t9=total=0
            dataAA = linesAAA.split(',')
            while linesAAA:
                if dataAA[0] == i :
                    if 0 < int(dataAA[1]) < 500 or 0 < int(dataAA[2]) < 2:
                        t1 = t1 + 1
                    elif 0 < int(dataAA[1]) < 800 or 0 < int(dataAA[2]) < 3:
                        t2 = t2 + 1
                    elif 0 < int(dataAA[1]) < 1500 or 0 < int(dataAA[2]) < 5:
                        t3 = t3 + 1
                    elif 0 < int(dataAA[1]) < 3000 or 0 < int(dataAA[2]) < 10:
                        t4 = t4 + 1
                    elif 0 < int(dataAA[1]) < 5000 or 0 < int(dataAA[2]) < 15:
                        t5 = t5 + 1
                    elif 0 < int(dataAA[1]) < 8000 or 0 < int(dataAA[2]) < 20:
                        t6 = t6 + 1
                    else:
                        total = total+1  

                    linesAA = inputfileA.readline()
                    linesAAA = linesAA.replace("none","0")
                    dataAA = linesAAA.split(',')
 
                else:
                    linesAA = inputfileA.readline()
                    linesAAA = linesAA.replace("none","0")
                    dataAA = linesAAA.split(',') 

            if t1+t2+t3+t4+t5+t6+total != 0:
                outputfileA.writelines("{},{},{},{},{},{},{},{}\n".format(i,t1,t1+t2,t1+t2+t3,t1+t2+t3+t4,t1+t2+t3+t4+t5,t1+t2+t3+t4+t5+t6,t1+t2+t3+t4+t5+t6+total))

            inputfileA.close()

        outputfileA.close()

        #Import template document
        path_script = os.path.dirname(os.path.realpath(__file__))
        template = DocxTemplate(os.path.join(path_script, 'temp_model_a.docx'))
        result = open(self.CSV_path_file+'/'+self.filename_in+"_model-A-result.csv", "r")

        code = self.code

        #TIME
        monb = self.mon
        mon = datetime.datetime.strptime(monb, "%m").strftime("%B")
        year = self.year

        #Import list airport
        air_list = open(os.path.join(path_script, 'airports.dat'), 'r')

        #find data of aiport
        lists = air_list.readline()
        while lists:
            if code in lists:
                airport = lists[0:-1]
                break
            else:
                lists = air_list.readline()        
        air_list.close()

        airport = airport.split(",")
        name = airport[1][1:-1]+' ('+airport[4][1:-1]+'), '+airport[2][1:-1]+', '+airport[3][1:-1]
        location = latlon.LatLon(float(airport[6]),float(airport[7]))
        location = location.to_string('d%° %m%\' %S%\" %H')
        lat = location[0]
        lng = location[1]
        high = airport[8]

        table_contents = []
        result_line = result.readlines()
        total = [0,0,0,0,0,0,0,0,0]
        mean = [0,0,0,0,0,0,0,0,0]

        for line in result_line:

            data = line.split(',')

            if int(data[7][0:-1]) == 0:
                for i in range(1,7):
                    total[i] = total[i] + int(data[i])
                    data[i] = "-"

            else:
                for i in range(1,7):
                    total[i] = total[i] + int(data[i])
                    data[i] = int(data[i])*100/int(data[7][0:-1])
                    data[i] = "{:.2f}".format(data[i])

            total[7] = total[7] + int(data[7])

            table_contents.append({
                't': data[0],
                'a': data[1],
                'b': data[2],
                'c': data[3],
                'd': data[4],
                'e': data[5],
                'f': data[6],
                'g': data[7][0:-1]
                })

        if total[7] == 0:
            for i in range(1,8):
                mean[i] = "-"

        else:    
            for i in range(1,8):
                mean[i] = total[i]*100/total[7]
                mean[i] = "{:.2f}".format(mean[i])

        context = {
            'airport': name,
            'code': code,
            'mon': mon,
            'year': year,
            'lat': lat,
            'lng': lng,
            'high': high,
            'table_contents': table_contents,
            'total1': total[1],
            'total2': total[2],
            'total3': total[3],
            'total4': total[4],
            'total5': total[5],
            'total6': total[6],
            'total7': total[7],
            'mean1': mean[1],
            'mean2': mean[2],
            'mean3': mean[3],
            'mean4': mean[4],
            'mean5': mean[5],
            'mean6': mean[6],
            'mean7': mean[7]
            }

        template.render(context)
        template.save(self.Climate_path_file+'/'+'CLIMATOLOGICAL_A_'+code+'_'+year+'_'+monb+'.docx')
        result.close()        
  
    # Frequencies (percent) of Visibility below specified values (in metres) at specified times
    def model_b(self):

        os.makedirs(self.Climate_path_file, exist_ok=True)
        a = "0000 0030 0100 0130 0200 0230 0300 0330 0400 0430 0500 0530 0600 0630 0700 0730 0800 0830 0900 0930 1000 1030 1100 1130 1200 1230 1300 1330 1400 1430 1500 1530 1600 1630 1700 1730 1800 1830 1900 1930 2000 2030 2100 2130 2200 2230 2300 2330"
        aa = a.split()
        outputfileB = open(self.CSV_path_file+'/'+self.filename_in+"_model-B-result.csv", "w")

        for i in aa:
            inputfileB = open(self.CSV_path_file+'/'+self.filename_in+"_model-B.csv", "r")
            linesBB = inputfileB.readline()
            t1=t2=t3=t4=t5=t6=t7=t8=t9=total=0
            dataBB = linesBB.split(',')
            while linesBB:
                if dataBB[0] == i :
                    if int(dataBB[1]) < 150:
                        t1 = t1 + 1
                    elif 150 <= int(dataBB[1]) < 350:
                        t2 = t2 + 1
                    elif 350 <= int(dataBB[1]) < 600:
                        t3 = t3 + 1
                    elif 600 <= int(dataBB[1]) < 800:
                        t4 = t4 + 1
                    elif 800 <= int(dataBB[1]) < 1500:
                        t5 = t5 + 1
                    elif 1500 <= int(dataBB[1]) < 3000:
                        t6 = t6 + 1
                    elif 3000 <= int(dataBB[1]) < 5000:
                        t7 = t7 + 1
                    elif 5000 <= int(dataBB[1]) < 8000:
                        t8 = t8 + 1

                    else:
                        total = total+1  
                    linesBB = inputfileB.readline()
                    dataBB = linesBB.split(',')
                else:
                    linesBB = inputfileB.readline()
                    dataBB = linesBB.split(',')  

            if t1+t2+t3+t4+t5+t6+t7+t8+total != 0:           
                outputfileB.writelines("{},{},{},{},{},{},{},{},{},{}\n".format(i,t1,t1+t2,t1+t2+t3,t1+t2+t3+t4,t1+t2+t3+t4+t5,t1+t2+t3+t4+t5+t6,t1+t2+t3+t4+t5+t6+t7,t1+t2+t3+t4+t5+t6+t7+t8,t1+t2+t3+t4+t5+t6+t7+t8+total))
            inputfileB.close()
        outputfileB.close()
                

        #Import template document
        path_script = os.path.dirname(os.path.realpath(__file__))
        template = DocxTemplate(os.path.join(path_script, 'temp_model_b.docx'))
        result = open(self.CSV_path_file+'/'+ self.filename_in+"_model-B-result.csv", "r")

        code = self.code

        #TIME
        monb = self.mon
        mon = datetime.datetime.strptime(monb, "%m").strftime("%B")
        year = self.year

        #Import list airport
        air_list = open(os.path.join(path_script, 'airports.dat'), 'r')

        #find data of aiport
        lists = air_list.readline()
        while lists:
            if code in lists:
                airport = lists[0:-1]
                break
            else:
                lists = air_list.readline()        
        air_list.close()

        airport = airport.split(",")
        name = airport[1][1:-1]+' ('+airport[4][1:-1]+'), '+airport[2][1:-1]+', '+airport[3][1:-1]
        location = latlon.LatLon(float(airport[6]),float(airport[7]))
        location = location.to_string('d%° %m%\' %S%\" %H')
        lat = location[0]
        lng = location[1]
        high = airport[8]

        table_contents = []
        result_line = result.readlines()
        total = [0,0,0,0,0,0,0,0,0,0,0]
        mean = [0,0,0,0,0,0,0,0,0,0,0]

        for line in result_line:

            data = line.split(',')

            if  int(data[9][0:-1]) == 0:
                    total[i] = total[i] + int(data[i])
                    data[i] = '-'
            else:    
                for i in range(1,9):
                    total[i] = total[i] + int(data[i])
                    data[i] = int(data[i])*100/int(data[9][0:-1])
                    data[i] = "{:.2f}".format(data[i])

            total[9] = total[9] + int(data[9][0:-1])

            table_contents.append({
                't': data[0],
                'a': data[1],
                'b': data[2],
                'c': data[3],
                'd': data[4],
                'e': data[5],
                'f': data[6],
                'g': data[7],
                'h': data[8],
                'i': data[9][0:-1]
                })

        if total[9] == 0:
            for i in range(1,10):
                mean[i] = "-"

        else:
            for i in range(1,10):
                mean[i] = total[i]*100/total[9]
                mean[i] = "{:.2f}".format(mean[i])

        context = {
            'airport': name,
            'code': code,
            'mon': mon,
            'year': year,
            'lat': lat,
            'lng': lng,
            'high': high,
            'table_contents': table_contents,
            'total1': total[1],
            'total2': total[2],
            'total3': total[3],
            'total4': total[4],
            'total5': total[5],
            'total6': total[6],
            'total7': total[7],
            'total8': total[8],
            'total9': total[9],
            'mean1': mean[1],
            'mean2': mean[2],
            'mean3': mean[3],
            'mean4': mean[4],
            'mean5': mean[5],
            'mean6': mean[6],
            'mean7': mean[7],
            'mean8': mean[8],
            'mean9': mean[9]
            }

        template.render(context)
        template.save(self.Climate_path_file+'/'+'CLIMATOLOGICAL_B_'+code+'_'+year+'_'+monb+'.docx')
        result.close()                

    # Frequencies (percent) of the Height of Base of the Cloud Layer of BKN or OVC specified values (in feets) at specified times
    def model_c(self):

        os.makedirs(self.Climate_path_file, exist_ok=True)
        a = "0000 0030 0100 0130 0200 0230 0300 0330 0400 0430 0500 0530 0600 0630 0700 0730 0800 0830 0900 0930 1000 1030 1100 1130 1200 1230 1300 1330 1400 1430 1500 1530 1600 1630 1700 1730 1800 1830 1900 1930 2000 2030 2100 2130 2200 2230 2300 2330"
        aa = a.split()
        outputfileC = open(self.CSV_path_file+'/'+self.filename_in+"_model-C-result.csv", "w")

        for i in aa:
            inputfileC = open(self.CSV_path_file+'/'+self.filename_in+"_model-C.csv", "r")
            linesCC = inputfileC.readline()
            linesCCC = linesCC.replace("non","0")
            t1=t2=t3=t4=t5=t6=t7=t8=t9=total=0
            dataCC = linesCCC.split(',')
            while linesCC:
                if dataCC[0] == i :
                    if 0 < int(dataCC[1]) < 2:
                        t1 = t1 + 1
                    elif 2 <= int(dataCC[1]) < 5:
                        t2 = t2 + 1
                    elif 5 <= int(dataCC[1]) < 10:
                        t3 = t3 + 1
                    elif 10 <= int(dataCC[1]) < 15:
                        t4 = t4 + 1
                    elif 15 <= int(dataCC[1]) < 20:
                        t5 = t5 + 1
                    elif 20 <= int(dataCC[1]) < 50:
                        t6 = t6 + 1
                    elif int(dataCC[1]) >= 50:
                        t7 = t7 + 1
                    else:
                        total = total + 1  
                    linesCC = inputfileC.readline()
                    linesCCC = linesCC.replace("non","0")
                    dataCC = linesCCC.split(',')
   
                else:
                    linesCC = inputfileC.readline()
                    linesCCC = linesCC.replace("non","0")
                    dataCC = linesCCC.split(',')      

            if t1+t2+t3+t4+t5+t6+t7+total != 0:        
                outputfileC.writelines("{},{},{},{},{},{},{},{},{}\n".format(i,t1,t1+t2,t1+t2+t3,t1+t2+t3+t4,t1+t2+t3+t4+t5,t1+t2+t3+t4+t5+t6,t1+t2+t3+t4+t5+t6+t7,t1+t2+t3+t4+t5+t6+t7+total))
            inputfileC.close()
        outputfileC.close()
                

        #Import template document
        path_script = os.path.dirname(os.path.realpath(__file__))
        template = DocxTemplate(os.path.join(path_script, 'temp_model_c.docx'))
        result = open(self.CSV_path_file+'/'+self.filename_in+"_model-C-result.csv", "r")

        code = self.code

        #TIME
        monb = self.mon
        mon = datetime.datetime.strptime(monb, "%m").strftime("%B")
        year = self.year

        #Import list airport
        air_list = open(os.path.join(path_script, 'airports.dat'), 'r')


        #find data of aiport
        lists = air_list.readline()
        while lists:
            if code in lists:
                airport = lists[0:-1]
                break
            else:
                lists = air_list.readline()        
        air_list.close()

        airport = airport.split(",")
        name = airport[1][1:-1]+' ('+airport[4][1:-1]+'), '+airport[2][1:-1]+', '+airport[3][1:-1]
        location = latlon.LatLon(float(airport[6]),float(airport[7]))
        location = location.to_string('d%° %m%\' %S%\" %H')
        lat = location[0]
        lng = location[1]
        high = airport[8]

        table_contents = []
        result_line = result.readlines()
        total = [0,0,0,0,0,0,0,0,0,0,0]
        mean = [0,0,0,0,0,0,0,0,0,0,0]

        for line in result_line:

            data = line.split(',')

            if int(data[8][0:-1]) == 0:
                for i in range(1,8):
                    total[i] = total[i] + int(data[i])
                    data[i] = "-"             
            else:
                for i in range(1,8):
                    total[i] = total[i] + int(data[i])
                    data[i] = int(data[i])*100/int(data[8][0:-1])
                    data[i] = "{:.2f}".format(data[i])

            total[8] = total[8] + int(data[8][0:-1])

            table_contents.append({
                't': data[0],
                'a': data[1],
                'b': data[2],
                'c': data[3],
                'd': data[4],
                'e': data[5],
                'f': data[6],
                'g': data[7],
                'h': data[8][0:-1]
                })

        if total[8] == 0:
            for i in range(1,9):
                mean[i] = "-"
        else:
            for i in range(1,9):
                mean[i] = total[i]*100/total[8]
                mean[i] = "{:.2f}".format(mean[i])

        context = {
            'airport': name,
            'code': code,
            'mon': mon,
            'year': year,
            'lat': lat,
            'lng': lng,
            'high': high,
            'table_contents': table_contents,
            'total1': total[1],
            'total2': total[2],
            'total3': total[3],
            'total4': total[4],
            'total5': total[5],
            'total6': total[6],
            'total7': total[7],
            'total8': total[8],
            'mean1': mean[1],
            'mean2': mean[2],
            'mean3': mean[3],
            'mean4': mean[4],
            'mean5': mean[5],
            'mean6': mean[6],
            'mean7': mean[7],
            'mean8': mean[8]
            }

        template.render(context)
        template.save(self.Climate_path_file+'/'+'CLIMATOLOGICAL_C_'+code+'_'+year+'_'+monb+'.docx')
        result.close()                

    # Frequencies (percent) of the Occurrence of Concurrent Wind Direction (in 30° Sections) and Speed (in knots) within specified ranges
    def model_d(self):

        os.makedirs(self.Climate_path_file, exist_ok=True)
        outputfileDD = open(self.CSV_path_file+'/'+self.filename_in+"_model-D-result.csv", "w")
        b =("VR","35 36 01","02 03 04","05 06 07","08 09 10","11 12 13","14 15 16","17 18 19","20 21 22","23 24 25","26 27 28","29 30 31","32 33 34")
        c = 0
        fd = open(self.CSV_path_file+'/'+self.filename_in+"_model-D.csv", "r")
        ld = fd.readline()
        dd = ld.split(',')
        while ld:

            if dd[1][0:-1] == "00":
                c = c + 1
                ld = fd.readline()
                dd = ld.split(',')
            else:
                ld = fd.readline()
                dd = ld.split(',')
        outputfileDD.writelines("{},0,0,0,0,0,0,0,0,0,0,0\n".format(c))

        for i in b:
            inputfileDD = open(self.CSV_path_file+'/'+self.filename_in+"_model-D.csv", "r")
            linesDD = inputfileDD.readline() 
            dataDD = linesDD.split(',')
            t1=t2=t3=t4=t5=t6=t7=t8=t9=t10=t11=t12=t13=0
            while linesDD:
                if dataDD[0] in i :
                    if dataDD[1] == "-":
                        t13 = t13 + 1
                    elif 0 <= int(dataDD[1]) < 1:
                        t1 = t1 + 1
                    elif 1 <= int(dataDD[1]) < 5:
                        t2 = t2 + 1
                    elif 5 <= int(dataDD[1]) < 10:
                        t3 = t3 + 1
                    elif 10 <= int(dataDD[1]) < 15:
                        t4 = t4 + 1
                    elif 15 <= int(dataDD[1]) < 20:
                        t5 = t5 + 1
                    elif 20 <= int(dataDD[1]) < 25:
                        t6 = t6 + 1
                    elif 25 <= int(dataDD[1]) < 30:
                        t7 = t7 + 1
                    elif 30 <= int(dataDD[1]) < 35:
                        t8 = t8 + 1
                    elif 35 <= int(dataDD[1]) < 40:
                        t9 = t9 + 1
                    elif 40 <= int(dataDD[1]) < 45:
                        t10 = t10 + 1
                    elif 45 <= int(dataDD[1]) < 50:
                        t11 = t11 + 1
                    elif int(dataDD[1]) >= 50:
                        t12 = t12 + 1
                        
            
                    linesDD = inputfileDD.readline()
                    dataDD = linesDD.split(',')
   
                else:
                    linesDD = inputfileDD.readline()
                    dataDD = linesDD.split(',')            
        
            outputfileDD.writelines("{},{},{},{},{},{},{},{},{},{},{},{}\n".format(t1,t2,t3,t4,t5,t6,t7,t8,t9,t10,t11,t12))
            inputfileDD.close()
        outputfileDD.close()


        #Import template document
        path_script = os.path.dirname(os.path.realpath(__file__))
        template = DocxTemplate(os.path.join(path_script, 'temp_model_d.docx'))
        result = open(self.CSV_path_file+'/'+self.filename_in+"_model-D-result.csv", "r")
        numdata = open(self.CSV_path_file+'/'+self.filename_in+"_model-D-result.csv", "r")

        code = self.code

        #TIME
        monb = self.mon
        mon = datetime.datetime.strptime(monb, "%m").strftime("%B")
        year = self.year

        #Import list airport
        air_list = open(os.path.join(path_script, 'airports.dat'), 'r')


        #find data of aiport
        lists = air_list.readline()
        while lists:
            if code in lists:
                airport = lists[0:-1]
                break
            else:
                lists = air_list.readline()        
        air_list.close()

        airport = airport.split(",")
        name = airport[1][1:-1]+' ('+airport[4][1:-1]+'), '+airport[2][1:-1]+', '+airport[3][1:-1]
        location = latlon.LatLon(float(airport[6]),float(airport[7]))
        location = location.to_string('d%° %m%\' %S%\" %H')
        lat = location[0]
        lng = location[1]
        high = airport[8]
        table_contents = []

        total = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
        mean = [0,0,0,0,0,0,0,0,0.0,0,0,0,0,0]
        tot = 0
        wind_direct = 'CALM VARIABLE 35-36-01 02-03-04 05-06-07 08-09-10 11-12-13 14-15-16 17-18-19 20-21-22 23-24-25 26-27-28 29-30-31 32-33-34'
        wind_direct = wind_direct.split(' ')

        for line in wind_direct:
            totalnum = numdata.readline()
            totalnum = totalnum.split(',')
            totalnum = [int(x) for x in totalnum]
            tot = tot + sum(totalnum)


        for line in wind_direct:
            data = result.readline()
            data = data.split(',')
            data = [int(x) for x in data]
            data.append(sum(data[0:12]))

            if int(tot) == 0:
                for i in range(0,13):
                    total[i] = total[i] + int(data[i])
                    data[i] = "-"

            else:    
                for i in range(0,13):
                    total[i] = total[i] + int(data[i])
                    data[i] = int(data[i])*100/int(tot)
                    data[i] = "{:.2f}".format(data[i])
    
            table_contents.append({
                't': line,
                'a': data[0],
                'b': data[1],
                'c': data[2],
                'd': data[3],
                'e': data[4],
                'f': data[5],
                'g': data[6],
                'h': data[7],
                'i': data[8],
                'j': data[9],
                'k': data[10],
                'm': data[11],
                'n': data[12]
                })

        if int(tot) == 0:
            for i in range(0,13):
                mean[i] = "-"
        else:    
            for i in range(0,13):
                mean[i] = total[i]*100/int(tot)
                mean[i] = "{:.2f}".format(mean[i])

        context = {
            'airport': name,
            'code': code,
            'mon': mon,
            'year': year,
            'lat': lat,
            'lng': lng,
            'high': high,
            'table_contents': table_contents,
            'mean1': mean[0],
            'mean2': mean[1],
            'mean3': mean[2],
            'mean4': mean[3],
            'mean5': mean[4],
            'mean6': mean[5],
            'mean7': mean[6],
            'mean8': mean[7],
            'mean9': mean[8],
            'mean10': mean[9],
            'mean11': mean[10],
            'mean12': mean[11],
            'mean13': mean[12]
            }

        template.render(context)
        template.save(self.Climate_path_file+'/'+'CLIMATOLOGICAL_D_'+code+'_'+year+'_'+monb+'.docx')
        result.close()
        numdata.close()        

    # Frequencies (percent) of Surface Temperature in specified ranges of 5 degrees Celsius at specified times
    def model_e(self):

        os.makedirs(self.Climate_path_file, exist_ok=True)
        a = "0000 0030 0100 0130 0200 0230 0300 0330 0400 0430 0500 0530 0600 0630 0700 0730 0800 0830 0900 0930 1000 1030 1100 1130 1200 1230 1300 1330 1400 1430 1500 1530 1600 1630 1700 1730 1800 1830 1900 1930 2000 2030 2100 2130 2200 2230 2300 2330"
        aa = a.split()

        outputfileE = open(self.CSV_path_file+'/'+self.filename_in+"_model-E-result.csv", "w")

        for i in aa:
            inputfileE = open(self.CSV_path_file+'/'+self.filename_in+"_model-E.csv", "r")
            linesEE = inputfileE.readline()
            t1=t2=t3=t4=t5=t6=t7=t8=t9=0
            dataEE = linesEE.split(',')
            while linesEE:
                if dataEE[0] == i:
                    if int(dataEE[1]) < 10:
                        t1 = t1 + 1
                    elif 10 <= int(dataEE[1]) < 15:
                        t2 = t2 + 1
                    elif 15 <= int(dataEE[1]) < 20:
                        t3 = t3 + 1
                    elif 20 <= int(dataEE[1]) < 25:
                        t4 = t4 + 1
                    elif 25 <= int(dataEE[1]) < 30:
                        t5 = t5 + 1
                    elif 30 <= int(dataEE[1]) < 35:
                        t6 = t6 + 1
                    elif 35 <= int(dataEE[1]) < 40:
                        t7 = t7 + 1
                    elif 40 <= int(dataEE[1]) < 45:
                        t8 = t8 + 1
                    elif int(dataEE[1]) > 45:
                        t9 = t9 + 1
   
                    linesEE = inputfileE.readline()
                    dataEE = linesEE.split(',')  
                else:
                    linesEE = inputfileE.readline()
                    dataEE = linesEE.split(',')     

            if t1+t2+t3+t4+t5+t6+t7+t8+t9 != 0:          
                outputfileE.writelines("{},{},{},{},{},{},{},{},{},{},{}\n".format(i,t1,t2,t3,t4,t5,t6,t7,t8,t9,t1+t2+t3+t4+t5+t6+t7+t8+t9))
            inputfileE.close()
        outputfileE.close()


        #Import template document
        path_script = os.path.dirname(os.path.realpath(__file__))
        template = DocxTemplate(os.path.join(path_script, 'temp_model_e.docx'))
        result = open(self.CSV_path_file+'/'+self.filename_in+"_model-E-result.csv", "r")

        code = self.code

        #TIME
        monb = self.mon
        mon = datetime.datetime.strptime(monb, "%m").strftime("%B")
        year = self.year

        #Import list airport
        air_list = open(os.path.join(path_script, 'airports.dat'), 'r')


        #find data of aiport
        lists = air_list.readline()
        while lists:
            if code in lists:
                airport = lists[0:-1]
                break
            else:
                lists = air_list.readline()        
        air_list.close()

        airport = airport.split(",")
        name = airport[1][1:-1]+' ('+airport[4][1:-1]+'), '+airport[2][1:-1]+', '+airport[3][1:-1]
        location = latlon.LatLon(float(airport[6]),float(airport[7]))
        location = location.to_string('d%° %m%\' %S%\" %H')
        lat = location[0]
        lng = location[1]
        high = airport[8]
        table_contents = []
        result_line = result.readlines()
        total = [0,0,0,0,0,0,0,0,0,0,0,0]
        mean = [0,0,0,0,0,0,0,0,0,0,0,0]

        for line in result_line:

            data = line.split(',')

            if int(data[10][0:-1]) == 0:
                for i in range(1,10):
                    total[i] = total[i] + int(data[i])
                    data[i] = "-"
            else:    
                for i in range(1,10):
                    total[i] = total[i] + int(data[i])
                    data[i] = int(data[i])*100/int(data[10][0:-1])
                    data[i] = "{:.1f}".format(data[i])

            total[10] = total[10] + int(data[10][0:-1])

            table_contents.append({
                't': data[0],
                'a': data[1],
                'b': data[2],
                'c': data[3],
                'd': data[4],
                'e': data[5],
                'f': data[6],
                'g': data[7],
                'h': data[8],
                'i': data[9],
                'j': data[10][0:-1]
                })

        if total[10] == 0:
            for i in range(1,11):
                mean[i] = "-"
        else:
            for i in range(1,11):
                mean[i] = total[i]*100/total[10]
                mean[i] = "{:.2f}".format(mean[i])

        context = {
            'airport': name,
            'code': code,
            'mon': mon,
            'year': year,
            'lat': lat,
            'lng': lng,
            'high': high,
            'table_contents': table_contents,
            'total1': total[1],
            'total2': total[2],
            'total3': total[3],
            'total4': total[4],
            'total5': total[5],
            'total6': total[6],
            'total7': total[7],
            'total8': total[8],
            'total9': total[9],
            'total10': total[10],
            'mean1': mean[1],
            'mean2': mean[2],
            'mean3': mean[3],
            'mean4': mean[4],
            'mean5': mean[5],
            'mean6': mean[6],
            'mean7': mean[7],
            'mean8': mean[8],
            'mean9': mean[9],
            'mean10': mean[10]
            }

        template.render(context)
        template.save(self.Climate_path_file+'/'+'CLIMATOLOGICAL_E_'+code+'_'+year+'_'+monb+'.docx')
        result.close()

    # Occurrences of Specified Phenomena (reports)
    def model_f(self):
        
        os.makedirs(self.Climate_path_file, exist_ok=True)
        outputfileFF = open(self.CSV_path_file+'/'+self.filename_in+"_model-F-result.csv", "w")
        rawmetar = open(self.filename_in, encoding="utf8", errors='ignore')
        datametar = rawmetar.readlines()[-1:]

        if re.findall("\d{6}Z", datametar[0]):
            DTZF = re.findall("\d{6}Z", datametar[0])
            DDF = DTZF[0][0:2]
            for day in range(1,int(DDF)+1):
                inputfileFF = open(self.filename_in, "r")
                inputfileFFF = open(self.CSV_path_file+'/'+self.filename_in+"_model-F.csv", "r")
                next(inputfileFF)
                next(inputfileFFF)
                linesFF = inputfileFF.readline()
                linesFFF = inputfileFFF.readline()
                linesFFFF = linesFFF.replace("No","0")
                DATAFFF = linesFFFF.split(',')
                t10=t15=t20=t25=t30=t35=t40=t45=t50=tot=0
    
                while linesFF:

                    if re.findall("\d{6}Z", linesFF):
                        DTZF = re.findall("\d{6}Z", linesFF)
                        DDF = DTZF[0][0:2]
                    else:
                        DDF = '0'

                    if re.findall('\w+\W+\w+\W+\w+\W+\w+\W+\w+\W+\w+\W+\w+\W+Q\d{4}', linesFF):
                        CUTDATAF = re.findall('\w+\W+\w+\W+\w+\W+\w+\W+\w+\W+\w+\W+\w+\W+Q\d{4}', linesFF)
                    else:
                        CUTDATAF = ['0']

                    if int(DDF) == day and re.findall("DZ", CUTDATAF[0]):
                        t10 = t10+1
                    elif int(DDF) == day and re.findall("\s\D{1}SHRA|\sSHRA", CUTDATAF[0]):
                        t20 = t20+1
                    elif int(DDF) == day and re.findall("\s\D{1}TSRA|\sTSRA", CUTDATAF[0]):
                        t30 = t30+1
                    elif int(DDF) == day and re.findall("\s\D{1}RA|\sRA", CUTDATAF[0]):
                        t15 = t15+1
                    elif int(DDF) == day and re.findall("\s\D{1}TS|\sTS", CUTDATAF[0]):
                        t25 = t25+1
                    elif int(DDF) == day and re.findall("FG", CUTDATAF[0]):
                        t35 = t35+1
                    elif int(DDF) == day and re.findall("BR", CUTDATAF[0]):
                        t40 = t40+1
                    elif int(DDF) == day and re.findall("HZ", CUTDATAF[0]):
                        t45 = t45+1
                    linesFF = inputfileFF.readline()    

                while linesFFF:
                    if int(DATAFFF[0]) == day and int(DATAFFF[1]) > 27:
                        t50 = t50 + 1
                    linesFFF = inputfileFFF.readline()
                    linesFFFF = linesFFF.replace("No","0")
                    DATAFFF = linesFFFF.split(',')

                inputfileFF.close()
                inputfileFFF.close()
    
                outputfileFF.writelines("{},{},{},{},{},{},{},{},{},{}\n".format(day,t10,t15,t20,t25,t30,t35,t40,t45,t50))


            outputfileFF.close()            
        else:
            outputfileFF.close()



        #Import template document
        path_script = os.path.dirname(os.path.realpath(__file__))
        template = DocxTemplate(os.path.join(path_script, 'temp_model_f.docx'))
        result = open(self.CSV_path_file+'/'+self.filename_in+"_model-F-result.csv", "r")

        code = self.code

        #TIME
        monb = self.mon
        mon = datetime.datetime.strptime(monb, "%m").strftime("%B")
        year = self.year

        #Import list airport
        air_list = open(os.path.join(path_script, 'airports.dat'), 'r')


        #find data of aiport
        lists = air_list.readline()
        while lists:
            if code in lists:
                airport = lists[0:-1]
                break
            else:
                lists = air_list.readline()        
        air_list.close()

        airport = airport.split(",")
        name = airport[1][1:-1]+' ('+airport[4][1:-1]+'), '+airport[2][1:-1]+', '+airport[3][1:-1]
        location = latlon.LatLon(float(airport[6]),float(airport[7]))
        location = location.to_string('d%° %m%\' %S%\" %H')
        lat = location[0]
        lng = location[1]
        high = airport[8]

        table_contents = []
        result_line = result.readlines()
        total = [0,0,0,0,0,0,0,0,0,0,0,0]
        mean = [0,0,0,0,0,0,0,0,0,0,0,0]

        for line in result_line:

            data = line.split(',')

            for i in range(1,10):
                total[i] = total[i] + int(data[i])



            table_contents.append({
                't': data[0],
                'a': data[1],
                'b': data[2],
                'c': data[3],
                'd': data[4],
                'e': data[5],
                'f': data[6],
                'g': data[7],
                'h': data[8],
                'i': data[9][0:-1]
                })


        context = {
            'airport': name,
            'code': code,
            'mon': mon,
            'year': year,
            'lat': lat,
            'lng': lng,
            'high': high,
            'table_contents': table_contents,
            'total1': total[1],
            'total2': total[2],
            'total3': total[3],
            'total4': total[4],
            'total5': total[5],
            'total6': total[6],
            'total7': total[7],
            'total8': total[8],
            'total9': total[9]
            }

        template.render(context)
        template.save(self.Climate_path_file+'/'+'CLIMATOLOGICAL_F_'+code+'_'+year+'_'+monb+'.docx')
        result.close()

    # Frequencies (percent) of the Occurrence of RVR (in meters) and/or the Height of the Base of 
    # the Lowest Cloud Layer (in feet) of BKN or OVC Extent below specified values at specified times
    def model_g(self):

        os.makedirs(self.Climate_path_file, exist_ok=True)
        a = "0000 0030 0100 0130 0200 0230 0300 0330 0400 0430 0500 0530 0600 0630 0700 0730 0800 0830 0900 0930 1000 1030 1100 1130 1200 1230 1300 1330 1400 1430 1500 1530 1600 1630 1700 1730 1800 1830 1900 1930 2000 2030 2100 2130 2200 2230 2300 2330"
        aa = a.split()

        outputfileG = open(self.CSV_path_file+'/'+self.filename_in+"_model-G-result.csv", "w")

        for i in aa:
            inputfileG = open(self.CSV_path_file+'/'+self.filename_in+"_model-G.csv", "r")
            linesGG = inputfileG.readline()
            linesGGG = linesGG.replace("No","0").replace("non","0")
            t1=t2=t3=t4=t5=t6=tot=0
            dataGG = linesGGG.split(',')
            while linesGGG:
                if dataGG[0] == i :
                    if 0 < int(dataGG[1]) < 100 or 0 < int(dataGG[2]) <= 2:
                        t1 = t1 + 1
                    elif 0 < int(dataGG[1]) < 150 or 0 < int(dataGG[2]) <= 4.5:
                        t2 = t2 + 1
                    elif 0 < int(dataGG[1]) < 350 or 0 < int(dataGG[2]) <= 8:
                        t3 = t3 + 1
                    elif 0 < int(dataGG[1]) < 550 or 0 <int(dataGG[2]) <= 15:
                        t4 = t4 + 1
                    elif 0 < int(dataGG[1]) < 800 or 0 <int(dataGG[2]) <= 30:
                        t5 = t5 + 1
                    elif 0 < int(dataGG[1]) < 1500 or 0 <int(dataGG[2]) <= 50:
                        t6 = t6 + 1
                    else:
                        tot = tot+1  

                    linesGG = inputfileG.readline()
                    linesGGG = linesGG.replace("No","0").replace("non","0")
                    dataGG = linesGGG.split(',')
 
                else:
                    linesGG = inputfileG.readline()
                    linesGGG = linesGG.replace("No","0").replace("non","0")
                    dataGG = linesGGG.split(',') 
                         
            if t1+t2+t3+t4+t5+t6+tot != 0: 
                outputfileG.writelines("{},{},{},{},{},{},{},{}\n".format(i,t1,t1+t2,t1+t2+t3,t1+t2+t3+t4,t1+t2+t3+t4+t5,t1+t2+t3+t4+t5+t6,t1+t2+t3+t4+t5+t6+tot))    
            inputfileG.close()
        outputfileG.close()


        #Import template document
        path_script = os.path.dirname(os.path.realpath(__file__))
        template = DocxTemplate(os.path.join(path_script, 'temp_model_g.docx'))
        result = open(self.CSV_path_file+'/'+self.filename_in+"_model-G-result.csv", "r")
        test = open(self.CSV_path_file+'/'+self.filename_in+"_model-G-result.csv", "r")
        numline = open(self.CSV_path_file+'/'+self.filename_in+"_model-G-result.csv", "r")

        code = self.code

        #TIME
        monb = self.mon
        mon = datetime.datetime.strptime(monb, "%m").strftime("%B")
        year = self.year

        #Import list airport
        air_list = open(os.path.join(path_script, 'airports.dat'), 'r')


        #find data of aiport
        lists = air_list.readline()
        while lists:
            if code in lists:
                airport = lists[0:-1]
                break
            else:
                lists = air_list.readline()        
        air_list.close()

        airport = airport.split(",")
        name = airport[1][1:-1]+' ('+airport[4][1:-1]+'), '+airport[2][1:-1]+', '+airport[3][1:-1]
        location = latlon.LatLon(float(airport[6]),float(airport[7]))
        location = location.to_string('d%° %m%\' %S%\" %H')
        lat = location[0]
        lng = location[1]
        high = airport[8]


        table_contents = []
        total = [0,0,0,0,0,0,0,0,0,0,0,0]
        mean = [0,0,0,0,0,0,0,0,0,0,0,0]
        data =['0',0,0,0,0,0,0,0,0,0,0,0]
        time = "00-01 01-02 02-03 03-04 04-05 05-06 06-07 07-08 08-09 09-10 10-11 11-12 12-13 13-14 14-15 15-16 16-17 17-18 18-19 19-20 20-21 21-22 22-23 23-24"
        time = time.split(' ')
        test1 = test.readline()
        test1 = test.readline()
        numline = numline.readlines()

        if test1[2:4] == '30':

            for line in time:

                data1 = result.readline()
                data1 = data1.split(',')
                data2 = result.readline()
                data2 = data2.split(',')

                if data1 == [''] :
                    for i in range(0,8):
                        data[i] = 0
                else:
                    for i in range(0,8):
                        data[i] = int(data1[i]) + int(data2[i])

                if int(data[7]) == 0:
                    for i in range(1,7):
                        total[i] = total[i] + int(data[i])
                        data[i] = "-"
                else:
                    for i in range(1,7):
                        total[i] = total[i] + int(data[i])
                        data[i] = int(data[i])*100/int(data[7])
                        data[i] = "{:.2f}".format(data[i])

                total[7] = total[7] + int(data[7])

                table_contents.append({
                    't': line,
                    'a': data[1],
                    'b': data[2],
                    'c': data[3],
                    'd': data[4],
                    'e': data[5],
                    'f': data[6],
                    'g': data[7]
                    })
                
            if total[7] == 0:
                for i in range(1,8):
                    mean[i] = "-"
            else:
                for i in range(1,8):
                    mean[i] = total[i]*100/total[7]
                    mean[i] = "{:.2f}".format(mean[i])

        elif test1[2:4] == '00':
            data1 = result.readline()
            data2 = data1.split(',')

            for line in numline:

                if data2 == [''] :
                    data[0] = '-'
                    for i in range(1,8):
                        data[i] = 0
                else:
                    data[0] = data2[0]
                    for i in range(1,8):
                        data[i] = int(data2[i])

                if int(data[7]) == 0:
                    for i in range(1,7):
                        total[i] = total[i] + int(data[i])
                        data[i] = "-"
                else:
                    for i in range(1,7):
                        total[i] = total[i] + int(data[i])
                        data[i] = int(data[i])*100/int(data[7])
                        data[i] = "{:.2f}".format(data[i])

                total[7] = total[7] + int(data[7])

                table_contents.append({
                    't': data[0],
                    'a': data[1],
                    'b': data[2],
                    'c': data[3],
                    'd': data[4],
                    'e': data[5],
                    'f': data[6],
                    'g': data[7]
                    })

                data1 = result.readline()
                data2 = data1.split(',')
  
            if total[7] == 0:
                for i in range(1,8):
                    mean[i] = "-"
            else:
                for i in range(1,8):
                    mean[i] = total[i]*100/total[7]
                    mean[i] = "{:.2f}".format(mean[i])

        context = {
            'airport': name,
            'code': code,
            'mon': mon,
            'year': year,
            'lat': lat,
            'lng': lng,
            'high': high,
            'table_contents': table_contents,
            'total1': total[1],
            'total2': total[2],
            'total3': total[3],
            'total4': total[4],
            'total5': total[5],
            'total6': total[6],
            'total7': total[7],
            'total8': total[8],
            'total9': total[9],
            'total10': total[10],
            'mean1': mean[1],
            'mean2': mean[2],
            'mean3': mean[3],
            'mean4': mean[4],
            'mean5': mean[5],
            'mean6': mean[6],
            'mean7': mean[7],
            'mean8': mean[8],
            'mean9': mean[9],
            'mean10': mean[10]
            }

        template.render(context)
        template.save(self.Climate_path_file+'/'+'CLIMATOLOGICAL_G_'+code+'_'+year+'_'+monb+'.docx')
        result.close()
        test.close()


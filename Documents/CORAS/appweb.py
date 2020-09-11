from datetime import datetime, timedelta
from dateutil import tz
import xlrd
import xlsxwriter
import openpyxl

path = input("Filename: ")
# path = "orders-2020-07-27-07-48-47.xlsx" #original file name
nweb = "Web_" +path
napp = "App_" +path
inputWorkbook = xlrd.open_workbook(path)
inputWorksheet = inputWorkbook.sheet_by_index(0)
row = inputWorksheet.nrows # no of rows
col = inputWorksheet.ncols # no of col
val = inputWorksheet.cell_value
i,j=1,1
app = xlsxwriter.Workbook(napp) #add date to name
web = xlsxwriter.Workbook(nweb)
wbold = web.add_format({'bold': True})
abold = app.add_format({'bold': True})
webSheet = web.add_worksheet()
appSheet = app.add_worksheet()


#app booking {Start DateTime, End DateTime} (x,12), (x,13)
#Start DateTime => float
#End DateTime => String
#web booking {Activity Start Date, Activitiy Start Time, Activity End Time} (x,25), (x,26), (x,27)


def con(ordinal, _epoch0=datetime(1899, 12, 31)):
    if ordinal >= 60:
        ordinal -= 1  # handle leap year bug
    return (_epoch0 + timedelta(days=ordinal)).replace(microsecond=0)


def InsertA (no): #app (order date)
    if no !=0:
        x = str(con(val(no,24))+ timedelta(hours=8))
        y = str(datetime.strptime(val(no,25), '%Y-%m-%d %H:%M:%S.%fZ') + timedelta(hours=8))
        dat = x[0:10]
        stime = x[10:]
        etime= y[10:]

    for c in range (26):
        appSheet.set_column(no,c,28)

        if(c <=8):
            if (no ==0 and c==8):
                appSheet.write(0,c, "Additional Guest E-mail(s)",abold)
            elif(no ==0):
                appSheet.write(0,c,val(0,c),abold)
            elif(c==8):
                appSheet.write(i,c,val(no,8))

            elif (no != 0 and c==0):
                da = str(con(val(no,0))+timedelta(hours=8))
                appSheet.write(i,c,da)
            else: 
                appSheet.write(i,c,val(no,c)) #val(no+4,c)

        elif(c<22):
            if(no ==0 and c ==18):
                appSheet.write(0,18,val(0,20),abold)
            elif(no==0 and c>18):
                appSheet.write(0,c,val(0,c+2),abold)
            elif(no==0):
                appSheet.write(0,c, val(0,c+1),abold)
            elif(c==20):
                appSheet.write(i,c,val(no,c+2))
            elif(c==21):
                appSheet.write(i,c,val(no,c))
            else:
                appSheet.write(i,c,val(no,c+1))
            

        elif(c>=22):
            if(no==0):
                appSheet.write(0,c,val(0,c+4), abold)

            elif(c==22):
                appSheet.write(i,c,dat)
            elif(c==23):
                appSheet.write(i,c,stime)
            elif(c==24):
                appSheet.write(i,c,etime)
            else:
                appSheet.write(i,c,val(no,c+4))
                

        
            
def InsertW (no): 
    for c in range (26):
        webSheet.set_column(no,c,28)

        if(c <=8):
            if (no ==0 and c==8):
                webSheet.write(0,c, "Additional Guest E-mail(s)",wbold)
            elif(no ==0):
                webSheet.write(0,c,val(0,c),wbold)
            elif(c==8):
                webSheet.write(j,c,val(no,9))

            elif (no != 0 and c==0):
                date = str(con(val(no,0))+timedelta(hours=8))
                webSheet.write(j,c,date)
            else: 
                webSheet.write(j,c,val(no,c)) #val(no+4,c)

        elif(c<22):
            if(no ==0 and c ==18):
                webSheet.write(0,18,val(0,20),wbold)
            elif(no==0 and c>18):
                webSheet.write(0,c,val(0,c+2),wbold)
            elif(c>=18):
                webSheet.write(j,c,val(no,c+2))
            elif(no==0):
                webSheet.write(0,c, val(0,c+1),wbold)
            else:
                 webSheet.write(j,c,val(no,c+1))
            

        elif(c>=22):
            if(no==0):
                webSheet.write(0,c,val(0,c+4), wbold)

            elif(c==23 or c==24):
                y = val(no,c+4)
                x = y[-2:].upper()
                xy = y[0:len(y)-2]+x
                in_time = datetime.strptime(xy, "%I:%M %p")+ timedelta(hours=8)
                time = datetime.strftime(in_time, "%H:%M")
                webSheet.write(j,c,time)

            else:
                webSheet.write(j,c,val(no,c+4))
    


InsertA(0)
InsertW(0)
# print(con(val(1,0))+timedelta(hours=8))

for r in range (1,row):
    if (val(r,24) != ""): #app
        print("app")
        InsertA(r)
        i+=1
             
    else:
        print("web")
        InsertW(r)
        j+=1
    

# print(val(0,8))







# for i in range (row): 
#     if (i== 0) :
#         for j in range (col):
#             appSheet.set_column(j,j,25)
#             appSheet.write(i,j, val(0,j),abold)
#             webSheet.set_column(j,j,25)
#             webSheet.write(i,j, val(0,j),wbold)

#     if( val(i,12) != None): #its app
#         for j in range (col):
#             outSheet.write()




# if (val(0,0) !=  None):
#     print(val(0,13))
#     print((val(1,13)))
#     e = datetime.fromisoformat(val(1,13)) # for string format
#     print(e)
#     print(val(0,12))
#     print(val(1,12))
#     print(con(val(1,12)))

    # excel_date = con(val(1,12))

    
    
# else:
#     print("empty")

app.close()
web.close()

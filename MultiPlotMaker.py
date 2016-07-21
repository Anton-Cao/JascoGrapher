import xlsxwriter #library to create and write to excel files
from tkinter.filedialog import askopenfilename #library to choose file

def strip(path): #method that removes preceding C://... and trailing .txt extention of the path
    lastSlash = 0 
    for i,c in enumerate(path): #finds last '/'
        if c == '/':
            lastSlash = i
    return path[lastSlash+1:-4]

def execute():
    print('What do you want the excel file to be called?') #name output file, .xlsx extension not needed
    filename = input()
    print('What do you want to name the chart? ("same" if same name as excel file)')
    chartname = input()
    if chartname == 'same':
        chartname = filename
    filename += '.xlsx'

    row = 801 #for ease of selecting data

    miny = 2 #store mininum y-value for formatting graph
    maxy = 0 #store maximum y-value for formatting graph

    wb = xlsxwriter.Workbook(filename) #create new excel workbook
    ws = wb.add_worksheet() #create new worksheet in that workbook
    chart = wb.add_chart({'type': 'scatter','subtype': 'smooth'}) #add scatterplot

    letters = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'] #for ease of reference
    numFile = 0 #number of input files
    cont = 'Y' #whether more files need to be read ('Y' - yes, 'N' - no)

    while cont == 'Y':
        row = 0
        numFile += 1
        dataname = askopenfilename() #get .txt file
        data = open(dataname,'r') #open data file
        for line in data: #goes through each line in .txt file
            curLine = line.split() #splits the line into array of words
            try: #if both words are numbers, this means this is actual data and not other info
                x=float(curLine[0]) 
                y=float(curLine[1])
                #write data to desired cells
                if numFile == 1:
                    ws.write(row,0,x) 
                ws.write(row,numFile,y)
                row+=1
                #update min and max values as necessary
                if y > maxy: 
                    maxy = y
                if y < miny:
                    miny = y
            except: #otherwise, do nothing
                1+1 #just a placeholder

        print('Data name?')
        curname = input()
        chart.add_series({
            #'name': '%s'%strip(dataname),
            'name': '%s'%curname,
            'categories': '=Sheet1!$A$1:$A$%d'%row,
            'values': '=Sheet1!$%s$1:$%s$%d'%(letters[numFile],letters[numFile],row),
        })
        
        print('More data? (Y/N)')
        cont = input()


    #set some info for chart
    chart.set_title({'name':chartname})
    chart.set_x_axis({'name':'Wavelength (nm)', 'min':350, 'max': 850})
    chart.set_y_axis({'name':'Absorbance', 'min':max(0,round(miny-0.1,1)), 'max': min(2,round(maxy+0.1,1)), 'minor_unit':0.2})
    if numFile == 1:
        chart.set_legend({'none':True}) 

    ws.insert_chart('%s2'%letters[numFile+2],chart) #put chart in graph


    wb.close() #close excel file

    print('Done!')

keepgoing = 'Y'
while keepgoing == 'Y':
    execute()
    print('Another graph? (Y/N)')
    keepgoing = input()

#Anton Cao 7/14/16
#use python 3.x >= 3.2
#have pip installed
#install xlsxwriter library using "pip install xlsxwriter" command in windows Powershell

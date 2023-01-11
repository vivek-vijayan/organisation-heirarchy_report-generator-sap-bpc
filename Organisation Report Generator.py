'''
Code :Organisation Heirarchy distribution
developer : Vivek Vijayan
Updated by  : Vivek Vijayan
Updated on 02 Sep 2019
--- SERCO _ SPECIFIC ---
'''
from datetime import date
import time, random
import xlsxwriter
import xlrd
import os
import getpass
import threading
from atpbar import atpbar, flush
import win32com.client as win32
from pathlib import Path
import xlsxwriter
import re

win32c = win32.constants
# Global Variables ----
# program variables
VERSION = 1.0
Current_user = None
Filename = "TEMPLATE.xlsx"

ENTITY = None
B_ENTITY = None
EP1 = None
BLACKLIST = None

# Main file engine
FILE_ENGINE = []

# Final report structure
Level1 = "SERCO"
Level1_Name = "Serco"

BlacklistPC = []
BlacklistDivision = []

ENTITY_PC = []
B_ENTITY_PC = [] # 22
B_ENTITY_LOCATION = [] # 20

B_ENTITY_LOCATION_UNIQUE = []

total_valid_members = 0
total_nsdl_locations = 0
# List begins
Level2              = []
Level2_Name         = [] 
Level3              = []
Level3_Name         = [] 
Level4              = []
Level4_Name         = [] 
Level5              = []
Level5_Name         = [] 
Level6              = []
Level6_Name         = [] 
Level7              = []
Level7_Name         = [] 
ProfitCenter        = []
ProfitCenter_Text   = []
Plant               = []
# --- Got in different way for NSDL and its location
Currency            = []
Regional_Currency   = []
PUB                 = []
PUB_Description     = []
Country             = []
Country_Description = []
Segment             = []
Segment_Description = []

get_pub = []
get_pub_desc = []
get_geo = []
get_geo_desc = []
get_seg = []
get_seg_desc = []

# ---------------------------

Heading = [
    "Level 1",
    "Level 1 Description",
    "Level 2",
    "Level 2 Description",
    "Level 3",
    "Level 3 Description",
    "Level 4",
    "Level 4 Description",
    "Level 5",
    "Level 5 Description",
    "Level 6",
    "Level 6 Description",
    "Level 7",
    "Level 7 Description",
    "Profit Center",
    "Profit Center Description",
    "Plant",
    "SAP ECC or NSDL PC",
    "Location",
    "Currency",
    "Regional Currency",
    "Private/Public",
    "Pub Description",
    "Country",
    "Country Description",
    "Segment",
    "Segment Description"
]

heirarchy_heading = [
    "Level 1",
    "Level 2",
    "Level 3",
    "Level 4",
    "Level 5",
    "Level 6",
    "Level 7",
    "Profit Center",
]

FILE_ENGINE.append(EP1)
FILE_ENGINE.append(ENTITY)
FILE_ENGINE.append(B_ENTITY)
FILE_ENGINE.append(BLACKLIST)

ROW_COUNT = [0,0,0,0]
COL_COUNT = [0,0,0,0]

def pivot_table(wb: object, ws1: object, pt_ws: object, ws_name: str, pt_name: str, pt_rows: list, pt_cols: list, pt_filters: list, pt_fields: list, row:int, col:int):
    # pivot table location
    pt_loc = len(pt_filters) + int(row)
    
    # grab the pivot table source data
    pc = wb.PivotCaches().Create(SourceType=win32c.xlDatabase, SourceData=ws1.UsedRange)
    
    # create the pivot table object
    pc.CreatePivotTable(TableDestination=f'{ws_name}!R{pt_loc}C{col}', TableName=pt_name)

    pt_ws.Select()
    pt_ws.Cells(pt_loc, 2).Select()

    for field_list, field_r in ((pt_filters, win32c.xlPageField), (pt_rows, win32c.xlRowField), (pt_cols, win32c.xlColumnField)):
        for i, value in enumerate(field_list):
            pt_ws.PivotTables(pt_name).PivotFields(value).Orientation = field_r
            pt_ws.PivotTables(pt_name).PivotFields(value).Position = i + 1

    for field in pt_fields:
        pt_ws.PivotTables(pt_name).AddDataField(pt_ws.PivotTables(pt_name).PivotFields(field[0]), field[1], field[2]).NumberFormat = field[3]

    pt_ws.PivotTables(pt_name).ShowValuesRow = True
    pt_ws.PivotTables(pt_name).ColumnGrand = False


def run_excel( f_name: str, sheet_name: str, row:int, col: int, item:list, nm:str, close:bool):
    filename = Path.cwd()/f_name
    # create excel object
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    # excel can be visible or not
    excel.Visible = True  # False
    # try except for file / path
    try:
        wb = excel.Workbooks.Open(filename)
    except:
        print("File open error occured")

    # set worksheet
    ws1 = wb.Sheets(sheet_name)
    
    # Setup and call pivot_table
    ws2_name = 'Summary'
    #wb.Sheets.Add().Name = ws2_name
    ws2 = wb.Sheets(ws2_name)
    
    pt_name = nm  
    pt_rows = item  
    pt_cols = []  
    pt_filters = []  
    # [0]: field name [1]: pivot table column name [3]: calulation method [4]: number format
    pt_fields = []
    
    pivot_table(wb, ws1, ws2, ws2_name, pt_name, pt_rows, pt_cols, pt_filters, pt_fields, row, col)
    print("     [ VALIDATING ] - Unlocking the file after validation")
    if close:
        wb.Close(True)
    print("     [ UNLOCKED ] - File unlocked for formatting")

def generate_template():
    workbook    = xlsxwriter.Workbook('TEMPLATE.xlsx')
    A1          = workbook.add_worksheet('EP1 Dump')
    A2          = workbook.add_worksheet('ENTITY Dump')
    A3          = workbook.add_worksheet('B_ENTITY Dump')
    A4          = workbook.add_worksheet('Other info tab')
    A4.write(0,0,"Blacklist PC",)
    A4.write(0,1,"Blacklist Division")
    A4.write(0,2,"PUB ID")
    A4.write(0,3,"PUB Description")
    A4.write(0,4,"GEO")
    A4.write(0,5,"GEO Description")
    A4.write(0,6,"SEGMENT")
    A4.write(0,7, "SEGMENT Description")

    workbook.close()
    os.system('TEMPLATE.xlsx')

# Multi thread execution --------------> ALLOWED

def get_file_data_to_engine(page, name) -> bool:
    global FILE_ENGINE
    global ROW_COUNT
    print("     --> Task " + name +" assigned to thread: {}  - ID: {}".format(threading.current_thread().name,os.getpid()))
    FILE_ENGINE[page] = xlrd.open_workbook(Filename).sheet_by_index(page)
    ROW_COUNT[page] = FILE_ENGINE[page].nrows
    COL_COUNT[page] = FILE_ENGINE[page].ncols
    print("   [ SUCCESS ]   Retrieved data from the sheet ID - " + str(name))


# Main execution flow
if __name__ == "__main__":    

    # clearing the screen
    os.system('cls')

    #generate_template()
    print('''
    ORGANISATION REPORT GENERATOR
    ----------------------------

    Developer : Vivek Vijayan
    Version: 1.0
    
    ''')
    choice = str(input("Do you want to upload template? (y/n)"))
    if choice == 'y':
        Filename = "TEMPLATE.xlsx"
    else:
        Filename = str(input("Enter the filename : "))

    # clearing the screen
    os.system('cls')
    print('''
    ORGANISATION REPORT GENERATOR
    ----------------------------

    Developer : Vivek Vijayan
    Version: 1.0
    
    ''')
    print("Processing the file : " + str(Filename))
    print("-----------------------------------------------------------------------------------------")

    t1 = threading.Thread(target=get_file_data_to_engine, args=(0, "EP1 Data"))
    t2 = threading.Thread(target=get_file_data_to_engine, args=(1, "ENTITY Data"))
    t3 = threading.Thread(target=get_file_data_to_engine, args=(2, "B_ENTITY Data"))
    t4 = threading.Thread(target=get_file_data_to_engine, args=(3, "Other info"))

    t1.start()
    t3.start()
    t2.start()
    t4.start()

    print("\n")
    x = 0 
    sym = ['/', '|', '\\', '-']
    while t1.is_alive() or t2.is_alive() or t3.is_alive():
        if x > 3:
            x = 0
        b = "   Reading data [" + str(sym[x]) +"]"
        print (b, end="\r")
        time.sleep(0.1)
        x = x + 1

    # process waiting
    t1.join()
    t2.join()
    t3.join()
    t4.join()

    # All the data has been recorded in the data engine

    print("\n\n     Generating the Organisation report...\n")
    output_filename = 'ORG REPORT.xlsx'
    output_workbook    = xlsxwriter.Workbook(output_filename)

    if output_workbook:
        print("     [ LOCKED ] - File created and locked for threading")
        Summary         = output_workbook.add_worksheet('Summary')
        Report          = output_workbook.add_worksheet('Organisation Table View')
        NSDL_Report          = output_workbook.add_worksheet('NSDL')
        NSDL_Report.hide()
        NSDL_Report.write(0,0, "NSDL_location")
        NSDL_Report.write(0,1, "PC")
        # Getting the Blacklist file
        for each_line in range(1, ROW_COUNT[3]):
            try:
                BlacklistPC.append(int(FILE_ENGINE[3].cell_value(each_line, 0)))
            except:
                BlacklistPC.append((FILE_ENGINE[3].cell_value(each_line, 0)))
            try:
                BlacklistDivision.append(int(FILE_ENGINE[3].cell_value(each_line, 1)))
            except:
                BlacklistDivision.append((FILE_ENGINE[3].cell_value(each_line, 1)))
            get_pub.append(FILE_ENGINE[3].cell_value(each_line, 2).strip())
            get_pub_desc.append(FILE_ENGINE[3].cell_value(each_line, 3).strip())
            get_geo.append(FILE_ENGINE[3].cell_value(each_line, 4).strip())
            get_geo_desc.append(FILE_ENGINE[3].cell_value(each_line, 5).strip())
            get_seg.append(FILE_ENGINE[3].cell_value(each_line, 6).strip())
            get_seg_desc.append(FILE_ENGINE[3].cell_value(each_line, 7).strip())

        # Counting the total line items
        # ROW_COUNT[0] --> EP1

        # MAIN DATA PUSH HAPPEN HERE        
        for each_line in range(1, ROW_COUNT[0]):
            ID = FILE_ENGINE[0].cell_value(each_line, 0).strip()
            try:
                DIV = int(FILE_ENGINE[0].cell_value(each_line, 2))
            except:
                DIV = (FILE_ENGINE[0].cell_value(each_line, 2))
            try:
                PC = int(FILE_ENGINE[0].cell_value(each_line, 14))
            except:
                PC = (FILE_ENGINE[0].cell_value(each_line, 14))
            if ID == Level1 and DIV not in BlacklistDivision and PC not in BlacklistPC:
                Level2.append((((FILE_ENGINE[0].cell_value(each_line, 2)))))
                Level2_Name.append(str(FILE_ENGINE[0].cell_value(each_line, 3)))
                Level3.append((((FILE_ENGINE[0].cell_value(each_line, 4)))))
                Level3_Name.append(str(FILE_ENGINE[0].cell_value(each_line, 5)))
                Level4.append((((FILE_ENGINE[0].cell_value(each_line, 6)))))
                Level4_Name.append(str(FILE_ENGINE[0].cell_value(each_line, 7)))
                Level5.append((((FILE_ENGINE[0].cell_value(each_line, 8)))))
                Level5_Name.append(str(FILE_ENGINE[0].cell_value(each_line, 9)))
                Level6.append((((FILE_ENGINE[0].cell_value(each_line, 10)))))
                Level6_Name.append(str(FILE_ENGINE[0].cell_value(each_line, 11)))
                Level7.append((((FILE_ENGINE[0].cell_value(each_line, 12)))))
                Level7_Name.append(str(FILE_ENGINE[0].cell_value(each_line, 13)))
                ProfitCenter.append((((FILE_ENGINE[0].cell_value(each_line, 14)))))
                ProfitCenter_Text.append(str(FILE_ENGINE[0].cell_value(each_line, 15)))
                Plant.append(str(FILE_ENGINE[0].cell_value(each_line, 16)).strip())

                total_valid_members = total_valid_members + 1
        # Entering the data as per each line item

        # Formatting
        bold = output_workbook.add_format({'bold': True, 'bg_color' : '#CA0F00', 'font_color': 'white'})
        pcf = output_workbook.add_format({'bg_color' : '#002B5B', 'font_color': 'white'})
        # FINAL OUTPUT 

        # 1. Heading    ------------------------------------------------------------- WRITING REPORT
        for each_line in range(0,len(Heading)):
            Report.write(1, each_line, Heading[each_line] , bold)

        # 2. Data ---------  MAIN  ( ID --> Plant) ---------------------------------- WRITING REPORT
        for each_line in range(2, total_valid_members):
                Report.write(each_line, 0, Level1,  )
                Report.write(each_line, 1, Level1_Name,  )
                Report.write(each_line, 2, Level2[each_line - 2],  )
                Report.write(each_line, 3, Level2_Name[each_line - 2],  )
                Report.write(each_line, 4, Level3[each_line - 2],  )
                Report.write(each_line, 5, Level3_Name[each_line - 2],  )
                Report.write(each_line, 6, Level4[each_line - 2],  )
                Report.write(each_line, 7, Level4_Name[each_line - 2],  )
                Report.write(each_line, 8, Level5[each_line - 2],  )
                Report.write(each_line, 9, Level5_Name[each_line - 2],  )
                Report.write(each_line, 10, Level6[each_line - 2],  )
                Report.write(each_line, 11, Level6_Name[each_line - 2],  )
                Report.write(each_line, 12, Level7[each_line - 2],  )
                Report.write(each_line, 13, Level7_Name[each_line - 2],  )
                Report.write(each_line, 14, ProfitCenter[each_line - 2], pcf )
                Report.write(each_line, 15, ProfitCenter_Text[each_line - 2],  )
                Report.write(each_line, 16, Plant[each_line - 2],  )

        # 3 XXX Getting the data for PC and NSDL location XXX   ------------------------------------------------- B_ENTITY
        for each_line in range(1, ROW_COUNT[2]):
            data = str(FILE_ENGINE[2].cell_value(each_line,1)).strip()
            nsdl_location = str(FILE_ENGINE[2].cell_value(each_line,0)).strip()
            if data not in B_ENTITY_PC and nsdl_location != "":
                B_ENTITY_PC.append(str(data))
                B_ENTITY_LOCATION.append(nsdl_location)

        # 4 XXX Getting the data for PC and Currency, Regional Currency, PUB, Country and Segmental XXX   ------- ENTITY
        for each_line in range(1, ROW_COUNT[1]):
            data = str(FILE_ENGINE[1].cell_value(each_line,0)).strip()
            currency = str(FILE_ENGINE[1].cell_value(each_line,13)).strip()
            reg_currency = str(FILE_ENGINE[1].cell_value(each_line,30)).strip()
            pub = str(FILE_ENGINE[1].cell_value(each_line,29)).strip()
            geo = str(FILE_ENGINE[1].cell_value(each_line,10)).strip()
            seg = str(FILE_ENGINE[1].cell_value(each_line,33)).strip()
            if data not in ENTITY_PC and len(data) == 14: # to confirm it is a PC
                ENTITY_PC.append(str(data))
                Currency.append(currency)
                Regional_Currency.append(reg_currency)
                PUB.append(pub)
                Segment.append(seg)
                Country.append(geo)


        # 4 Data ---- NSDL or SAPECC and its Location --------------------------------------------------- WRITING REPORT
        inc = 1
        for each_line in range(2, total_valid_members):
            pc = ProfitCenter[each_line - 2]
            try:
                pc = "PC000" + str(int(pc)).strip() + "BS"
            except:
                pass
            location = ""
            if pc in B_ENTITY_PC:
                pc_id_in_b_entity = B_ENTITY_PC.index(pc)
                location = B_ENTITY_LOCATION[pc_id_in_b_entity]
                Report.write(each_line, 17, "NSDL",  )
                Report.write(each_line, 18, location,  )
                NSDL_Report.write(inc, 0, str(int(float(str(ProfitCenter[each_line]).strip()))) + " - " +  str(ProfitCenter_Text[each_line]))
                NSDL_Report.write(inc, 1, location)
                inc += 1
            else:
                Report.write(each_line, 17, "SAP ECC",  )
                Report.write(each_line, 18, location,  )

        

        # 5 - XXX Getting the data for Currency, Regional Currency XXX 
        for each_line in range(2, total_valid_members):
            pc = ProfitCenter[each_line - 2]
            try:
                pc = "C_PC000" + str(int(pc)).strip()
            except:
                pass
            if pc in ENTITY_PC:
                pc_id_in_entity = ENTITY_PC.index(pc)
                Report.write(each_line, 19, Currency[pc_id_in_entity] ,  )
                Report.write(each_line, 20, Regional_Currency[pc_id_in_entity],  )
                Report.write(each_line, 21, PUB[pc_id_in_entity],  )
                Report.write(each_line, 23, Country[pc_id_in_entity],  )
                Report.write(each_line, 25, Segment[pc_id_in_entity],  )

                # getting the descriptions
                try:
                    pub_id = get_pub.index(PUB[pc_id_in_entity])
                    Report.write(each_line, 22, get_pub_desc[pub_id],  )
                except:
                    Report.write(each_line, 22, "",  )
                try:
                    geo_id = get_geo.index(Country[pc_id_in_entity])
                    Report.write(each_line, 24, get_geo_desc[geo_id],  )
                except:
                    Report.write(each_line, 24, "",  )
                try:
                    seg_id = get_seg.index(Segment[pc_id_in_entity])
                    Report.write(each_line, 26, get_seg_desc[seg_id],  )
                except:                    
                    Report.write(each_line, 26, "",  )

        print("     [ SUCCESS ] - Organisation Table View Generated successfully")
        
        format1 = output_workbook.add_format({'num_format': ';;;'})
        # heirarchy ___ HIDDEN ____ report creation
        Heir_report          = output_workbook.add_worksheet('HIDDEN')
        Heir_report.hide()

        
        h = xlrd.open_workbook("ORG REPORT.xlsx").sheet_by_index(0)
        for each_line in range(0,len(heirarchy_heading)):
            Heir_report.write(1, each_line, heirarchy_heading[each_line ] , bold)

        for each_line in range(2, total_valid_members):
            Heir_report.write(each_line, 0, str(Level1) + " " + str(Level1_Name),  )
            go_next = True
            row_complete = False
            # from level 2
            m = re.match(r"(\w+)\.(\w+)", str(Level2[each_line-2]).strip() + "d.c")
            m.groups()
            
            if m.groups()[0].isnumeric():
                Heir_report.write(each_line, 1, str(int(float(str((Level2[each_line - 2]))))) + "  " + str(Level2_Name[each_line - 2]) ,  )
                m = re.match(r"(\w+)\.(\w+)", str(Level3[each_line-2]).strip() + "d.c")
                m.groups()
                
                if m.groups()[0].isnumeric():

                    Heir_report.write(each_line, 2, str(int(float(str((Level3[each_line - 2]))))) + "  " + str(Level3_Name[each_line - 2]) ,  )
                    m = re.match(r"(\w+)\.(\w+)", str(Level4[each_line-2]).strip() + "d.c")
                    m.groups()
                    
                    if m.groups()[0].isnumeric():
    
                        Heir_report.write(each_line, 3, str(int(float(str((Level4[each_line - 2]))))) + "  " + str(Level4_Name[each_line - 2]) ,  )
                        m = re.match(r"(\w+)\.(\w+)", str(Level5[each_line-2]).strip() + "d.c")
                        m.groups()
                        
                        if m.groups()[0].isnumeric():
        
                            Heir_report.write(each_line,4, str(int(float(str((Level5[each_line - 2]))))) + "  " + str(Level5_Name[each_line - 2]) ,  )
                            m = re.match(r"(\w+)\.(\w+)", str(Level6[each_line-2]).strip() + "d.c")
                            m.groups()
                            
                            if m.groups()[0].isnumeric():
            
                                Heir_report.write(each_line, 5, str(int(float(str((Level6[each_line - 2]))))) + "  " + str(Level6_Name[each_line - 2]) ,  )
                                m = re.match(r"(\w+)\.(\w+)", str(Level7[each_line-2]).strip() + "d.c")
                                m.groups()
                                
                                if m.groups()[0].isnumeric():
                
                                    Heir_report.write(each_line, 6, str(int(float(str((Level7[each_line - 2]))))) + "  " + str(Level7_Name[each_line - 2]) ,  )
                                else:
                                    Heir_report.write(each_line, 6, str(int(float(str((ProfitCenter[each_line - 2]))))) + "  " + str(ProfitCenter_Text[each_line - 2]) ,  )
                            else:
                                Heir_report.write(each_line, 5, str(int(float(str((ProfitCenter[each_line - 2]))))) + "  " + str(ProfitCenter_Text[each_line - 2]) ,  )
                        else:
                            Heir_report.write(each_line, 4, str(int(float(str((ProfitCenter[each_line - 2]))))) + "  " + str(ProfitCenter_Text[each_line - 2]) ,  )
                    else:
                        Heir_report.write(each_line, 3, str(int(float(str((ProfitCenter[each_line - 2]))))) + "  " + str(ProfitCenter_Text[each_line - 2]) ,  )
                else:
                    Heir_report.write(each_line, 2, str(int(float(str((ProfitCenter[each_line - 2]))))) + "  " + str(ProfitCenter_Text[each_line - 2]) ,  )
            else:
                Heir_report.write(each_line, 1, str(int(float(str((ProfitCenter[each_line - 2]))))) + " " + str(ProfitCenter_Text[each_line - 2]) ,  )

        UNILOCAL = []

        for x in B_ENTITY_LOCATION:
            if x not in UNILOCAL:
                UNILOCAL.append(x)

        # Writing the summary Inforamtion
        h1 = output_workbook.add_format({'bold': True, 'font_size' : 23})
        p = output_workbook.add_format({'bold': True, 'font_size' : 10})
        h2 = output_workbook.add_format({'bold': True, 'font_size' : 11, 'bg_color' : '#B32600', 'font_color': 'white'})
        h22 = output_workbook.add_format({'bold': True, 'font_size' : 11, 'bg_color' : '#002B5B', 'font_color': 'white'})

        Summary.write(2,1,"Organisation Heirarchy Summary",h1)
        Summary.write(3,1,"Developed by Automated program on " + str(time.ctime()), p)
        Summary.write(5,1,"Main Heirarchy view", h2)
        Summary.write(5,4,"NSDL Location - PC Map", h22)


        Summary.hide_gridlines(2)
        output_workbook.close()
        print("     [ LOCKED ] - Locking file for Heirarchy creation")

        # Main heirarchy
        run_excel("ORG REPORT.xlsx", "HIDDEN", 7, 2, ['Level 1', 'Level 2', 'Level 3','Level 4', 'Level 5', 'Level 6', 'Level 7', 'Profit Center'], 'Organisation Heirarchy', True)

        # NSDL Heirarchy
        run_excel("ORG REPORT.xlsx", "NSDL", 7, 5, [ 'PC', 'NSDL_location',], 'NSDL', False)
        
    else:
        print("     [ FAILED ] - Failed to create a new file")


    

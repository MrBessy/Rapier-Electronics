import xlsxwriter 
from datetime import datetime
import subprocess

def create_excel_sheet(BOM, PnP, BareBOM, BarePnP1, BarePnP2=None):
    
    def print_missing(myDict, itemColumn, descColumn):
        
        ind = 0
        for part in myDict:
            worksheet1.write((1 +ind), itemColumn, part)
            description = myDict[part]
            worksheet1.write((1+ind), descColumn, description)
            ind += 1

    def find_missing(dict1, dict2):
        
        missing_dict = {}
        for key in dict1:
            if key not in dict2:
                carry_value = dict1[key]
                missing_dict[key] = carry_value
        print(missing_dict)
        return missing_dict
    
    # Waiting for approval on how new files names should be created. 
    '''
    def create_fileName(file_passed):
        file_name = file_passed.name
        file_name.strip("BOM")
        return file_name

    #fileName = create_fileName(File)
    '''

    '''
    current_time = datetime.now().strftime("%B-%d-%Y_%H-%M-%S")
    file_name = current_time + ".xlsx"
    '''

    # For testing purposes, when ever a file is created it is named 'test'. Remove when ready to push.
    file_name = "Test_mega.xlsx"


    workbook = xlsxwriter.Workbook(file_name)
    worksheet1 = workbook.add_worksheet('Comparrison')
    worksheet2 = workbook.add_worksheet('BOM')                # worksheets 2 and 3 will be needed forr future, but no need for current version. 
    worksheet3 = workbook.add_worksheet('Pick and Place File 1')
    worksheet4 = workbook.add_worksheet('Pick and Place File 2')
    

    # Setting up the columns
    worksheet1.write(0, 0, "#")
    worksheet1.set_column('A:A', 6.43)
    
    worksheet1.write(0, 1, "Designators In PnP")
    worksheet1.set_column('B:B', 17.14)

    worksheet1.write(0, 2, "Designators in BOM")
    worksheet1.set_column('C:C', 17.14)

    worksheet1.write(0, 4, "Missing from BOM")
    worksheet1.set_column('E:E', 18.29)

    worksheet1.write(0, 5, "BOM Description")
    worksheet1.set_column('F:F', 85.00)

    worksheet1.write(0, 6, "Missing form PnP")
    worksheet1.set_column('G:G', 18.29)

    worksheet1.write(0, 7, "PnP Description")
    worksheet1.set_column('H:H', 27.00)

    worksheet2.write(0, 0, BareBOM)

    worksheet3.write(0,0, BarePnP1)

    if not BarePnP2:
        worksheet4.write(0, 0, BarePnP2)
    
    # Define cell formats for good data (green) and bad data (red)
    found_des_format = workbook.add_format({'bold': False, 'bg_color': '8dee73'})
    missing_des_format = workbook.add_format({'bold': True, 'bg_color': 'ee7577'})
    suspicious_des_format = workbook.add_format({'bold': True, 'bg_color': 'eee775'})
    
    
    ### Suspicious values can be designators with a digit within them, eg: DES, XWV, etc.
    def sus_des(des):
        
        digits = [str(i) for i in range(10)]     # string values of numbers from 0 to 9.
        
        # sets the designator to suspicous until it passes logic test
        suspiciousValue = True
        for ind in range(len(des)):
            if des[ind] in digits:
                suspiciousValue = False
                if len(des) <= 5:
                    suspiciousValue = False
                else:
                    suspiciousValue = True
        return suspiciousValue

    rowCount = 1
    for bom_value, pnp_value in zip(BOM, PnP):
        
        if bom_value in PnP:
            # If there's a match, write both values in their respective columns
            pnp_value = bom_value
            if sus_des(pnp_value): 
                worksheet1.write(rowCount, 1, pnp_value, suspicious_des_format)
                worksheet1.write(rowCount, 2, bom_value, suspicious_des_format)

            else:
                worksheet1.write(rowCount, 1, pnp_value, found_des_format)
                worksheet1.write(rowCount, 2, bom_value, found_des_format)
        else:
            # If there's no match, write BOM value in the second column
            worksheet1.write(rowCount, 2, bom_value, missing_des_format)
        rowCount += 1

    for pnp_value in PnP:
        if pnp_value not in BOM:
            # If there's no match, write PnP value in the first column
            worksheet1.write(rowCount, 1, pnp_value, missing_des_format)
            rowCount += 1
        
    BOMMissing = find_missing(BOM, PnP)
    PnPMissing = find_missing(PnP, BOM)
    
    print_missing(BOMMissing, 4, 5)
    print_missing(PnPMissing, 6, 7)

    workbook.close()

    # Open the generated Excel file
    try:
        subprocess.Popen(['start', 'excel', file_name], shell=True)
    except Exception as e:
        print(f"Error: {e}")
import re  # Import the 're' module for regular expressions
import excel
from excel import *

def compareFiles(BOM, PNP1, PNP2, Output="newWindow"):
    
    def split_row(row, bad_values, file=False):
        '''This function is used to split a row further if the rows items are in the bad_values list. 
        If file is set to True, it will also return the new length of the row.'''

        row = row.lower()

        # Split the row using either '\t' or 3 '\s'.
        columns = re.split(r'\s{3,}|\t', row.strip())

        columns = [item for item in columns if item != ""] # Removes any unwanted empty spaces

        # Handles the case when column needs spliting insuffient spacing between headers  
        # using a list of potentilly wrong values.
        for col in range(len(columns)):
            if columns[col] in bad_values:
                split_col = columns[col].split()
                columns[0] = split_col[col]
                columns.insert(1, split_col[(col +1)])

        #when file is set to true, len is also returned
        if file:        
            return columns, len(columns)+1
        else:
            return columns

    # Define a nested function to find the column index containing the designators in the header row
    def find_column_index(header_row, column_name):
        
        header_row_lower = []
        if column_name == 'des':
            acceptable_parameters = ["des", 'designator']
        elif column_name == 'desc':
            acceptable_parameters = ['desc', 'description']         #Could add foot print here but need to makes sure desc is picked first  ..  , 'footprint'

        for i in range(len(header_row)):
            header_row_lower.append(header_row[i].lower())  # Convert the entire header row to lowercase
            for name in range(len(acceptable_parameters)):
                if acceptable_parameters[name] == header_row_lower[i]:  
                    column_idx = i
                    return column_idx   # Return the index of the column containing "Designator"
        
        print("Uh oh")
        return None  # Return None if the column is not found
    
    # Intializing variables to be used
    BOMpartDescriptions = {}
    PnPpartDescriptions = {}
    

    designator_column_idx = None 

    # Define a list of values to ignore or consider as invalid
    bad_values = ["designator comment", "", " ", "designator footprint", "comment description", None]  # Need a more dynaimc list of bad values
    
    # Beginnig Comparsion

    # Open and read the BOM file
    with open(BOM, 'r') as f:
        lines = f.readlines()
        for line in lines:
            
            if "des" in line.lower():  # Checks for "Des" in the line. Note: Description is located on the same line as Designator
                
                header_row = line
                header_row_columns, numOfColumns= split_row(header_row, bad_values, True)
                
                designator_column_idx = find_column_index(header_row_columns, 'des')  # Get the column index for description
                description_column_idx = find_column_index(header_row_columns, "desc")
                columnData = [designator_column_idx, description_column_idx]
                break

            elif 'des' not in line.lower():
                if 'Des' not in line:
                    print("Designator column not found in the BOM file.")
                elif 'Desc' not in line:
                    print("Desciption column not found in the BOM file.")

        # Extract the data from beneath the header row
        data_lines = lines[lines.index(header_row) + 1:]  # Start from the line after the header
        
        temp_line = []
        skipLine = False
        for data in range(len(columnData)):
            
            
            for line in data_lines:
                fields = line.strip().split('\t')  # Split the line into fields using tab as delimiter
                length = len(fields) + 1

                print(f'len of fields: {len(fields)}') 

                if skipLine == True:
                    skipLine = False
                    pass

                elif length < numOfColumns:
                    print(fields)
                    for element in fields:
                        temp_line.append(element)

                    print(f'line len: {len(temp_line)+1} \n{temp_line}')
                    
                    if (len(temp_line)+1) < numOfColumns:
                        skipLine = True

                elif length == numOfColumns:
                    description_value = fields[data]
                    designator_value = fields[columnData[0]].replace('"', '').replace(' ', '').split(",")
                    for i in designator_value:                # Multiple designators may share the same description.
                        BOMpartDescriptions[i] = description_value
                        temp_line = []

                else:
                    print("Death to the stormcloaks")
                    print(f'Cols: {numOfColumns}')
                    print(f'fields: {len(fields)}')
                    print(f'line: {len(temp_line)}')
                """
                
                data_field = fields[column]         # Creates issue when lines in BOM are partial lines
                #partInfo = data_field.split(',')

                if data == designator_column_idx:
                    
                # Add each designator to the set, stripping whitespace and removing quotes
                for des in partInfo:
                    designator = des.strip().replace('"', '')                               # Possibly just delete these lines and 
                                                                                                # just use partDescriptiosn
                    # A basic filter to eliminate any unessesary entries
                    if designator not in bad_values:
                        designators_set.add(designator)  
            
                                                                            elif designator is None:
                        print("There is an issue with designator column being assigned.")
                """
                    #elif data == description_column_idx:
                    
                
                '''if fields[data] == "":
                    description_value = fields[data - 1]        #Shouldnt exist
                else:
                    '''
                
                
            # Wating on approval for which data to display
            """if designator_value not in bad_values:
                comment_and_description = description_value.split('description')

                if len(comment_and_description) > 1:
                   if "comment" in description_column_idx.lower() and "description" in description_column_idx.lower():
                    # Combine the data from both columns
                    comment_and_description = description_value.split('\t')
                    if len(comment_and_description) > 1:
                        combined_data = comment_and_description[0].strip() + " " + comment_and_description[1].strip()
                    else:
                        combined_data = description_value.strip()
                else:
                    combined_data = description_value.strip()

                descriptions[designator_value] = combined_data"""

   
    # Handles if only one file is uploaded to compare against the BOM File
    if PNP2 is None:
        File = [PNP1]
    else:
        File = [PNP1, PNP2]

    # Iterate through the PNP files specified in 'File'
    for fileName in File:
        temp_line = []
        with open(fileName, "r") as infile:
            lines = infile.readlines()

            # lines[0] is used here because a PnP files will have header informaion as the first line.      ##### This is not true
            columns = split_row(lines[0], bad_values)
            for line in lines:
                data = re.split(r'\s{3,}|\t', line.strip())  # Split by three or more spaces

                if len(data) == len(columns):
                    designatorValue = data[0]
                    PnPpartDescriptions[designatorValue] = data[1]
                
                elif (len(data) + len(temp_line)) < len(columns):
                    temp_line = []  
                    for i in data:                                                  # This is A quick fix
                        temp_line.append(i)
                    designatorValue=None              
                    
                elif (len(temp_line) + len(data)) == len(columns):
                    designatorValue = temp_line[0]
                    PnPpartDescriptions[designatorValue] = temp_line[1]
                else:
                    print("error in processing PnP File")
                    pass

        for key in PnPpartDescriptions:
            if key.lower() in bad_values:
                del PnPpartDescriptions[key]
                break

                '''if data[0].lower() not in bad_values:
                    PnPParts_designators_set.add(designatorValue)'''

    print(f'BOM: {BOMpartDescriptions}')
    print("")
    print(f'PNP: {PnPpartDescriptions}')

    ''' # Determine which designators are missing from which file
    missing_from_PnPParts = designators_set - PnPParts_designators_set
    missing_from_BOM = PnPParts_designators_set - designators_set

    #print(f"PNP: {missing_from_PnPParts}")

    for part in missing_from_PnPParts:
        if part in partDescriptions:
            #print(part)
            value_to_move = partDescriptions[part]
            #print(value_to_move)
            missing_parts_description[part] = value_to_move

    for part in missing_from_BOM:
        if part in partDescriptions:
            #print(part)
            value_to_move = partDescriptions[part]
            #print(value_to_move)
            missing_parts_description[part] = value_to_move
    
    # Handle cases where missing designators are empty by setting them to None
    if missing_from_PnPParts == {""}:
        missing_from_PnPParts = None
    if missing_from_BOM == {""}:
        missing_from_BOM = None'''

    # The user will select either to view the results in a new window or in a Excel sheet.
    if Output == "Excel File":

        #Creates a Excel sheet using calculated data.
        create_excel_sheet(BOMpartDescriptions, PnPpartDescriptions)
        

    elif Output == "New Window":
        print("New Window goes here")
        # Also give the option to transfer into Excel from here.

    else:
        print("Error")
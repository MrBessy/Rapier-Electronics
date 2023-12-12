def compareFiles(BOM, PNP1, PNP2, Output="newWindow"):

    File = [PNP1, PNP2]

    PnPParts_designators_set = set()
    for fileName in File:
        with open(fileName, "r") as infile:
            lines = infile.readlines()
            for line in lines:
                data = line.split('\t')
                if len(data) >= 3:
                    designatorValue = data[0].strip().replace('"', '')  # strip off extra whitespaces and quote marks
                    PnPParts_designators_set.add(designatorValue)

    # Load the designators from the BOM file
    designators_set = set()

    with open(BOM, 'r') as f:
        lines = f.readlines()
        data_lines = lines[13:]  # skip non-data lines
        for line in data_lines:
            fields = line.split('\t')

            if len(fields) >= 3:  # Ensure there are enough fields
                designator_field = fields[2]  # 4th field is "Designator"
                designators_for_part = designator_field.split(',')

                # Add each designator to the set
                for designator in designators_for_part:
                    designator = designator.strip().replace('"', '')  # strip off extra whitespaces and quote marks
                    designators_set.add(designator)

    # Determine which designators are missing from which file
    missing_from_PnPParts = designators_set - PnPParts_designators_set
    missing_from_BOM = PnPParts_designators_set - designators_set

    if Output == "Excel File":


        print("In excel file")
        # Print the missing designators
        print("Total Desiginators from PnP: ",len(PnPParts_designators_set))
        print("Total Desiginators from BOM: ",len(designators_set))

        if len(missing_from_PnPParts) > 0:
            for missing in missing_from_PnPParts:
                print("Missing from PnPParts:", missing)
            print("Total Missing from Pnp: ",len(missing_from_PnPParts))
            print(missing_from_PnPParts)

        if len(missing_from_BOM) > 0:
            print("Missing from BOM: ")
            for missing in missing_from_BOM:
                print(missing)
            print("Total Missing from BOM: ",len(missing_from_BOM))

    elif Output == "New Window":

        print("Well damn, i need to create a new window")

        # also give the option  to transfer into excel from here. 

    else: 
        print("Error")
import pandas as pd

"""
GOAL OF THIS SCRIPT:
 We will be formatting the SiteLink all leads activity and the move in activity (over any given period of time)
 to be uploaded into monday.com. This script will add relevant Location Names to the file, select only the important 
 columns, and apply the correct 'category. 
"""

# Site ID's. These are the SiteLink ID's to be changed to 'Location Name' Column.
SITE_ID = {
    33507: "Turtle Creek",
    32937: "Mt. Pleasant",
    33497: "Bridgeville",
    33348: "Warren",
    33876: "Brinton",
    33940: "Murrysville",
    33807: "McKees Rocks",
    34011: "Robinson",
    34107: "New Kensington",
    34232: "Southside",
    34305: "Etna",
    34306: "Office Express",
    43461: "Slippery Rock",
    48377: "Ambridge",
}

# Categories that will be added to the 'Category' column depending on the 'sTypeName' column from SiteLink.
# Changes can be made by either adding an 'sTypeName' (left value) to the list and a corresponding 'Category'(right val)
CATEGORY = {
    'Admin Purpose Only': 'Self Storage',
    'AO Integral Part of Property': 'Self Storage',
    'Drive Up Flex Space': 'Flex Space',
    'Flex Space': 'Flex Space',
    'Heated Motorcycle Parking': 'Motorcycle',
    'Indoor CC Unit': 'Self Storage',
    'Indoor Heated Car Parking': 'Car',
    'Indoor Heated Unit': 'Self Storage',
    'Locker': 'Self Storage',
    'Office': 'Office Space',
    'Outdoor Car Parking': 'Car',
    'Rehearsal Space': 'Rehearsal',
    'Art': 'Art',
    'Drive In CC Unit': 'Self Storage',
    "Drive In Flex Space": "Flex Space",
    "Drive In Flex Space w / Office": "Flex Space",
    "Drive In Heated Unit": "Self Storage",
    "Drive In Unit": "Self Storage",
    "Drive Up CC Unit": "Self Storage",
    "Drive Up Unit": "Self Storage",
    "Flex Space Acces In & Out": "Flex Space",
    "Flex Space w / Office": "Flex Space",
    "Indoor Car Club Parking": "Car",
    "Indoor Heated RV In / Out": "RV",
    "Indoor Heated RV Parking": "RV",
    "Indoor Unit": "Self Storage",
    "Locker - CC": "Self Storage",
    "Outdoor RV Parking": "RV",
    "Unfinished Warehouse": "Flex Space",
    "Wine": "Self Storage",
    "Wine Locker": "Self Storage",
    'Flex Space w/ Office': "Flex Space",
    "Indoor Heated RV In/Out": "RV",
    "Drive In Flex Space w/Office": "Flex Space"
}

# These are the only relevant columns to be uploaded from the SiteLink document.
columnsFromDoc = ["SiteID", "TenantID", "sTypeName", "dPlaced", "dLease", "sTenantName",
                  "sEmployeeName", "sEmployeeConvertedToMoveIn", "sInquiryType", "sMarketingDesc", "sRentalType"]


# Renameing the columns here. Want to change the column name? Add it here
new_column_names = {
    "SiteID": "SiteID",
    "TenantID": "TenantID",
    "sTypeName": "Unit Type (SiteLink)",
    "dPlaced": "Date Placed",
    "dLease": "Lease Date",
    "sTenantName": "Tenant Name",
    "sEmployeeName": "Employee Placed",
    "sEmployeeConvertedToMoveIn": " Employee Moved In",
    "sInquiryType": "Inquiry Type",
    "sMarketingDesc": "Marketing Description",
    "sRentalType": "Rental Type"
}

# !!!! READ IN DATA FROM GIVEN FILE !!!!
df_data = pd.read_excel('file-to-format.xls')

# filter out columns
df_data = df_data[columnsFromDoc]
# Add column and location name
df_data['General Category'] = df_data['sTypeName'].map(CATEGORY)
df_data['Location Name'] = df_data['SiteID'].map(SITE_ID)

# filling empty data with 'other'
df_data['sMarketingDesc'] = df_data['sMarketingDesc'].fillna('Other')

# todo: rename columns for easy import.
df_data.rename(new_column_names, axis=1, inplace=True)


# Export File to excel.
df_data.to_excel("November2022.xls")

print(df_data)


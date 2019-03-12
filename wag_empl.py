# pip install xlrd
import pandas as pd


# Read file
while True:
    filename = input("Name of the file with employees data (with .xlsx): ")
    try:
        data = pd.read_excel(filename)
        break
    except:
        print("no such file in directory, please check the name")

# List of column's names, which we need
column_names = ["Organizační struktura",
                "Cost Center",
                "Business Unit"
                ]

# Select multiple columns, empty cells fill with word "blanks"
empl_data = data[column_names].fillna("blanks") 

# New column names 
new_names = ["company name",
             "profit center",
             "BU"
            ]

# Create mapper dict {column_name: new name}
mapper = {}
for i in range(3):
    mapper[column_names[i]] = new_names[i]

# Rename columns wit new_names
empl_data = empl_data.rename(columns=mapper)

# Use split method to convert string to a list
# and choose the second value (index = 1), which is the company name
empl_data["company name"] = empl_data["company name"].str.split(" - ").map(lambda x: x[1])

# Ask for cost type and add it to new column
cost_type = input("Provide cost type: ")
empl_data["cost type"] = cost_type
new_names.append("cost type")

# Calculate the licence's cost per employee and add it to new column
total_empl = empl_data["company name"].size
while True:
    try:
        licence_cost = float(input("How much are the licences? "))
        break
    except:
        print("You didn't provide valide number")
        print("do not forget to use decimal point instead of comma")
licence_cost_empl = round(licence_cost / total_empl, 2)
empl_data["cost"] = licence_cost_empl

# Check the difference caused by rounding
difference = int(empl_data["cost"].sum() * 100) - int(licence_cost * 100)
for i in range(abs(difference)):
    if difference > 0:
        empl_data["cost"][i] = licence_cost_empl - 0.01
    if difference < 0:
        empl_data["cost"][i] = licence_cost_empl + 0.01
        
# Aggregate data
grouped_data = empl_data.groupby(new_names, as_index=False).sum()
# Delete world "blanks" in column "BU"
grouped_data.loc[grouped_data["BU"]=="blanks", ["BU"]] = ""           

# Save to excel
writer = pd.ExcelWriter('licences.xlsx')
grouped_data.to_excel(writer,sheet_name='data', index=False)
writer.save()

print("********************************************************")
print("New excel file with name licences.xlsx has been created!")
print("********************************************************")


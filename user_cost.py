import pandas as pd


filename = "users_result.xlsx"  # This file was created by users.py

# Read file
try:
    data = pd.read_excel(filename)
except:
    print("no such file in directory, please check the name on line 4")

# List of column's names, which we need
column_names = ["company name",
                "profit center",
                "BU"
                ]

# Select multiple columns
users_data = data[column_names].loc[data["company name"].notnull()] # select only rows, where is a value in column "company name"
users_data = users_data[column_names].fillna("blanks")              # fill empty cell with word "blanks"

# Ask for cost type and add it to new column
cost_type = input("Provide cost type: ")
users_data["cost type"] = cost_type
column_names.append("cost type")

# Calculate the licence's cost per employee and add it to new column
total_users = users_data["company name"].size                       # how many users we have
while True:
    try:
        licence_cost = float(input("How much are the licences? "))
        break
    except:
        print("You didn't provide valide number")
        print("do not forget to use decimal point instead of comma")
licence_cost_user = round(licence_cost / total_users, 2)            # calculates the price per user
users_data["cost"] = licence_cost_user                              # add the unit price to new column "cost"

# Check the difference caused by rounding
difference = int(users_data["cost"].sum() * 100) - int(licence_cost * 100)
for i in range(abs(difference)):
    if difference > 0:
        users_data["cost"][i] = licence_cost_user - 0.01
    if difference < 0:
        users_data["cost"][i] = licence_cost_user + 0.01


# Aggregate data
grouped_data = users_data.groupby(column_names, as_index=False).sum()
grouped_data.loc[grouped_data["BU"]=="blanks", ["BU"]] = ""

# Save to excel
writer = pd.ExcelWriter('user_licences.xlsx')
grouped_data.to_excel(writer,sheet_name='data', index=False)
writer.save()

print("*************************************************************")
print("New excel file with name user_licences.xlsx has been created!")
print("*************************************************************")


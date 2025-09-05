import pandas as pd 

# Read Excel file
excel = pd.read_excel("Book3.xlsx")

print("Full data:")
print(excel)

# --- Statistics ---
print("\nStatistics:")
print("Oldest person:", excel.loc[excel['Age'].idxmax()]['Name'], "-", excel['Age'].max())
print("Youngest person:", excel.loc[excel['Age'].idxmin()]['Name'], "-", excel['Age'].min())
print("Average age:", excel['Age'].mean())

# --- Search Example ---
search_name = "Crawford"  # change this value to test
found = excel[excel['Name'].str.contains(search_name, case=False)]
print("\nSearch results for", search_name)
print(found)

# --- Add New Person ---
new_person = {"Name": "Alice Wonderland", "Nickname": "Ally", "Age": 24}

# ✅ Convert dictionary into a DataFrame before concatenating
new_row = pd.DataFrame([new_person])

# ✅ Now concatenate properly
excel = pd.concat([excel, new_row], ignore_index=True)

# Save the updated file
excel.to_excel("Book3.xlsx", index=False)
print("\nExcel file updated successfully!")


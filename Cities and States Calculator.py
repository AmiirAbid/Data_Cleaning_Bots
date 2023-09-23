import pandas as pd

# Load the Excel file
excel_file = '../Data sheets/FINAL DATA.xlsx'  # Replace with the path to your Excel file
df = pd.read_excel(excel_file)

# Calculate the occurrence of each city with state name
city_state_counts = df.groupby(['Location', 'State']).size().reset_index(name='Count')

# Calculate the occurrence of each state
state_counts = df['State'].value_counts().reset_index()
state_counts.columns = ['State', 'Count']

# Export results to a CSV file
with open('occurrence_results.csv', 'w') as f:
    f.write("City,State,Count\n")
    f.write(city_state_counts.to_csv(index=False, header=False))
    f.write("\nState,Count\n")
    f.write(state_counts.to_csv(index=False, header=False))

print("Results exported to occurrence_results.csv")
# -*- coding: utf-8 -*-
"""
Created on Sat Sep 28 15:15:50 2024

@author: Andres
"""

# List of Excel file paths
files = ['sales_data_2013.xlsx', 'sales_data_2014.xlsx', 'sales_data_2015.xlsx', 'sales_data_2016.xlsx']

# Read each Excel file and store them in a list of DataFrames
data_frames = [pd.read_excel(file) for file in files]

# Concatenate all DataFrames into one
combined_data = pd.concat(data_frames, ignore_index=True)

# Save the combined DataFrame to a new Excel file
combined_data.to_excel('combined_sales_data.xlsx', index=False)

print("Data combined and saved successfully!")

# Read the Excel file
df = pd.read_excel('combined_sales_data.xlsx')

# Filter rows where the Profit is between 500 and 750
filtered_df = df[(df['Profit'] >= 500) & (df['Profit'] <= 750)]

# Save the filtered data to a new Excel file
filtered_df.to_excel('filtered_sales_data.xlsx', index=False)

print("Filtered data saved to 'filtered_sales_data.xlsx'.")

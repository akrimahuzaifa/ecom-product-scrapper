import ast

import pandas as pd

# Read the Excel file
df = pd.read_excel('extracted_products.xlsx')

# Display the first few rows
#print(df.head())

features_column = df['Features']
#print(features_column)

def features_to_html_table(features_str):
    # Convert string to dictionary
    features = ast.literal_eval(features_str)
    table_style = (
        'font-family:arial,sans-serif;'
        'border-collapse:collapse;'
        'width:100%;'
    )
    th_td_style = (
        'border:1px solid #dddddd;'
        'text-align:left;'
        'padding:8px;'
    )
    html = f'<table style="{table_style}">\n'
    html += f'  <tr><th style="{th_td_style}">Feature</th><th style="{th_td_style}">Value</th></tr>\n'
    for i, (key, value) in enumerate(features.items()):
        row_bg = 'background-color:#dddddd;' if i % 2 == 1 else ''
        html += (
            f'  <tr style="{row_bg}">'
            f'<td style="{th_td_style}">{key}</td>'
            f'<td style="{th_td_style}">{str(value).replace(chr(10), "<br>")}</td>'
            f'</tr>\n'
        )
    html += '</table>'
    return html

for index, row in features_column.items():
    # Convert each features string to an HTML table
    html_table = features_to_html_table(row)
    # Save the HTML table to a new column in the DataFrame
    df.at[index, 'Features_HTML'] = html_table

# Save the updated DataFrame to a new Excel file
df.to_excel('extracted_products_with_html_tables.xlsx', index=False)
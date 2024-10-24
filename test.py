from datetime import datetime
import pandas as pd
import os
from openpyxl import Workbook

# Load the Excel file (replace with your actual file path)
file_path = 'your_excel_file.xlsx'  
df = pd.read_excel(file_path)

# Extract the relevant columns
df_filtered = df[['CUSTOMER NAME', 'PROVINCE', 'STYLECLRSIZE', 'UNIT PRICE', 'QTY', 'ORDER DATE']]

# Group the orders by 'CUSTOMER NAME' and 'ORDER DATE'
grouped_orders = df_filtered.groupby(['CUSTOMER NAME', 'ORDER DATE'])

# Create a directory to store the order files
output_dir = 'orders_output'  # You can change this to your desired directory
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Helper function to sanitize file names
def sanitize_filename(filename):
    return filename.replace(':', '-').replace('/', '-').replace('\\', '-')

# Function to write order information into the template
def write_order_to_template(customer_name, province, order_data, output_file, start_item=0, end_item=None):
    # Create a new workbook or load a template if you have one
    wb = Workbook()
    ws = wb.active

    # Set column widths
    column_widths = {'A': 2.83, 'B': 6.66, 'C': 8.33, 'D': 16.83, 'E': 7.16, 'F': 16.5, 
                    'G': 11.66, 'H': 20.83, 'I': 16.66, 'J': 14.16, 'K': 18.33, 'L': 6.5}

    for col, width in column_widths.items():
        ws.column_dimensions[col].width = width

    # Set row heights
    row_heights = {1: 11.45, 2: 11.45, 3: 11.45, 4: 11.45, 5: 14.25, 6: 11.45, 
                7: 11.45, 8: 13.5, 9: 15.0, 10: 4.5, 11: 12.0, 12: 17.25, 
                13: 21.75, 16: 12.75, 17: 11.25, 18: 11.25, 19: 11.25, 
                20: 11.25, 21: 11.25, 22: 11.25, 23: 11.25, 24: 11.25, 
                25: 11.25, 26: 11.25, 27: 9.0, 28: 10.5, 29: 10.5, 30: 10.5, 
                31: 10.5, 32: 10.5, 33: 10.5, 34: 10.5, 35: 9.75, 36: 7.5, 
                37: 12.0, 38: 15.0, 39: 18.75, 40: 30.0}

    for row, height in row_heights.items():
        if height:  # Only set height if it is defined
            ws.row_dimensions[row].height = height
    
    # Write customer information
    ws['D9'] = customer_name
    ws['D11'] = province

    # Set date
    ws['K11'] = datetime.now().strftime("%B %d, %Y")

    # Set order number
    ws.merge_cells(f'I13:J13')
    # ws['I13'] = order_number

    # Starting row for styles
    start_row = 15
    total_value = 0  # Initialize the order total
    p_count = 0  # Initialize the order count

    # If end_item is None, process all items
    if end_item is None:
        end_item = len(order_data)

    # Loop through each row in the order data to populate styles
    for index, row in order_data.iloc[start_item:end_item].iterrows():
        style_total = row['QTY'] * row['UNIT PRICE']
        total_value += style_total  # Add to the total value
        p_count += row['QTY']  # add qty to running total

        # Merge columns C and D for Style and Color
        ws.merge_cells(f'C{start_row}:D{start_row}')
        ws[f'C{start_row}'] = row['STYLECLRSIZE']  # Merged Style and Color and Size

        ws[f'H{start_row}'] = row['QTY']  # Quantity
        ws[f'I{start_row}'] = row['UNIT PRICE']  # Unit Price
        ws[f'J{start_row}'] = style_total  # Total (QTY * Unit Price)

        start_row += 1

    # Calculate the number of dashes to add (10 minus the number of styles in the order)
    num_styles = end_item - start_item
    num_dashes = max(10 - num_styles, 0)

    # Add dashes in column J for the remaining rows (10 - number of styles)
    for i in range(num_dashes):
        ws[f'J{start_row}'] = '-'
        start_row += 1

    # Write the total order value in column F on the row after the last dash
    ws[f'F{start_row}'] = total_value / 1.12  # VAT less
    start_row += 2

    ws[f'F{start_row}'] = total_value  # with VAT
    start_row += 1
    
    ws[f'F{start_row}'] = total_value * 0.12  # VAT
    start_row += 1

    ws[f'H{start_row}'] = p_count
    ws[f'I{start_row}'] = 'P'
    ws[f'J{start_row}'] = total_value

    # Save the file
    wb.save(output_file)


# Save each order to a separate Excel file, splitting if there are more than 10 items
for (customer, date), group in grouped_orders:
    customer_name = group['CUSTOMER NAME'].iloc[0]
    province = group['PROVINCE'].iloc[0]

    # Sanitize the filename to avoid invalid characters
    date_str = str(date).replace(':', '-').replace(' ', '_')
    customer_str = sanitize_filename(customer_name)

    # Number of items in the order
    num_items = len(group)
    
    # Process in batches of 10 items
    batch_size = 10
    batch_number = 1
    for start_item in range(0, num_items, batch_size):
        end_item = min(start_item + batch_size, num_items)
        
        file_name = f"{customer_str}_{date_str}_batch{batch_number}.xlsx"
        output_file = os.path.join(output_dir, file_name)
        
        # Write the order data to the template
        write_order_to_template(customer_name, province, group, output_file, start_item, end_item)
        batch_number += 1

print(f"Order files saved to {output_dir}")

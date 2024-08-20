import smartsheet
import re
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)

# Define the Smartsheet API access token
access_token = 'XM9QwMkgayOSMo0fQVVrbkKqA3JutUFDPgdgz'

# Initialize the Smartsheet client
smartsheet_client = smartsheet.Smartsheet(access_token)

# Function to get column ID by column title
def get_column_id(sheet, column_title):
    for column in sheet.columns:
        if column.title == column_title:
            return column.id
    raise ValueError(f"Column '{column_title}' not found in the sheet.")

# Function to read and parse the Results.txt file
def parse_results(file_path):
    results = []
    with open(file_path, 'r', encoding='utf-8') as file:
        content = file.read()
        
        # Split content by sequence sections
        sequence_blocks = content.split('Results for file:')
        
        for block in sequence_blocks:
            sequence_match = re.search(r'Sequence:\s*(\w+)', block)
            if sequence_match:
                sequence = sequence_match.group(1)
                result = {'Sequence': f'SEQ {sequence}'}
                
                # Extract each activity percentage and adjust the percentage
                for activity in ['Processing', 'Fabrication', 'Paint', 'Shipping']:
                    activity_match = re.search(f'{activity}:\s*([\d\.]+)%', block)
                    if activity_match:
                        result[activity] = float(activity_match.group(1)) / 100  # Convert to decimal percentage
                    else:
                        result[activity] = 0.0
                
                results.append(result)
    
    logging.info(f"Parsed results: {results}")
    return results

# Function to update Smartsheet with the parsed data
def update_smartsheet(sheet_id, results):
    # Get the sheet
    sheet = smartsheet_client.Sheets.get_sheet(sheet_id)
    
    # Retrieve column IDs
    primary_column_id = get_column_id(sheet, 'Primary Column')
    activity_column_id = get_column_id(sheet, 'Activity')
    tekla_complete_column_id = get_column_id(sheet, 'Tekla % Complete')
    
    rows_to_update = []
    
    # Iterate through the parsed results
    for result in results:
        sequence = result['Sequence']
        logging.info(f"Processing sequence: {sequence}")
        
        # Find the row that contains this sequence in the Primary Column
        sequence_row = None
        for row in sheet.rows:
            for cell in row.cells:
                if cell.column_id == primary_column_id and cell.value == sequence:
                    sequence_row = row
                    break
            if sequence_row:
                break

        if sequence_row:
            logging.info(f"Found sequence row ID: {sequence_row.id}")
            # Find and update each corresponding activity
            for i, activity in enumerate(['Processing', 'Fabrication', 'Paint', 'Shipping']):
                # Get the row corresponding to this activity (next rows)
                activity_row = sheet.rows[sheet.rows.index(sequence_row) + i]
                activity_cell = next((cell for cell in activity_row.cells if cell.column_id == activity_column_id and cell.value == activity), None)
                
                if activity_cell:
                    # Update the Tekla % Complete column for this activity
                    tekla_complete_cell = next((cell for cell in activity_row.cells if cell.column_id == tekla_complete_column_id), None)
                    if tekla_complete_cell:
                        percentage = result[activity]
                        logging.debug(f"Updating row {activity_row.id} - Sequence: {sequence_row.id}, Activity: {activity}, Percentage: {percentage}%")
                        tekla_complete_cell.value = percentage
                        rows_to_update.append(smartsheet.models.Row({
                            'id': activity_row.id,
                            'cells': [tekla_complete_cell]
                        }))
        else:
            logging.warning(f"No matching row found for sequence: {sequence}")
    
    if rows_to_update:
        # Update the rows in the sheet
        smartsheet_client.Sheets.update_rows(sheet_id, rows_to_update)
        logging.info("Updated rows in Smartsheet successfully.")
    else:
        logging.warning("No rows were updated.")

# Main function to execute the script
def main():
    # Define the Smartsheet sheet ID
    sheet_id = '7479818832531332'
    
    # Path to the Results.txt file
    results_file_path = input("Enter the file path for the results file: ").strip().strip('"').strip("'")
    
    # Parse the results
    results = parse_results(results_file_path)
    
    # Update Smartsheet with the parsed results
    update_smartsheet(sheet_id, results)

if __name__ == "__main__":
    main()

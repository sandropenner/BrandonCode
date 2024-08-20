import os
import re
import sys
from datetime import datetime

def parse_station_summary_with_weight_and_save(file_path, output_file):
    with open(file_path, 'r', encoding='utf-8-sig') as file:
        lines = file.readlines()

    job_number = None
    sequence = None
    stations = {}
    
    current_station = None
    
    for line in lines:
        # Look for Job # and Sequence
        if "Job #:" in line:
            job_number = line.split("Job #:")[1].strip()
        if "Sequence:" in line:
            sequence = line.split("Sequence:")[1].strip()
            sequence = re.sub(r'[^\x00-\x7F]+', '', sequence)  # Remove special characters

        # Identify station names
        station_match = re.match(r'\s*(Fit/Bolt|Weld|QC|Inspected|Process|Stored|Ship Ready)', line)
        if station_match:
            current_station = station_match.group(1)
            stations[current_station] = {"Completed_Weight": 0, "Remaining_Weight": 0}
        
        # Capture the weight data
        weight_match = re.findall(r'(\d{1,3}(?:,\d{3})*)#', line)
        if weight_match and current_station:
            if len(weight_match) == 2:  # Ensure there are exactly two weights
                completed_weight = int(weight_match[0].replace(',', ''))
                remaining_weight = int(weight_match[1].replace(',', ''))
                stations[current_station]["Completed_Weight"] += completed_weight
                stations[current_station]["Remaining_Weight"] += remaining_weight

    # Calculate completion percentage by weight
    station_results = {}
    for station, data in stations.items():
        total_weight = data['Completed_Weight'] + data['Remaining_Weight']
        if total_weight > 0:
            completion_percentage = data['Completed_Weight'] / total_weight * 100
        else:
            completion_percentage = 0.0
        station_results[station] = completion_percentage

    # Rename stations as needed for the output
    station_name_mapping = {
        "Process": "Processing",
        "Weld": "Fabrication"
    }

    # Append results to the output file
    output_file.write(f"Results for file: {os.path.basename(file_path)}\n")
    output_file.write(f"Job #: {job_number}\n")
    output_file.write(f"Sequence: {sequence}\n\n")
    output_file.write("Station Completion Percentages:\n")
    for station, percentage in station_results.items():
        # Apply the name mapping if applicable
        output_station_name = station_name_mapping.get(station, station)
        output_file.write(f"{output_station_name}: {percentage:.2f}%\n")
    output_file.write("\n" + "-"*50 + "\n\n")

def process_directory(directory_path, output_directory):
    if not os.path.isdir(directory_path):
        print(f"Error: The path '{directory_path}' is not a directory.")
        sys.exit(1)

    if not os.path.isdir(output_directory):
        print(f"Error: The output directory '{output_directory}' is not valid.")
        sys.exit(1)

    # Get current date to use as the output file name
    current_date = datetime.now().strftime('%Y-%m-%d')
    output_file_name = f"p-results-{current_date}.txt"
    output_file_path = os.path.join(output_directory, output_file_name)
    
    with open(output_file_path, 'w', encoding='utf-8') as output_file:
        # Walk through all directories and subdirectories
        for root, dirs, files in os.walk(directory_path):
            for filename in files:
                if filename.endswith(".txt"):
                    file_path = os.path.join(root, filename)
                    parse_station_summary_with_weight_and_save(file_path, output_file)

    print(f"All results have been saved to {output_file_path}")

if __name__ == "__main__":
    # Ask the user for the path to the directory and the output directory
    directory_path = input("Please enter the path to the directory containing the text files: ").strip('\'"')
    output_directory = "N:\\Production\\Production Status"  # Fixed path for the results file
    process_directory(directory_path, output_directory)

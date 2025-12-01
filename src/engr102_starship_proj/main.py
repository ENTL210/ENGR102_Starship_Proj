import os
import time
import json 
import random
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from pathlib import Path
import datetime


def generate_answer_keys(export_directory):
    print("-" * 75)
    print((" " * 22) + "[2] Generating Answer Keys" + (" " * 22))
    
    path = Path(export_directory)
    datasets_count = 0
    result = []
    
    # Getting numbers of datapoints...
    datapoint_counts = len(generate_time_stamps())
    
    # Check if there is an dataset...
    for item in path.iterdir():
        if item.is_dir() and "Dataset" in item.name:
            datasets_count += 1
    
    # If there is no dataset...
    if datasets_count == 0:
        print("\n       There is no dataset in this directory.")
        print("\n       An answer key can't be generated. Exiting in 3 seconds...")
        print("-" * 75)
        for i in range(0,3):
            print(f"          Exiting program in {3-i}...") 
            time.sleep(1)
        print("-" * 75)
        exit()
    else: 
        print(f"\n      Numbers of datasets found: {datasets_count}")
        print("\n      Begin to loop through each dataset...")
    
    for item in path.iterdir():
        if item.is_dir() and "Dataset" in item.name:
            print(f"\n      {item.name}")
            print(f"\n           Counts of data points: {datapoint_counts}")
            
            # Calculating avg and max of temperature...
            temperature_file_name = "temps_in_C.xlsx"
            temperature_file_path = os.path.join(item, temperature_file_name)
            print(f"\n           Calculating avg & max temperatures:")
            
            # Loading the workbook...
            wb = load_workbook(temperature_file_path)
            ws = wb.active
            temperature_values = []
            for row_num in range(2, datapoint_counts + 2):
                cell_value = ws.cell(row_num, 2).value
                temperature_values.append(cell_value)
            avg_temperature_value = round(sum(temperature_values) / len(temperature_values), 2)
            max_temperature = round(max(temperature_values), 2)
            print(f"              Avg Temperature in C: {avg_temperature_value}°C")
            print(f"              Max Temperature in C: {max_temperature}°C")
            
            # Sum standard power input...
            standard_power_input_file_name = "power_input_std.xlsx"
            standard_power_input_file_path = os.path.join(item, standard_power_input_file_name)
            print(f"\n           Summing the total power input of the standard model:")
            # Loading the workbook...
            wb = load_workbook(standard_power_input_file_path)
            ws = wb.active
            standard_power_input_values = []
            for row_num in range(2, datapoint_counts + 2):
                cell_value = ws.cell(row_num, 2).value
                standard_power_input_values.append(cell_value)
            sum_of_standard_power_input = round(sum(standard_power_input_values), 2)
            print(f"              Sum of Power Input of the Standard Model: {sum_of_standard_power_input} Watts-Minutes")
            
            # Sum PV power input...
            pv_power_input_file_name = "power_input_pv.xlsx"
            pv_power_input_file_path = os.path.join(item, pv_power_input_file_name)
            print(f"\n           Summing the total power input of the proposed model:")
            # Loading the workbook...
            wb = load_workbook(pv_power_input_file_path)
            ws = wb.active
            pv_power_input_values = []
            for row_num in range(2, datapoint_counts + 2):
                cell_value = ws.cell(row_num, 2).value
                pv_power_input_values.append(cell_value)
            sum_of_pv_power_input = round(sum(pv_power_input_values), 2)
            print(f"              Sum of Power Input of the Proposed Model: {sum_of_pv_power_input} Watts-Minutes")
            
            
            
            
            
            
        
                    
    
    
def generate_random_nums(min, avg, max, counts):
    # Initialize array of nums...
    nums = []
    
    # Randomize numbers until reaching the target number of numbers...
    while len(nums) < counts:
        randomized_number = random.uniform(min, max)
        nums.append(randomized_number)
    
    # Find average of the set of randomized numbers...
    nums_average = sum(nums) / len(nums)
    # Find the difference between the target average and the randomized average...
    average_difference = round(avg - nums_average, 0) # Round to 0 deci points to add some noises to the dataset
    # Adjusted randomized numbers to reach the approximate avg...
    adjusted_nums = [num + average_difference for num in nums] 
    
    return adjusted_nums

def generate_time_stamps():
    timestamps = []
    # Set time interval where data being collected (7:30 AM - 10:00 PM)
    start = datetime.datetime.strptime('7:30:00', '%H:%M:%S')
    end = datetime.datetime.strptime('22:00:00', '%H:%M:%S')
    
    # Set delta between point of collecting...
    delta = datetime.timedelta(minutes=5)
    
    # Inititalize while loop with t...
    t = start
    while t <= end:
        timestamps.append(f"{t.strftime('%H:%M:%S')}") # Converting datetime obj to str...
        t += delta # Add 5 minutes...
    
    return timestamps
    

def generate_problem(export_directory, data):
    print("-" * 75) # seperator
    print((" " * 25) + "[1] Generating Problems" + (" " * 25))
    print("-" * 75) # seperator
    
    nums_of_datasets = len(data)
    
    for dataset in data:
        
        # The directory to the dataset that in progress of making...
        making_dataset = None
        
        while True:
            # Base on numbers of datasets in json file, random the dataset index...
            dataset_index = random.randint(1,4)
            # base on randomized dataset index, generate a folder path to that dataset...
            dataset_path = os.path.join(export_directory, f"Dataset {dataset_index}")
            
            # if the dataset directory is already exists, randomize another index and path...
            if os.path.exists(dataset_path):
                continue
            else:
                # if the dataset directory chosen is not exists...
                os.mkdir(dataset_path)
                making_dataset = dataset_path
                break
        
        print(f"\n       Generating {os.path.basename(making_dataset)}...")
        
        print(f"\n       Generated Dataset Path: {making_dataset}")
        
        # Declaring key to be ignored...
        skipped_key_value = None
        key_to_ignore = "station_power_input"
        
        # Loop through all the keys except station_power_input...
        for key in dataset:
            start_time = time.time()
            
            if key == "season":
                continue
            if key == key_to_ignore:
                skipped_key_value = key
                continue
            
            # Initializing workbook...
            wb_basename = key
            wb_path = os.path.join(making_dataset, f"{key}.xlsx")
            wb = Workbook()
            ws = wb.active
            
            # Generating (min, avg, max) of each data file...
            print(f"\n       Generating {key}.xlsx...")
            print(f"\n            (", end=" ")
            for key,value in dataset[key].items():
                print(f"{key.capitalize()}: {value}", end=" ")
            print(f")")
            
            print("\n            Generating timestamp column...")
            
            # Generating timestamp list...
            time_stamps = generate_time_stamps()
            
            # Generating timestamp column & label...
            time_stamps_column = 1
            ws.column_dimensions[get_column_letter(time_stamps_column)].width = 20
            ws.cell(1, 1, "Time (%HH:%MM:%SS)")
            
            for index, value in enumerate(time_stamps, 1):
                ws.cell(index + 1, time_stamps_column, value)
                
            print("\n            Generating values column...")
                
            # Generating randomized numbers list...
            values = generate_random_nums(dataset[wb_basename]["min"], dataset[wb_basename]["avg"], dataset[wb_basename]["max"], len(time_stamps))
            
            # Generating value column & label...
            values_column = 2
            ws.column_dimensions[get_column_letter(values_column)].width = 20
            ws.cell(1, 2, f"Values ({dataset[wb_basename]["unit"] if "unit" in dataset[wb_basename] else ""})")
            
            for index, value in enumerate(values, 1):
                ws.cell(index + 1, values_column, value)
                
            wb.save(wb_path)

            end_time = time.time()
            # set elapsed time to ms and 2 deci points...
            elapsed_time = round((end_time - start_time) * 1000, 3) 
            # print out the elapsed time...
            print(f"\n            {wb_basename}.xlsx has been generated ({elapsed_time}ms)") 
            
        # Generating station_power_input_std.xlsx & modify power_input_pv
        if skipped_key_value != None:
            start_time = time.time()
            basename = skipped_key_value
            
            # Initialize workbooks
            wb = Workbook()
            ws = wb.active
            wb_path = os.path.join(making_dataset, f"power_input_std.xlsx")
            
            # Generating (min, avg, max) of each data file...
            print(f"\n       Generating power_input_std.xlsx...")
            print(f"\n            (", end=" ")
            for key,value in dataset[basename].items():
                print(f"{key.capitalize()}: {value}", end=" ")
            print(f")")
            
            print("\n            Generating timestamp column...")
            
            # Generating timestamp list...
            time_stamps = generate_time_stamps()
            
            # Generating timestamp column & label...
            time_stamps_column = 1
            ws.cell(1, 1, "Time (%HH:%MM:%SS)")
            ws.column_dimensions[get_column_letter(time_stamps_column)].width = 20
            
            for index, value in enumerate(time_stamps, 1):
                ws.cell(index + 1, time_stamps_column, value)
                
            print("\n            Generating values column...")
                
            # Generating value column & label...
            values_column = 2
            ws.column_dimensions[get_column_letter(values_column)].width = 20
            ws.cell(1, 2, f"Values ({dataset[basename]["unit"] if "unit" in dataset[basename] else ""})")
            
            # Set values columns to 0 initially...
            current_row = 1
            while current_row <= len(time_stamps):
                ws.cell(current_row + 1, values_column, 0)
                current_row += 1
                
            # Numbers of times robo went to the charging stations...
            total_nums_of_charging_charging_intervals = random.randint(20, len(time_stamps) // 3)

            # When robot first charge of the day...
            start_interval = 0
            
            for current_nums_of_charging_intervals in range(1, total_nums_of_charging_charging_intervals + 1):
                # Randomized how long the robot will be charged...
                duration_of_intervals = random.randint(1,3)
                end_interval = start_interval + duration_of_intervals
                
                while start_interval <= end_interval:
                    # Getting value from json...
                    value = dataset[basename]["value"]
                    # Add noise
                    value = generate_random_nums(value - 0.5, value, value + 0.5, 1)[0]
                    ws.cell(start_interval + 2, values_column, value)

                    start_interval += 1
                
                start_interval = random.randint(1, len(time_stamps))
                current_nums_of_charging_intervals += 1
                
            wb.save(wb_path)
            
            end_time = time.time()
            # set elapsed time to ms and 2 deci points...
            elapsed_time = round((end_time - start_time) * 1000, 3) 
            # print out the elapsed time...
            print(f"\n            {wb_basename}.xlsx has been generated ({elapsed_time}ms)") 
            
        print("-" * 75)        

def read_data_json():
    with open('data.json', 'r') as f:
        return json.load(f)["data"]

def get_folder_path(test_path="/Users/edwardl210/Documents/Coding/ENGR102_Starship_Proj/src/engr102_starship_proj"):
    print((" " * 22) + "Selecting Export Destination" + (" " * 22))
    print("-" * 75)

    # While loop until a valid path is returned...
    while True:
        if test_path:
            path = test_path
        else:
            path = input("\n       Please provide a export destination: ")

        print(f"\n       Provided Path: {path}")
        confirmation = input("\n       Is this correct: (1) Yes (2) No: ")

        if confirmation == "1":
            if os.path.isdir(path):
                print("\n       Path has been verified")
                print("")
                
                export_path = os.path.join(path, "engr102_starship_problems")
                os.makedirs(os.path.join(export_path, "Answer Keys"), exist_ok=True)
                
                return export_path
            else: 
                print("       Provided path is invalid. Please try again")
                continue
        elif confirmation == '2':
            print("       Path rejected. Please try entering the correct path again.")
            continue
        else:
            print("       Invalid selection. Please enter '1' for Yes or '2' for No.")
            
def get_user_options():
    
    # While loop until users choose a valid option...
    while True:
        print("-" * 75) # Separator
        # Print the initial greeting only once
        print((" " * 10) + "Welcome to the ENGR102 Starship Problem Generator." + (" " * 10))
        print("-" * 75) # Separator
        print("\n       Options: ")
        print("         [1] Generate Problems") # Generates 4 days of data..
        print("         [2] Generate Answers Key") # Generates 4 days of data..
        print("         [3] Quit Tool\n") # Quit tools
        print("-" * 75) # Seperator
        selection = input("       Enter a menu option in the keyboard [1,2,3]: ")
        
        print("-" * 75)
        
        if selection == "1":
            return "1"
        elif selection == "2":
            return "2"
        elif selection == "3":
            return "3"
        else: 
            print("Invalid selection. Please choose one of the options from the menu")
            continue

def main():
    # Get user selection...
    user_selection = get_user_options()
    
    if user_selection == "3":
        print("-" * 75)
        for i in range(0,3):
            print(f"          Exiting program in {3-i}...") 
            time.sleep(1)
        print("-" * 75)
        exit()
    
    # Reading data.json
    data = read_data_json()
    
    if user_selection == "1":
        # Get export destination
        export_path = get_folder_path()
        generate_problem(export_directory=export_path, data=data)
        
    if user_selection == "2":
        # Get the location of the datasets
        export_path = get_folder_path()
        generate_answer_keys(export_directory=export_path)
    
    
    
    
if __name__ == "__main__":
    main()
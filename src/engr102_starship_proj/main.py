import os
import time
import json 
import random
from openpyxl import Workbook
import datetime

def generate_answer_keys():
    print("-" * 75)
    print((" " * 22) + "[2] Generating Answer Keys" + (" " * 22))
    
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
        timestamps.append = f"{t}" # Converting datetime obj to str...
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
        
        for key in dataset:
            start_time = time.time()
            
            if key == "season":
                continue
            if key == "station_power_input":
                continue
            
            # Initializing workbook...
            wb_path = os.path.join(os.path.basename(making_dataset), f"{key}.xlsx")
            wb = Workbook()
            ws = wb.active
            
            # Generating (min, avg, max) of each data file...
            print(f"\n       Generating {key}.xlss...")
            print(f"\n            (", end=" ")
            for key,value in dataset[key].items():
                print(f"{key.capitalize()}: {value}", end=" ")
            print(f")")
            
            # Generating timestamp list...
            time_stamps = generate_time_stamps()
            
            # Generating timestamp column & label...
            time_stamps_column = 1
            ws.cell(1, 1, "Time (%HH:%MM:%SS)")
            
            for index, value in enumerate(time_stamps, 1):
                ws.cell(index + 1, time_stamps_column, value)
            
            

            end_time = time.time()
            # set elapsed time to ms and 2 deci points...
            elapsed_time = round((end_time - start_time) * 1000, 3) 
            # print out the elapsed time...
            print(f"\n            {os.path.basename(wb_path)} has been generated ({elapsed_time}ms)") 
            
            
        
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
                return path
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
        print("         [1] Generate problems") # Generates 4 days of data..
        print("         [2] Quit Tool\n") # Quit tools
        print("-" * 75) # Seperator
        selection = input("       Enter a menu option in the keyboard [1,2]: ")
        
        print("-" * 75)
        
        if selection == "1":
            return "1"
        elif selection == "2":
            return "2"
        else: 
            print("Invalid selection. Please choose one of the options from the menu")
            continue

def main():
    # Get user selection...
    user_selection = get_user_options()
    
    if user_selection == "2":
        print("-" * 75)
        for i in range(0,3):
            print(f"          Exiting program in {3-i}...") 
            time.sleep(1)
        print("-" * 75)
        exit()
    
    # Get export destination...
    export_path = get_folder_path()
    
    # Reading data.json
    data = read_data_json()
    
    generate_problem(export_directory=export_path, data=data)
    
    
    

if __name__ == "__main__":
    main()
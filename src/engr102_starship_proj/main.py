import os
import time
import json 
import random
from openpyxl import Workbook

def generate_answer_keys():
    print("-" * 75)
    print((" " * 22) + "[2] Generating Answer Keys" + (" " * 22))

def generate_problem(export_directory, data):
    # Initialize openpyxl Wb...
    wb = Workbook()
    
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
            if key == "season":
                continue
            print(f"\n       Generated :")
        
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
        selection = input("       Enter a menu option in the keyboard [1,2,3,4]: ")
        
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
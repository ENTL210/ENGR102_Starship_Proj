import os
import json 

folder_path = ""
data = {}

def read_data_json():
    with open('data.json', 'r') as f:
        return json.load(f)

def get_folder_path():
    # Print the initial greeting only once
    print("Welcome to the ENGR102 Starship Problem Generator.")
    print("-" * 45) # Separator

    # While loop until a valid path is returned...
    while True:
        print("\nPlease provide a folder path where you would like to save the generated problem: ", end="")
        path = input()

        print(f"\nProvided Path: {path}")
        confirmation = input("\nIs this correct: (1) Yes (2) No: ")

        if confirmation == "1":
            if os.path.isdir(path):
                print("\nPath confirmed and verified.")
                return path 
            else:
                print(f"Error: '{path}' is not a valid directory. Please check and try again.")
        elif confirmation == '2':
            print("Path rejected. Please try entering the correct path again.")
            continue
        else:
            print("Invalid selection. Please enter '1' for Yes or '2' for No.")

def main():
    # Greeting & request for folder path...
    # folder_path = get_folder_path()

    # Reading data.json
    data = read_data_json()
    




    
    
    

if __name__ == "__main__":
    main()
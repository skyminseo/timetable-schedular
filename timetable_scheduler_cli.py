import csv
import os
from datetime import datetime


# Get Timetable Data
class TimetableData:
    def __init__(self, Description, Module_Code, Study_Mode, Cohort, Allocated_Location_Name, Planned_Size, Allocated_Staff_Name, Zone_Name, Activity_Dates_Individual, Scheduled_Days, Scheduled_Start_Time, Scheduled_End_Time, Duration, Class_Type):
        # Make all attributes private with leading underscores
        self.__Description = Description
        self.__Module_Code = Module_Code
        self.__Study_Mode = Study_Mode
        self.__Cohort = Cohort
        self.__Allocated_Location_Name = Allocated_Location_Name
        self.__Planned_Size = Planned_Size
        self.__Allocated_Staff_Name = Allocated_Staff_Name
        self.__Zone_Name = Zone_Name
        self.__Activity_Dates_Individual = Activity_Dates_Individual
        self.__Scheduled_Days = Scheduled_Days
        self.__Scheduled_Start_Time = Scheduled_Start_Time
        self.__Scheduled_End_Time = Scheduled_End_Time
        self.__Duration = Duration
        self.__Class_Type = Class_Type

    def get_items(self):
        # Return all attributes as a dictionary
        return {
            'Description': self.__Description,
            'Module_Code': self.__Module_Code,
            'Study_Mode': self.__Study_Mode,
            'Cohort': self.__Cohort,
            'Allocated_Location_Name': self.__Allocated_Location_Name,
            'Planned_Size': self.__Planned_Size,
            'Allocated_Staff_Name': self.__Allocated_Staff_Name,
            'Zone_Name': self.__Zone_Name,
            'Activity_Dates_Individual': self.__Activity_Dates_Individual,
            'Scheduled_Days': self.__Scheduled_Days,
            'Scheduled_Start_Time': self.__Scheduled_Start_Time,
            'Scheduled_End_Time': self.__Scheduled_End_Time,
            'Duration': self.__Duration,
            'Class_Type': self.__Class_Type
        }

    def __str__(self):
        return f"""
Module Name: {self.__Description}
Module Code: {self.__Module_Code}
Study Mode: {self.__Study_Mode}
Cohort: {self.__Cohort}
Location: {self.__Allocated_Location_Name} ({self.__Zone_Name})
Planned Size: {self.__Planned_Size}
Lecturer: {self.__Allocated_Staff_Name}
Schedule: {self.__Activity_Dates_Individual}
Scheduled Day: {self.__Scheduled_Days}
Start Time: {self.__Scheduled_Start_Time}
End Time: {self.__Scheduled_End_Time}
Duration: {self.__Duration}
Class Type: {self.__Class_Type}
"""


# Implement Heap Sort and Binary Search Algorithm
class TimetableManager:
    def __init__(self):
        self.data_manager = DataManager()

    def heap_sort(self, arr, reverse=False):
        n = len(arr)

        for i in range(n // 2 - 1, -1, -1):
            self.heapify(arr, n, i, reverse)

        for i in range(n - 1, 0, -1):
            arr[i], arr[0] = arr[0], arr[i]
            self.heapify(arr, i, 0, reverse)

    def heapify(self, arr, n, i, reverse=False):
        largest = i
        left = 2 * i + 1
        right = 2 * i + 2

        # Compare based on the 'Activity_Dates_Individual' attribute as datetime objects
        if left < n and ((datetime.strptime(arr[i].get_items()["Activity_Dates_Individual"], "%d/%m/%Y") < datetime.strptime(arr[left].get_items()["Activity_Dates_Individual"], "%d/%m/%Y")) if not reverse else (datetime.strptime(arr[i].get_items()["Activity_Dates_Individual"], "%d/%m/%Y") > datetime.strptime(arr[left].get_items()["Activity_Dates_Individual"], "%d/%m/%Y"))):
            largest = left

        if right < n and ((datetime.strptime(arr[largest].get_items()["Activity_Dates_Individual"], "%d/%m/%Y") < datetime.strptime(arr[right].get_items()["Activity_Dates_Individual"], "%d/%m/%Y")) if not reverse else (datetime.strptime(arr[largest].get_items()["Activity_Dates_Individual"], "%d/%m/%Y") > datetime.strptime(arr[right].get_items()["Activity_Dates_Individual"], "%d/%m/%Y"))):
            largest = right

        if largest != i:
            arr[i], arr[largest] = arr[largest], arr[i]
            self.heapify(arr, n, largest, reverse)

    def binary_search(self, csv_filename, search_key, search_criteria):
        data = self.data_manager.data_by_file[csv_filename]["timetable_data_list"]
        left, right = 0, len(data) - 1
        results = []

        while left <= right:
            mid = (left + right) // 2
            current_item = data[mid].get_items()[search_criteria]

            if current_item == search_key:
                results.append(data[mid])  # Found a match
                # Continue searching for additional matches to the left
                left = mid - 1
                while left >= 0 and data[left].get_items()[search_criteria] == search_key:
                    results.append(data[left])
                    left -= 1
                # Continue searching for additional matches to the right
                right = mid + 1
                while right < len(data) and data[right].get_items()[search_criteria] == search_key:
                    results.append(data[right])
                    right += 1
                break
            elif current_item < search_key:
                left = mid + 1
            else:
                right = mid - 1

        return results  # Return list of schedules matching the search criteria


# Filter the Data
class DataManager:
    def __init__(self):
        self.data_by_file = {}

    def data_filter(self, csv_filename):
        # Initialize the data dictionary for this file
        self.data_by_file[csv_filename] = {
            "timetable_data_list": []
        }

        with open(csv_filename, 'r') as csv_file:
            csv_reader = csv.reader(csv_file)
            for column in csv_reader:
                name = column[1]
                parts = name.split("_")
                # Check if there are enough parts before accessing indices. Prevent (IndexError: list index out of range)
                if len(parts) >= 2:
                    timetable_data = TimetableData(
                        Description=column[2],
                        Module_Code=parts[3],
                        Study_Mode=parts[2],
                        Cohort=parts[0] + " " + parts[1],
                        Allocated_Location_Name=column[8],
                        Planned_Size=column[9],
                        Allocated_Staff_Name=column[10],
                        Zone_Name=column[11],
                        Activity_Dates_Individual=column[3],
                        Scheduled_Days=column[4],
                        Scheduled_Start_Time=column[5],
                        Scheduled_End_Time=column[6],
                        Duration=column[7],
                        Class_Type=parts[4]
                    )

                    data = self.data_by_file[csv_filename]
                    data["timetable_data_list"].append(timetable_data)

    def list_schedules_by_module_name(self, csv_filename, module_name):
        data = self.data_by_file[csv_filename]
        filtered_schedules = []
        for timetable_data in data["timetable_data_list"]:
            if module_name in timetable_data.get_items()["Description"]:
                filtered_schedules.append(timetable_data)

        return filtered_schedules

    def list_schedules_by_lecturer_name(self, csv_filename, lecturer_name):
        data = self.data_by_file[csv_filename]
        filtered_schedules = []
        for timetable_data in data["timetable_data_list"]:
            if lecturer_name in timetable_data.get_items()["Allocated_Staff_Name"]:
                filtered_schedules.append(timetable_data)

        return filtered_schedules

    def list_schedules_by_date_range(self, csv_filename, start_date, end_date):
        data = self.data_by_file[csv_filename]
        start_date = datetime.strptime(start_date, "%d/%m/%Y")
        end_date = datetime.strptime(end_date, "%d/%m/%Y")
        filtered_schedules = []
        for timetable_data in data["timetable_data_list"]:
            schedule_date = datetime.strptime(
                timetable_data.get_items()["Activity_Dates_Individual"], "%d/%m/%Y")
            if start_date <= schedule_date <= end_date:
                filtered_schedules.append(timetable_data)

        return filtered_schedules

    def list_schedules_by_location(self, csv_filename, location_name):
        data = self.data_by_file[csv_filename]
        filtered_schedules = []
        for timetable_data in data["timetable_data_list"]:
            if location_name in timetable_data.get_items()["Allocated_Location_Name"]:
                filtered_schedules.append(timetable_data)

        return filtered_schedules

    def list_schedules_by_specific_time(self, csv_filename, specific_time):
        data = self.data_by_file[csv_filename]
        filtered_schedules = []
        for timetable_data in data["timetable_data_list"]:
            if specific_time in timetable_data.get_items()["Scheduled_Start_Time"]:
                filtered_schedules.append(timetable_data)

        return filtered_schedules

    def list_schedules_by_duration(self, csv_filename, duration):
        data = self.data_by_file[csv_filename]
        filtered_schedules = []
        for timetable_data in data["timetable_data_list"]:
            if duration in timetable_data.get_items()["Duration"]:
                filtered_schedules.append(timetable_data)

        return filtered_schedules

    def list_schedules_by_day(self, csv_filename, day):
        data = self.data_by_file[csv_filename]
        filtered_schedules = []
        for timetable_data in data["timetable_data_list"]:
            if day in timetable_data.get_items()["Scheduled_Days"]:
                filtered_schedules.append(timetable_data)

        return filtered_schedules

    def print_data(self, csv_filename):
        data = self.data_by_file[csv_filename]["timetable_data_list"]
        for timetable_data in data:
            print(timetable_data)
        print(f"Data from: {csv_filename}\n")


# Main Function
class Main:
    def __init__(self):
        self.timetable_manager = TimetableManager()

    def load_csv_files(self, directory_path):
        csv_filepaths = []  # Create a list to store valid CSV file paths
        # List all files in the directory
        all_files = os.listdir(directory_path)

        # Filter for CSV files by checking the file extension
        csv_files = [file for file in all_files if file.endswith('.csv')]

        if not csv_files:
            print("No CSV files found in the selected directory.")
            return []

        # Iterate through the list of CSV files and create the full file paths
        for csv_filename in csv_files:
            csv_filepath = os.path.join(directory_path, csv_filename)
            csv_filepaths.append(csv_filepath)  # Add the file path to the list
        print("Loaded CSV files:", list(csv_filepaths))

        return csv_filepaths

    def run(self):
        directory_path = input("Enter the directory path: ")
        csv_filepaths = self.load_csv_files(directory_path)

        if not csv_filepaths:
            # If no CSV files found, ask the user to select a valid path again
            while True:
                directory_path = input("Enter a valid directory path again: ")
                csv_filepaths = self.load_csv_files(directory_path)
                if csv_filepaths:
                    break

        # Check if the user wants to select a second path
        select_second_path = input(
            "Do you want to select a second directory? (y/n): ").lower()

        if select_second_path == 'y':
            directory_path2 = input("Enter the second directory path: ")
            csv_filepaths2 = self.load_csv_files(directory_path2)

            if not csv_filepaths2:
                # If no CSV files found, ask the user to select a valid path again
                while True:
                    directory_path2 = input(
                        "Enter a valid directory path again: ")
                    csv_filepaths2 = self.load_csv_files(directory_path2)
                    if csv_filepaths2:
                        break
        else:
            csv_filepaths2 = []

        # Load and process all CSV files in the specified directories
        for csv_filepath in csv_filepaths:
            self.timetable_manager.data_manager.data_filter(csv_filepath)

        for csv_filepath in csv_filepaths2:
            self.timetable_manager.data_manager.data_filter(csv_filepath)

        while True:
            print("Options:")
            print("1. Search schedules by Module Name")
            print("2. Search schedules by Lecturer Name")
            print("3. Search schedules by Location")
            print("4. Search schedules by Specific Time")
            print("5. Search schedules by Duration")
            print("6. Search schedules by Day")
            print("7. Print All Schedules")
            print("8. Quit")

            choice = input("Enter your choice: ")

            if choice == "1":
                module_name = input("Enter Module Name to search: ")
                self.search_schedules("Description", module_name)

            elif choice == "2":
                lecturer_name = input("Enter Lecturer Name to search: ")
                self.search_schedules("Allocated_Staff_Name", lecturer_name)

            elif choice == "3":
                location_name = input("Enter Location Name to search: ")
                self.search_schedules(
                    "Allocated_Location_Name", location_name)

            elif choice == "4":
                specific_time = input(
                    "Enter Specific Time to search (HH:MM:SS): ")
                self.search_schedules("Scheduled_Start_Time", specific_time)

            elif choice == "5":
                duration = input("Enter Duration to search: ")
                self.search_schedules("Duration", duration)

            elif choice == "6":
                day = input("Enter Day to search: ")
                self.search_schedules("Scheduled_Days", day)

            elif choice == "7":
                for csv_filename in self.timetable_manager.data_manager.data_by_file:
                    self.timetable_manager.data_manager.print_data(
                        csv_filename)

            elif choice == "8":
                break

            else:
                print("Invalid choice. Please select a valid option.")

    def search_schedules(self, search_criteria, search_key):
        sort_option = input(
            "Select Sorting Option \n1. Ascending Order\n2. Descending Order\nEnter your choice:")
        ascending = True
        if sort_option == "2":
            ascending = False

        for csv_filename in self.timetable_manager.data_manager.data_by_file:
            results = self.timetable_manager.binary_search(
                csv_filename, search_key, search_criteria)
            if len(results) > 0:
                print(
                    f"\nSchedules found for '{search_criteria}' with '{search_key}' in path '{csv_filename}':")
                # Sort the results based on user's choice
                self.timetable_manager.heap_sort(
                    results, reverse=not ascending)
                for result in results:
                    print(result)


if __name__ == "__main__":
    main = Main()
    main.run()

import csv
import os
from datetime import datetime
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
from openpyxl import Workbook
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.platypus import PageBreak


# Manage and Get Timetable Data
class TimetableData:
    def __init__(self, Description, Module_Code, Study_Mode, Cohort, Allocated_Location_Name, Planned_Size, Allocated_Staff_Name, Zone_Name, Activity_Dates_Individual, Scheduled_Days, Scheduled_Start_Time, Scheduled_End_Time, Duration, Class_Type):
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
            'Module Name': self.__Description,
            'Module Code': self.__Module_Code,
            'Study Mode': self.__Study_Mode,
            'Cohort': self.__Cohort,
            'Location': f"{self.__Allocated_Location_Name}({self.__Zone_Name})",
            'Planned Size': self.__Planned_Size,
            'Lecturer': self.__Allocated_Staff_Name,
            'Scheduled Date': self.__Activity_Dates_Individual,
            'Scheduled Day': self.__Scheduled_Days,
            'Lecture Start Time': self.__Scheduled_Start_Time,
            'Lecture End Time': self.__Scheduled_End_Time,
            'Duration': self.__Duration,
            'Class Type': self.__Class_Type
        }


# Implement Heap Sort and Binary Search Algorithm
class DataManager:
    def __init__(self):
        self.timetable_manager = TimetableManager()

    # 'reverse' parameter for descending and 'key' parameter to specify sorting key
    def heapify(self, arr, n, i, reverse=False, key=lambda x: x):
        largest = i     # Initialize largest as root
        leftChild = 2 * i + 1
        rightChild = 2 * i + 2

        # Detect if left child of root exists and which is greater than root
        if leftChild < n and ((key(arr[i].get_items()) < key(arr[leftChild].get_items())) if not reverse else (key(arr[i].get_items()) > key(arr[leftChild].get_items()))):
            largest = leftChild

        if rightChild < n and ((key(arr[largest].get_items()) < key(arr[rightChild].get_items())) if not reverse else (key(arr[largest].get_items()) > key(arr[rightChild].get_items()))):
            largest = rightChild

        # If the largest element is not the root, swap them
        if largest != i:
            arr[i], arr[largest] = arr[largest], arr[i]     # Swap the elements
            self.heapify(arr, n, largest, reverse, key)

    # The main function to heap sort an array
    def heap_sort(self, arr, reverse=False, key=lambda x: x):
        n = len(arr)

        # Build a max-heap
        for i in range(n // 2 - 1, -1, -1):
            self.heapify(arr, n, i, reverse, key)

        # Swap the root (largest element) with the current last element
        for i in range(n - 1, 0, -1):
            arr[i], arr[0] = arr[0], arr[i]     # Swap
            self.heapify(arr, i, 0, reverse, key)

    def binary_search(self, csv_filename, search_key, search_criteria):
        data = self.timetable_manager.data_by_file[csv_filename]["timetable_data_list"]

        # Initialize the left and right pointers for binary search
        left, right = 0, len(data) - 1

        results = []

        while left <= right:
            # Calculate the middle index
            mid = (left + right) // 2

            # Get the value of the search criteria for the current item at the middle index
            current_item = data[mid].get_items()[search_criteria]

            # Check if the current item matches the search key
            if current_item == search_key:
                results.append(data[mid])

                # Continue searching for additional matches to the left of the current item
                left = mid - 1
                while left >= 0 and data[left].get_items()[search_criteria] == search_key:
                    results.append(data[left])
                    left -= 1

                right = mid + 1
                while right < len(data) and data[right].get_items()[search_criteria] == search_key:
                    results.append(data[right])
                    right += 1

                # Break out of the loop, as it found all matching items
                break
            elif current_item < search_key:
                # If the current item is less than the search key, update the left pointer
                left = mid + 1
            else:
                right = mid - 1

        return results


# Filter the Timetable Data Items
class TimetableManager:
    def __init__(self):
        self.data_by_file = {}

    def data_filter(self, csv_filename):
        # Initialize the timetable data dictionary for this file
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


# Implementation of the GUI
class Window:

    search_term_suggestions = {
        "Module Name": [
            "SET Computer Hacking Forensics Investigator (CHFI)",
            "SET Data Communication & Networking (DCNG)",
            "SET Data Communication & Networking",
            "SET Database Design & Modelling (DDMG)",
            "SET Database Design & Modelling",
            "SET Discrete Mathematics (DM)",
            "SET Ethical Hacking & Countermeasures (EHC)",
            "SET Introduction to Computer Operating Systems (Group A)",
            "SET Information Systems and the Organisations (ISOG)",
            "SET Introduction to Software Engineering (ISEG)",
            "SET Introduction to Programming (IP)",
            "SET Network Defense (ND)",
            "SET Security Specialist (SS)"
        ],
        "Module Code": [
            "CHFI",
            "DDMG",
            "DCNG",
            "DDMG",
            "DM",
            "EHC",
            "ICOS",
            "IP",
            "ISEG",
            "ISOG",
            "ND",
            "SS"
        ],
        "Study Mode": [
            "FT", "PT"
        ],
        "Cohort": ["DICT-DNDFC 221"],
        "Location": [
            "A01(Marina)",
            "B05/B06(Marina)",
            "C01(Jackson)",
            "C02/C03(Jackson)",
            "D01(Jackson)",
            "D03(Jackson)",
            "D06(Jackson)",
            "LT3(Marina)",
            "MBA1(Marina)",
            "Online Learning(Marina)",
            "Online Learning(Jackson)"
        ],
        "Planned Size": [
            "20", "24", "30", "48", "50", "60", "80", "84", "90", "100"
        ],
        "Lecturer": [
            "Ankit Saurabh",
            "Chang Wing Hong Edmund",
            "Dr Liau Vui Kien",
            "Dr Tan Boon Leing",
            "Kelvin Wu",
            "Lee Han John",
            "Lau Jun Tian Terence",
            "Ong Chin Ann",
            "Yegna Ramanchandran Vijayalakshmi",
            "Yee Sook Liang"
        ],
        "Scheduled Day": [
            "Monday", "Tuesday", "Wednesday", "Thursday", "Friday"
        ],
        "Lecture Start Time": [
            "8:30:00",
            "08:30:00",
            "12:00:00",
            "13:00:00",
            "15:30:00",
            "19:00:00"
        ],
        "Lecture End Time": [
            "11:30:00",
            "15:00:00",
            "15:30:00",
            "18:00:00",
            "18:30:00",
            "22:00:00"
        ],
        "Duration": [
            "2:30",
            "03:00",
            "3:00",
            "3:30"
        ],
        "Class Type": [
            "Lab02GrpA/1", "Lec09", "Lec15", "Lab03GrpA/1", "Lec12", "Lec06A/1", "Lab03GrpB/1", "Lec01", "Lec06",
            "Lec04B/1", "Lec06B/1", "Lab02GrpB/1", "Lec10", "Lec08 (Lab Grp1)", "Lab05GrpB/1", "Lab05GrpA/1", "Lec04",
            "Lab04GrpB/1", "Lec07", "Lab04GrpA/1", "Lec04A/1", "Lab01GrpB/1", "Lab01GrpA/1", "Lec02", "Lab02Grp01/1",
            "Lec01/1 (Group B)", "Lec01/1", "Lab04Grp01/1", "Lab05Grp01/1", "Lab06Grp01/1", "Lec01/1 (Group A)",
            "Lec02/1 (Group B)", "Lec02/1", "Lec02/1 (Group A)", "Lab01Grp02/1", "Lab03Grp02/1", "Lab02Grp03/1", "Lec03/1",
            "Lab03Grp03/1", "Lec03/1 (Group B)", "Lec03/1 (Group A)", "Lec04/1", "Lec04/1 (Group A)", "Lec04/1 (Group B)",
            "Lec05/1", "Lec05/1 (Group B)", "Lec05/1 (Group A)", "Lec06/1 (Group B)", "Lec06/1 (Group A)", "Lec06/1",
            "Lec07/1 (Group A)", "Lec07/1 (Group B)", "Lec07/1", "Lec08/1", "Lec08/1 (Group A)", "Lec08/1 (Group B)",
            "Lec09/1 (Group B)", "Lec09/1", "Lec09/1 (Group A)", "Lec10/1 (Group B)", "Lec10/1", "Lec11/1", "Lec11/1 (Group A)",
            "Lec11/1 (Group B)", "Lec12/1 (Group A)", "Lec12/1 (Group B)", "Lec12/1", "Lec13/1 (Group B)", "Lec13/1 (Group A)",
            "Lec13/1", "Lec14/1", "Lec14/1 (Group A)", "Lec14/1 (Group B)", "Lec15/1 (Group A)", "Lec15/1", "Lec15/1 (Group B)",
            "Lec11 (Lab Grp 1)", "Lec05 (Lab Grp 1)", "Lec14 (Lab Grp 1)", "Lec13 (Lab Grp 1)", "Lec03 (Lab Grp 1)",
            "Lec13 (Lab Grp 2)", "Lec11 (Lab Grp 2)", "Lec08 (Lab Grp 2)", "Lec14 (Lab Grp 2)", "Lec03 (Lab Grp 2)", "Lec05 (Lab Grp 2)"
        ]
    }

    def __init__(self, root):
        self.timetable_manager = DataManager()
        self.folder_paths = []

        self.root = root
        self.root.title("Timetable Viewer & Generator")
        self.root.geometry("1000x700")

        self.selected_folder_path = ""

        self.folder_path_label = tk.Label(root, text="")
        self.folder_path_label.pack()

        self.searched_data = []

        # Create a frame of the treeview
        tree_frame = tk.Frame(root)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        # Create the treeview widget
        self.tree = ttk.Treeview(tree_frame, show="headings", columns=(
            "Module Name",
            "Module Code",
            "Study Mode",
            "Cohort",
            "Location",
            "Planned Size",
            "Lecturer",
            "Scheduled Date",
            "Scheduled Day",
            "Lecture Start Time",
            "Lecture End Time",
            "Duration",
            "Class Type"
        ), height=10)

        # Set default column widths
        column_widths = [
            180, 80, 80, 100, 150, 80, 100, 100, 100, 100, 100, 80, 100
        ]
        for i, column_id in enumerate(self.tree["columns"]):
            self.tree.column(column_id, width=column_widths[i])

        self.tree.heading("#1", text="Module Name")
        self.tree.heading("#2", text="Module Code")
        self.tree.heading("#3", text="Study Mode")
        self.tree.heading("#4", text="Cohort")
        self.tree.heading("#5", text="Location")
        self.tree.heading("#6", text="Planned Size")
        self.tree.heading("#7", text="Lecturer")
        self.tree.heading("#8", text="Scheduled Date")
        self.tree.heading("#9", text="Scheduled Day")
        self.tree.heading("#10", text="Lecture Start Time")
        self.tree.heading("#11", text="Lecture End Time")
        self.tree.heading("#12", text="Duration")
        self.tree.heading("#13", text="Class Type")

        # Create vertical scrollbar
        v_scrollbar = ttk.Scrollbar(
            tree_frame, orient="vertical", command=self.tree.yview)
        v_scrollbar.pack(side="right", fill="y")

        # Create horizontal scrollbar
        h_scrollbar = ttk.Scrollbar(
            tree_frame, orient="horizontal", command=self.tree.xview)
        h_scrollbar.pack(side="bottom", fill="x")

        self.tree.configure(yscrollcommand=v_scrollbar.set)
        self.tree.configure(xscrollcommand=h_scrollbar.set)

        self.tree.pack(fill=tk.BOTH, expand=True)

        # Create a frame for buttons
        button_frame = tk.Frame(root)
        button_frame.pack(pady=15)

        load_button = tk.Button(
            button_frame, text="Load CSV Folder", command=self.load_csv)
        load_button.grid(row=0, column=1, padx=10, pady=20)

        add_folder_button = tk.Button(
            button_frame, text="Add CSV Folder", command=self.add_csv_folder)
        add_folder_button.grid(row=0, column=3, padx=10, pady=20)

        self.criteria_label = tk.Label(
            button_frame, text="Select Search Criteria:")
        self.criteria_label.grid(row=1, column=0, padx=10, pady=10)

        self.criteria_var = tk.StringVar()
        self.criteria_var.set("Module Name")  # Default search criteria
        criteria_list = [
            "Module Name",
            "Module Code",
            "Study Mode",
            "Cohort",
            "Location",
            "Planned Size",
            "Lecturer",
            "Scheduled Day",
            "Lecture Start Time",
            "Lecture End Time",
            "Duration",
            "Class Type"
        ]

        # Create a combobox for search criteria
        criteria_dropdown = ttk.Combobox(
            button_frame, textvariable=self.criteria_var, values=criteria_list)
        criteria_dropdown.grid(row=1, column=1, padx=10, pady=10)
        self.criteria_var.trace_add('write', self.criteria_change)

        search_label = tk.Label(button_frame, text="Search Term:", width=10)
        search_label.grid(row=1, column=2, padx=10, pady=10)

        # Create an entry widget for search term
        self.search_entry = ttk.Combobox(button_frame, width=35)
        self.search_entry.grid(row=1, column=3, padx=10, pady=10)

        search_button = tk.Button(
            button_frame, text="Search", command=self.display_searched_data)
        search_button.grid(row=1, column=4, padx=10, pady=10)

        search_label = tk.Label(button_frame, text="Sort By:")
        search_label.grid(row=2, column=0, padx=10, pady=10)

        # Create a combobox for sorting attribute
        self.sort_attribute_var = tk.StringVar()
        self.sort_attribute_var.set("Module Name")
        sort_attribute_dropdown = ttk.Combobox(
            button_frame, textvariable=self.sort_attribute_var, values=criteria_list)
        sort_attribute_dropdown.grid(row=2, column=1, padx=10, pady=10)

        sort_order_label = tk.Label(
            button_frame, text="Sort Order:")
        sort_order_label.grid(row=2, column=2, padx=10, pady=10)

        # Create a combobox for sorting order
        self.sort_order_var = tk.StringVar()
        self.sort_order_var.set("Ascending")
        sort_order_combobox = ttk.Combobox(
            button_frame, textvariable=self.sort_order_var, values=["Ascending", "Descending"])
        sort_order_combobox.grid(row=2, column=3, padx=10, pady=10)

        sort_button = tk.Button(
            button_frame, text="Sort", command=self.display_sorted_data)
        sort_button.grid(row=2, column=4, padx=10, pady=10)

        self.show_all_button = tk.Button(
            button_frame, text="Show All Schedules", command=self.show_all_data)
        self.show_all_button.grid(row=3, column=1, padx=10, pady=10)

        sort_order_label = tk.Label(
            button_frame, text="Export Option:")
        sort_order_label.grid(row=3, column=2, padx=10, pady=10)

        self.export_options = ttk.Combobox(
            button_frame, values=["Excel (.xlsx)", "PDF (.pdf)"])
        self.export_options.set("Excel (.xlsx)")
        self.export_options.grid(row=3, column=3, padx=10, pady=10)

        export_button = tk.Button(
            button_frame, text="Export", command=self.export_data)
        export_button.grid(row=3, column=4, padx=10, pady=10)

        # Bind header click events to sort by the selected attribute
        for col_id in self.tree["columns"]:
            self.tree.heading(
                col_id, command=lambda c=col_id: self.sort_by_attribute(c))

    def get_loaded_data(self):
        loaded_data = []
        for item in self.tree.get_children():
            values = self.tree.item(item, 'values')
            if len(values) >= 14:
                timetable_data = TimetableData(
                    Description=values[0],
                    Module_Code=values[1],
                    Study_Mode=values[2],
                    Cohort=values[3],
                    Allocated_Location_Name=values[4],
                    Planned_Size=values[5],
                    Allocated_Staff_Name=values[6],
                    Zone_Name=values[7],
                    Activity_Dates_Individual=values[8],
                    Scheduled_Days=values[9],
                    Scheduled_Start_Time=values[10],
                    Scheduled_End_Time=values[11],
                    Duration=values[12],
                    Class_Type=values[13]
                )
                loaded_data.append(timetable_data)
        return loaded_data

    def add_csv_folder(self):
        folder_path = filedialog.askdirectory()

        if not folder_path:
            return

        # Check if there are CSV files in the selected folder
        csv_files = [file for file in os.listdir(
            folder_path) if file.endswith('.csv')]

        if not csv_files:
            # No CSV files found in the selected directory, show an error message
            messagebox.showerror(
                "No CSV Files Found", "No CSV files were found in the selected directory.\nPlease select a folder with CSV files.")
            return

        self.folder_paths.append(folder_path)

        self.load_csv_data(folder_path)

        self.folder_path_label.config(
            text=f"Selected Folders: {', '.join(self.folder_paths)}")

    def load_csv_data(self, folder_path):
        csv_filepaths = []
        all_files = os.listdir(folder_path)

        csv_files = [file for file in all_files if file.endswith('.csv')]

        if not csv_files:
            messagebox.showerror(
                "No CSV Files Found", "No CSV files were found in the selected directory.\nPlease select a valid folder.")
            return

        for csv_filename in csv_files:
            csv_filepath = os.path.join(folder_path, csv_filename)
            csv_filepaths.append(csv_filepath)

        for csv_filepath in csv_filepaths:
            self.timetable_manager.data_filter(csv_filepath)

        self.loaded_data = self.get_loaded_data()
        self.display_data()

    def load_csv(self):
        # Ask the user to select a directory
        directory_path = filedialog.askdirectory()

        if not directory_path:
            return

        self.folder_paths.append(directory_path)

        self.folder_path_label.config(
            text=f"Selected Folders: {', '.join(self.folder_paths)}")

        self.load_csv_data(directory_path)

    def load_csv_data(self, folder_path):
        csv_filepaths = []
        all_files = os.listdir(folder_path)

        csv_files = [file for file in all_files if file.endswith('.csv')]

        if not csv_files:
            messagebox.showerror(
                "No CSV Files Found", "No CSV files were found in the selected directory.\nPlease select a valid folder.")
            return

        for csv_filename in csv_files:
            csv_filepath = os.path.join(folder_path, csv_filename)
            csv_filepaths.append(csv_filepath)

        for csv_filepath in csv_filepaths:
            self.timetable_manager.timetable_manager.data_filter(csv_filepath)

        self.loaded_data = self.get_loaded_data()
        self.display_data()

    def show_all_data(self):
        self.display_data()
        self.searched_data = []

    def display_data(self):
        self.tree.delete(*self.tree.get_children())

        for csv_filename in self.timetable_manager.timetable_manager.data_by_file:
            data = self.timetable_manager.timetable_manager.data_by_file[
                csv_filename]["timetable_data_list"]

            for timetable_data in data:
                items = timetable_data.get_items()
                self.tree.insert("", "end", values=(
                    items["Module Name"],
                    items["Module Code"],
                    items["Study Mode"],
                    items["Cohort"],
                    items["Location"],
                    items["Planned Size"],
                    items["Lecturer"],
                    items["Scheduled Date"],
                    items["Scheduled Day"],
                    items["Lecture Start Time"],
                    items["Lecture End Time"],
                    items["Duration"],
                    items["Class Type"]
                ))

    def sort_by_attribute(self, col_id):
        # Sort the data based on the column clicked
        sort_attribute = col_id.lstrip("#")
        self.sort_attribute_var.set(sort_attribute)
        self.display_sorted_data()

    def display_sorted_data(self):
        # Get the selected sorting attribute from the dropdown
        sort_attribute = self.sort_attribute_var.get()

        # Determine whether to sort in ascending or descending order
        sort_order = self.sort_order_var.get()
        reverse = (sort_order == "Descending")

        data_to_sort = self.loaded_data if not self.searched_data else self.searched_data

        # Define a key function to extract the sorting value
        def key_func(item):
            value = item[sort_attribute]

            # Convert date and time strings to datetime objects for specific attributes
            if sort_attribute == 'Scheduled Date':
                return datetime.strptime(value, "%d/%m/%Y")
            elif sort_attribute == 'Scheduled Day':
                days_of_week = ["Monday", "Tuesday", "Wednesday",
                                "Thursday", "Friday", "Saturday", "Sunday"]
                return days_of_week.index(value)
            elif sort_attribute in ['Lecture Start Time', 'Lecture End Time']:
                return datetime.strptime(value, "%H:%M:%S")
            else:
                return value

        try:
            self.timetable_manager.heap_sort(
                data_to_sort, reverse=reverse, key=key_func)
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return

        self.display_results(data_to_sort)

    def display_searched_data(self):
        search_criteria = self.criteria_var.get()
        search_key = self.search_entry.get()

        results = []

        for csv_filename in self.timetable_manager.timetable_manager.data_by_file:
            results.extend(self.timetable_manager.binary_search(
                csv_filename, search_key, search_criteria))

        self.searched_data = results

        self.display_results(results)

    def display_results(self, results):
        self.tree.delete(*self.tree.get_children())

        for result in results:
            items = result.get_items()
            self.tree.insert("", "end", values=(
                items["Module Name"],
                items["Module Code"],
                items["Study Mode"],
                items["Cohort"],
                items["Location"],
                items["Planned Size"],
                items["Lecturer"],
                items["Scheduled Date"],
                items["Scheduled Day"],
                items["Lecture Start Time"],
                items["Lecture End Time"],
                items["Duration"],
                items["Class Type"]
            ))

    def criteria_change(self, *args):
        selected_criteria = self.criteria_var.get()
        search_terms = self.search_term_suggestions.get(selected_criteria, [])

        self.search_entry['values'] = search_terms
        self.search_entry.set("")

    def export_data(self):
        export_format = self.export_options.get()

        if export_format == "Excel (.xlsx)":
            self.export_to_excel()
        elif export_format == "PDF (.pdf)":
            self.export_to_pdf()

    def export_to_excel(self):
        wb = Workbook()
        ws = wb.active

        # Define the headers of excel
        headers = [
            "Module Name",
            "Module Code",
            "Study Mode",
            "Cohort",
            "Location",
            "Planned Size",
            "Lecturer",
            "Scheduled Date",
            "Scheduled Day",
            "Lecture Start Time",
            "Lecture End Time",
            "Duration",
            "Class Type"
        ]

        ws.append(headers)

        data_to_export = self.loaded_data if not self.searched_data else self.searched_data

        # Iterate through the data and add it to the worksheet
        for item in data_to_export:
            values = item.get_items()
            row_data = [
                values["Module Name"],
                values["Module Code"],
                values["Study Mode"],
                values["Cohort"],
                values["Location"],
                values["Planned Size"],
                values["Lecturer"],
                values["Scheduled Date"],
                values["Scheduled Day"],
                values["Lecture Start Time"],
                values["Lecture End Time"],
                values["Duration"],
                values["Class Type"]
            ]
            ws.append(row_data)

        excel_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if excel_file:
            wb.save(excel_file)
            messagebox.showinfo("Export Successful",
                                "Data has been exported to Excel successfully.")

    def export_to_pdf(self):
        data_to_export = self.loaded_data if not self.searched_data else self.searched_data

        pdf_file = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])

        if not pdf_file:
            return

        doc = SimpleDocTemplate(pdf_file, pagesize=landscape(letter))
        story = []

        headers = [
            "Module Name",
            "Module Code",
            "Study Mode",
            "Cohort",
            "Location",
            "Planned Size",
            "Lecturer",
            "Scheduled Date",
            "Scheduled Day",
            "Lecture Start Time",
            "Lecture End Time",
            "Duration",
            "Class Type"
        ]

        col_widths = [
            130, 40, 40, 80, 80, 40, 50, 50, 50, 60, 60, 30, 50
        ]

        data = [headers]

        for item in data_to_export:
            values = item.get_items()
            row_data = [
                values["Module Name"],
                values["Module Code"],
                values["Study Mode"],
                values["Cohort"],
                values["Location"],
                values["Planned Size"],
                values["Lecturer"],
                values["Scheduled Date"],
                values["Scheduled Day"],
                values["Lecture Start Time"],
                values["Lecture End Time"],
                values["Duration"],
                values["Class Type"]
            ]
            data.append(row_data)

        table = Table(data, colWidths=col_widths)

        # Apply table styles
        style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.lightblue),
            ('GRID', (0, 0), (-1, -1), 1, colors.white),
            ('FONTSIZE', (0, 0), (-1, -1), 6)
        ])
        table.setStyle(style)

        story.append(table)

        story.append(PageBreak())

        doc.build(story)
        messagebox.showinfo("Export Successful",
                            "Timetable has been exported to PDF successfully.")


if __name__ == "__main__":
    root = tk.Tk()
    app = Window(root)
    root.mainloop()

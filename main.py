import tkinter as tk
from tkinter import ttk
import openpyxl
import matplotlib.pyplot as plt
import folium
import webbrowser
from geopy.geocoders import Nominatim
from fake_useragent import UserAgent

class LakeDataApp:

    EXCEL_START_ROW = 6
    EXCEL_START_COL = 4
    ELEMENTS_START_COL = 7
    ELEMENTS_END_COL = 13
    EXCEL_FILE_PATH = 'ocena-jakosci-osadow-jeziora-2021_wybrane_pierwisatki.xlsx'

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Zawartość pierwiastków w jeziorze")
        self.root.geometry("600x400")
        self.content_frame = ttk.Frame(self.root)
        self.content_frame.place(relx=0.5, rely=0.4, anchor=tk.CENTER)
        self.lake_names = self.get_lake_names()
        self.selected_lake = tk.StringVar()
        self.element_vars = []
        self.wb = openpyxl.load_workbook(self.EXCEL_FILE_PATH)
        self.sheet = self.wb.active
        self.build()

    def build(self):
        self.build_lake_dropdown()
        self.build_elements_checkbuttons()
        self.build_display_chart_button()
        self.build_display_map_button()

    def get_lake_names(self):
        wb = openpyxl.load_workbook(self.EXCEL_FILE_PATH)
        sheet = wb.active
        lake_names = []

        for row in sheet.iter_rows(min_row=self.EXCEL_START_ROW, min_col=self.EXCEL_START_COL, values_only=True):
            lake_name = row[0]
            if lake_name is not None:
                lake_names.append(lake_name)
        
        return lake_names

    def build_lake_dropdown(self):
        lake_label = ttk.Label(self.content_frame, text="Wybierz jezioro:")
        lake_label.pack(pady=10)
        lake_dropdown = ttk.Combobox(self.content_frame, textvariable=self.selected_lake, values=self.lake_names)
        lake_dropdown.pack()

    def build_elements_checkbuttons(self):
        elements_label = ttk.Label(self.content_frame, text="Wybierz pierwiastki:")
        elements_label.pack(pady=10)
        checkbox_frame = ttk.Frame(self.content_frame)
        checkbox_frame.pack(pady=10)

        for index, row in enumerate(self.sheet.iter_rows(min_row=self.EXCEL_START_ROW-1, max_row=self.EXCEL_START_ROW-1, max_col=self.ELEMENTS_END_COL, values_only=True)):
            for col, element in enumerate(row[self.ELEMENTS_START_COL-1:self.ELEMENTS_END_COL]):
                element_var = tk.BooleanVar()
                element_checkbox = ttk.Checkbutton(checkbox_frame, text=element, variable=element_var)
                element_checkbox.grid(row=index, column=col, padx=10)
                self.element_vars.append(element_var)

    def build_display_chart_button(self):
        display_button = ttk.Button(self.content_frame, text="Wyświetl wykres", command=self.display_chart)
        display_button.pack(pady=10)

    def build_display_map_button(self):
        display_map_button = ttk.Button(self.content_frame, text="Wyświetl mapę", command=self.display_map)
        display_map_button.pack(pady=10)


    def display_chart(self):
        lake_name = self.selected_lake.get()
        element_values, element_labels = self.get_element_values_and_labels(lake_name)
        selected_values, selected_labels = self.get_selected_element_values_and_labels(element_values, element_labels)

        fig, ax = plt.subplots(figsize=(18, 8))
        ax.axis('equal')
        ax.set_title(f'Zawartość pierwiastków w {lake_name} w roku 2021')
        selected_explode = [0.4 if label == 'Pb' else 0 for label in selected_labels]
        wedges, labels = ax.pie(selected_values, labels=selected_values, explode=selected_explode)
        ax.legend(wedges, selected_labels, title="Pierwiastki w mg/kg sm", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
        ax.text(0, -1.5, f'pH: {element_values[0]}', fontsize=12, ha='center')
        plt.show()

    def get_element_values_and_labels(self, lake_name):
        element_values = []
        element_labels = []

        for row in self.sheet.iter_rows(min_row=self.EXCEL_START_ROW, max_col=self.ELEMENTS_END_COL, values_only=True):
            if row[self.EXCEL_START_COL-1] == lake_name:
                element_values = row[self.ELEMENTS_START_COL-1:self.ELEMENTS_END_COL]

        for row in self.sheet.iter_rows(min_row=self.EXCEL_START_ROW-1, max_row=self.EXCEL_START_ROW-1, max_col=self.ELEMENTS_END_COL, values_only=True):
            element_labels = row[self.ELEMENTS_START_COL-1:self.ELEMENTS_END_COL]
        
        return element_values, element_labels

    def get_selected_element_values_and_labels(self, element_values, element_labels):
        selected_values = []
        selected_labels = []

        for i, var in enumerate(self.element_vars):
            if var.get():
                selected_values.append(element_values[i])
                selected_labels.append(element_labels[i])
        
        return selected_values, selected_labels

    def display_map(self):
        lake_name = self.selected_lake.get().split("-")[0].strip()
        location = self.get_location(lake_name)
        
        if location is not None:
            self.show_map(location, lake_name)
        else:
            print("Nie udało się odnaleźć koordynat dla tego jeziora.")

    def get_location(self, lake_name):
        user_agent = UserAgent()
        geolocator = Nominatim(user_agent=user_agent.random)
        location = geolocator.geocode(lake_name)
        
        return location

    def show_map(self, location, lake_name):
        lat, lon = location.latitude, location.longitude
        lake_map = folium.Map(location=[lat, lon], zoom_start=13)

        folium.Marker(
            location=[lat, lon],
            popup=lake_name,
            icon=folium.Icon(color='blue')
        ).add_to(lake_map)

        folium.LatLngPopup().add_to(lake_map)
        lake_map.save('lake_map.html')
        webbrowser.open('lake_map.html')

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = LakeDataApp()
    app.run()

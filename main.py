import matplotlib.pyplot as plt
import openpyxl
import tkinter as tk
from tkinter import ttk
import folium
import webbrowser
from geopy.geocoders import Nominatim
from fake_useragent import UserAgent

# Wczytaj plik Excel
wb = openpyxl.load_workbook('ocena-jakosci-osadow-jeziora-2021_wybrane_pierwisatki.xlsx')
sheet = wb.active

# Utwórz listę do przechowywania nazw jezior
lake_names = []

# Pobierz nazwy jezior z kolumny D i zapisz je w liście nazwy_jezior
for row in sheet.iter_rows(min_row=6, min_col=4, values_only=True):
    lake_name = row[0]
    if lake_name is not None:
        lake_names.append(lake_name)
# Utwórz interfejs użytkownika
root = tk.Tk()
root.title("Zawartość pierwiastków w jeziorze")

# Ustaw rozmiar okna
root.geometry("600x400")

# Utwórz główną ramkę dla zawartości
content_frame = ttk.Frame(root)
content_frame.place(relx=0.5, rely=0.4, anchor=tk.CENTER)

# Utwórz rozwijaną listę do wyboru jeziora
selected_lake = tk.StringVar()
lake_label = ttk.Label(content_frame, text="Wybierz jezioro:")
lake_label.pack(pady=10)
lake_dropdown = ttk.Combobox(content_frame, textvariable=selected_lake, values=lake_names)
lake_dropdown.pack()

# Dodaj etykietę dla wyboru pierwiastków
elements_label = ttk.Label(content_frame, text="Wybierz pierwiastki:")
elements_label.pack(pady=10)

# Utwórz ramkę dla przycisków CheckBox
checkbox_frame = ttk.Frame(content_frame)
checkbox_frame.pack(pady=10)

# Utwórz przyciski CheckBox dla każdego pierwiastka
element_vars = []
for index, row in enumerate(sheet.iter_rows(min_row=5, max_col=13, max_row=5, values_only=True)):
    for col, element in enumerate(row[7:13]):
        element_var = tk.BooleanVar()
        element_checkbox = ttk.Checkbutton(checkbox_frame, text=element, variable=element_var)
        element_checkbox.grid(row=index, column=col, padx=10)
        element_vars.append(element_var)

#Utwórz funkcję wyświetlającą wykres kołowy dla wybranego jeziora
def display_chart():
    # Pobierz nazwę wybranego jeziora
    lake_name = selected_lake.get()
    # Pobierz wartości pierwiastków dla wybranego jeziora
    root_values = []
    root_labels = []
    for row in sheet.iter_rows(min_row=6, max_col=13, values_only=True):
        name = row[3]
        if name == lake_name:
            for cell in row[6:13]:
                root_values.append(cell)

    for row in sheet.iter_rows(min_row=5, max_col=13, max_row=5, values_only=True):
        for cell in row[6:13]:
            root_labels.append(cell)
 
    #Wybierz tylko wybrane pierwiastki
    selected_values = []
    selected_labels = []

    for i, var in enumerate(element_vars):
        if var.get():
            selected_values.append(root_values[i + 1])
            selected_labels.append(root_labels[i + 1])

    #Utwórz wykres kołowy
    fig, ax = plt.subplots(figsize=(18, 8))
    ax.axis('equal')
    ax.set_title(f'Zawartość pierwiastków w {lake_name} w roku 2021')
    selected_explode = [0.4 if etykieta == 'Pb' else 0 for etykieta in selected_labels]
    wedges, labels = ax.pie(selected_values, labels=selected_values, explode=selected_explode)

    #Utwórz legendę
    ax.legend(wedges, selected_labels, title="Pierwiastki w mg/kg sm", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))

    #Label z pH pod wykresem
    ax.text(0, -1.5, f'pH: {root_values[0]}', fontsize=12, ha='center')

    #Wyświetlanie wykresu
    plt.show()

# Utwórz przycisk, aby wyświetlić wykres dla wybranego jeziora
display_button = ttk.Button(content_frame, text="Wyświetl wykres", command=display_chart)
display_button.pack(pady=10)

#Utwórz przycisk, aby wyświetlić mapę dla wybranego jeziora
def display_map():
    # Pobierz nazwę wybranego jeziora
    lake_name = selected_lake.get()

    # Usuń myślniki i wszystko po nich z nazwy jeziora
    lake_name = lake_name.split("-")[0].strip()

    # Użyj geokodera, aby pobrać koordynaty geograficzne dla wybranego jeziora
    user_agent = UserAgent()
    geolocator = Nominatim(user_agent=user_agent.random)
    location = geolocator.geocode(lake_name)
    if location is not None:
        lat, lon = location.latitude, location.longitude
        # Utwórz mapę z wybranym jeziorem
        lake_map = folium.Map(location=[lat, lon], zoom_start=13)

        # Dodaj marker z nazwą jeziora i koordynatami
        folium.Marker(
            location=[lat, lon],
            popup=lake_name,
            icon=folium.Icon(color='blue')
        ).add_to(lake_map)

        # Dodaj interaktywny popup z koordynatami
        folium.LatLngPopup().add_to(lake_map)

        # Wyświetl mapę w przeglądarce
        lake_map.save('lake_map.html')
        webbrowser.open('lake_map.html')
    else:
        print("Nie udało się odnaleźć koordynat dla tego jeziora.")

# Utwórz przycisk, aby wyświetlić mapę dla wybranego jeziora
display_map_button = ttk.Button(content_frame, text="Wyświetl mapę", command=display_map)
display_map_button.pack(pady=10)    

def on_close():
    print("Zamykanie programu...")
    root.destroy()
    
# Obsługa zdarzenia zamknięcia okna
root.protocol("WM_DELETE_WINDOW", on_close)

#Uruchom GUI
root.mainloop()

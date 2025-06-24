import pandas as pd
import os
import tkinter as tk
from tkinter import messagebox
from datetime import datetime, timedelta

def degrees_to_cardinal(degrees):
    """
    Converts wind direction in degrees to cardinal direction (N, NE, E, etc.).
    """
    cardinal_directions = [
        "N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE",
        "S", "SSW", "SW", "WSW", "W", "WNW", "NW", "NNW"
    ]
    idx = int((degrees % 360) / 22.5)
    return cardinal_directions[idx]

def mps_to_kph(speed_mps):
    """
    Converts speed from meters per second (m/s) to kilometers per hour (km/h).
    """
    return round(speed_mps * 3.6)

def prepare_text(location_name, data, next_day, day_after_next):
    """
    Prepares the formatted text output for wind data.
    """
    coastal = data["coastal"]
    offshore = data["offshore"]
    
    text = f"Datos de viento para {location_name}:\n"
    
    # Coastal Zone
    avg_direction = (coastal['wind_direction'].iloc[2] + coastal['wind_direction'].iloc[3]) / 2
    avg_direction_cardinal = degrees_to_cardinal(avg_direction)
    
    text += (
        f"Durante la mañana del {next_day} el viento estará del {degrees_to_cardinal(coastal['wind_direction'].iloc[0])} "
        f"con velocidades entre {mps_to_kph(coastal['wind_speed'].iloc[0])} km/h y {mps_to_kph(coastal['max_wind'])} km/h.\n"
        f"Durante la tarde del {next_day} el viento estará del {degrees_to_cardinal(coastal['wind_direction'].iloc[1])} "
        f"con velocidades entre {mps_to_kph(coastal['wind_speed'].iloc[1])} km/h y {mps_to_kph(coastal['max_wind'])} km/h.\n"
        f"Durante la noche del {next_day} y madrugada del {day_after_next} el viento estará del "
        f"{avg_direction_cardinal} "
        f"con velocidades entre {mps_to_kph(max(coastal['wind_speed'].iloc[2], coastal['wind_speed'].iloc[3]))} km/h y {mps_to_kph(coastal['max_wind'])} km/h.\n\n"
    )
    
    # Offshore Zone
    text += (
        f"Datos de viento para {location_name} mar afuera:\n"
        f"Durante la mañana del {next_day} el viento estará del {degrees_to_cardinal(offshore['wind_direction'].iloc[0])} "
        f"con velocidades entre {mps_to_kph(offshore['wind_speed'].iloc[0])} km/h y {mps_to_kph(offshore['max_wind'])} km/h.\n"
        f"Durante la tarde del {next_day} el viento estará del {degrees_to_cardinal(offshore['wind_direction'].iloc[1])} "
        f"con velocidades entre {mps_to_kph(offshore['wind_speed'].iloc[1])} km/h y {mps_to_kph(offshore['max_wind'])} km/h.\n"
        f"Durante la noche del {next_day} el viento estará del {degrees_to_cardinal(offshore['wind_direction'].iloc[2])} "
        f"con velocidades entre {mps_to_kph(offshore['wind_speed'].iloc[2])} km/h y {mps_to_kph(offshore['max_wind'])} km/h.\n"
        f"Durante la madrugada del {day_after_next} el viento estará del {degrees_to_cardinal(offshore['wind_direction'].iloc[3])} "
        f"con velocidades entre {mps_to_kph(offshore['wind_speed'].iloc[3])} km/h y {mps_to_kph(offshore['max_wind'])} km/h.\n"
    )
    
    return text

def process_wind_data():
    # Get the current date to build the path
    today = datetime.now()
    tomorrow = today + timedelta(days=1)
    day_after_tomorrow = today + timedelta(days=2)
    yy = today.strftime("%y")
    mm = today.strftime("%m")
    dd = today.strftime("%d")
    next_day = tomorrow.strftime("%y/%m/%d")
    day_after_next = day_after_tomorrow.strftime("%y/%m/%d")

    # Define the file path dynamically based on the date
    relative_path = f"/mnt/gfs-wave/06/GFS-Wave_{yy}{mm}{dd}/Oleaje_Viento.xlsx"
    file_path = os.path.abspath(relative_path)

    # Check if the file exists
    if not os.path.exists(file_path):
        print(f"Error: No se encontró el archivo en la ruta {file_path}")
        return

    # Load the Excel file
    excel_data = pd.ExcelFile(file_path)

    # Load the relevant sheets for analysis
    pcoc_data = pd.read_excel(file_path, sheet_name="PCOC", header=None)
    gofo_data = pd.read_excel(file_path, sheet_name="GOFO", header=None)

    # Generate text reports
    output_text = prepare_text("Acajutla", {
        "coastal": {"wind_speed": pcoc_data.loc[8:11, 5], "wind_direction": pcoc_data.loc[8:11, 6], "max_wind": pcoc_data.loc[34, 5]},
        "offshore": {"wind_speed": pcoc_data.loc[8:11, 14], "wind_direction": pcoc_data.loc[8:11, 15], "max_wind": pcoc_data.loc[34, 14]}
    }, next_day, day_after_next)
    output_text += "\n" + prepare_text("La Unión", {
        "coastal": {"wind_speed": gofo_data.loc[8:11, 5], "wind_direction": gofo_data.loc[8:11, 6], "max_wind": gofo_data.loc[34, 5]},
        "offshore": {"wind_speed": gofo_data.loc[8:11, 14], "wind_direction": gofo_data.loc[8:11, 15], "max_wind": gofo_data.loc[34, 14]}
    }, next_day, day_after_next)

    output_file_path = os.path.expanduser("~/Pronostico_Marino/Datos_Viento.txt")
    archive_path = os.path.expanduser(f"~/Pronostico_Marino/archivo/Datos_Viento_{yy}{mm}{tomorrow.strftime('%d')}.txt")

    os.makedirs(os.path.dirname(output_file_path), exist_ok=True)
    os.makedirs(os.path.dirname(archive_path), exist_ok=True)
    
    with open(output_file_path, "w") as file:
        file.write(output_text)
    with open(archive_path, "w") as file:
        file.write(output_text)
    
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("Proceso Completado", "El proceso ha finalizado exitosamente.")

if __name__ == "__main__":
    process_wind_data()

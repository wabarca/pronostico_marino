import pandas as pd
import os
import tkinter as tk
from tkinter import messagebox, simpledialog
from datetime import datetime, timedelta
import math

def degrees_to_cardinal(degrees):
    cardinal_directions = [
        "N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE",
        "S", "SSW", "SW", "WSW", "W", "WNW", "NW", "NNW"
    ]
    idx = int((degrees % 360) / 22.5)
    return cardinal_directions[idx]


def mps_to_kph(speed_mps):
    return math.ceil(speed_mps * 3.6)

def m_to_ft(height_m):
    return math.ceil(height_m * 3.28084)


def viento_texto(v_min, v_max):
    v_min_kph = mps_to_kph(v_min)
    v_max_kph = mps_to_kph(v_max)
    if v_min_kph == v_max_kph:
        return f"con velocidades rondando {v_max_kph} km/h"
    else:
        return f"con velocidades entre {v_min_kph} km/h y {v_max_kph} km/h"


def prepare_text(location_name, coastal_df, offshore_df, next_day, day_after_next):
    text = f"Datos de viento para {location_name}:\n"

    max_wind_coast = coastal_df["SPEWI_C"].max()
    max_wind_offshore = offshore_df["SPEWI_O"].max()

    avg_dir_night = (coastal_df.iloc[2]["DIRWI_C"] + coastal_df.iloc[3]["DIRWI_C"]) / 2

    text += (
        f"Durante la mañana del {next_day}, El viento estará del {degrees_to_cardinal(coastal_df.iloc[0]['DIRWI_C'])} "
        f"{viento_texto(coastal_df.iloc[0]['SPEWI_C'], max_wind_coast)}.\n"
        f"Durante la tarde del {next_day}, El viento estará del {degrees_to_cardinal(coastal_df.iloc[1]['DIRWI_C'])} "
        f"{viento_texto(coastal_df.iloc[1]['SPEWI_C'], max_wind_coast)}.\n"
        f"Durante la noche del {next_day} y madrugada del {day_after_next}, El viento estará del {degrees_to_cardinal(avg_dir_night)} "
        f"{viento_texto(max(coastal_df.iloc[2]['SPEWI_C'], coastal_df.iloc[3]['SPEWI_C']), max_wind_coast)}.\n\n"
    )

    text += f"Datos mar afuera para {location_name}:\n"
    for i, label in enumerate(["mañana", "tarde", "noche", "madrugada"]):
        fecha = next_day if i < 3 else day_after_next
        htsgw = offshore_df.iloc[i]['HTSGW_O']
        text += (
            f"Durante la {label} del {fecha}, El viento estará del {degrees_to_cardinal(offshore_df.iloc[i]['DIRWI_O'])} "
            f"{viento_texto(offshore_df.iloc[i]['SPEWI_O'], max_wind_offshore)}. "
            f"El oleaje estará predominando del {degrees_to_cardinal(offshore_df.iloc[i]['DIRPW_O'])} "
            f"con alturas de {htsgw:.1f} m, aproximadamente equivalente a {m_to_ft(htsgw)} pie.\n"
        )

    return text

def process_wind_data():
    if os.environ.get("DISPLAY", "") == "":
        print("No hay entorno gráfico disponible. Ejecute con entorno X o use un entorno de escritorio.")
        return

    root = tk.Tk()
    root.withdraw()

    user_date_str = simpledialog.askstring("Fecha requerida", "Ingrese la fecha de los datos a utilizar (formato: dd/mm/yyyy):")
    if not user_date_str:
        messagebox.showerror("Error", "No se ingresó ninguna fecha.")
        return

    try:
        base_date = datetime.strptime(user_date_str, "%d/%m/%Y")
    except ValueError:
        messagebox.showerror("Error", "Formato de fecha inválido. Use el formato dd/mm/yyyy.")
        return

    yy = base_date.strftime("%y")
    mm = base_date.strftime("%m")
    dd = base_date.strftime("%d")
    target_date = base_date + timedelta(days=1)
    next_day = target_date.strftime("%Y-%m-%d")
    day_after_next = (target_date + timedelta(days=1)).strftime("%Y-%m-%d")

    relative_path = f"/mnt/gfs-wave/06/GFS-Wave_{yy}{mm}{dd}/Oleaje_Viento.xlsx"
    file_path = os.path.abspath(relative_path)

    if not os.path.exists(file_path):
        messagebox.showerror("Error", f"No se encontró el archivo en la ruta {file_path}")
        return

    def load_filtered_data(sheet):
        df = pd.read_excel(file_path, sheet_name=sheet, header=1)
        df.columns = [
            "FECHA_C", "HORA_C", "DIRPW_C", "HTSGW_C", "PERPW_C", "SPEWI_C", "DIRWI_C",
            "POWPW_C", "Extra_C", "FECHA_O", "HORA_O", "DIRPW_O", "HTSGW_O",
            "PERPW_O", "SPEWI_O", "DIRWI_O", "POWPW_O"
        ]
        year = base_date.year
        df["FECHA_C"] = pd.to_datetime(df["FECHA_C"].astype(str) + f" {year}", format="%d %b %Y", errors='coerce')
        df["FECHA_O"] = pd.to_datetime(df["FECHA_O"].astype(str) + f" {year}", format="%d %b %Y", errors='coerce')
        filtered = df[df["FECHA_C"].dt.date == target_date.date()]
        if filtered.empty:
            raise ValueError(f"No se encontraron datos para la fecha {target_date.strftime('%d/%m/%Y')} en la hoja {sheet}.")
        return filtered

    try:
        pcoc_data = load_filtered_data("PCOC")
        gofo_data = load_filtered_data("GOFO")
    except ValueError as e:
        messagebox.showerror("Error", str(e))
        return

    text_pcoc = prepare_text("Acajutla", pcoc_data, pcoc_data, next_day, day_after_next)
    text_gofo = prepare_text("La Unión", gofo_data, gofo_data, next_day, day_after_next)

    output_text = text_pcoc + "\n" + text_gofo

    output_file_path = os.path.expanduser("~/Pronostico_Marino/Datos_Viento.txt")
    archive_path = os.path.expanduser(f"~/Pronostico_Marino/archivo/Datos_Viento_{yy}{mm}{dd}.txt")

    os.makedirs(os.path.dirname(output_file_path), exist_ok=True)
    os.makedirs(os.path.dirname(archive_path), exist_ok=True)

    with open(output_file_path, "w") as file:
        file.write(output_text)
    with open(archive_path, "w") as file:
        file.write(output_text)

    messagebox.showinfo("Proceso Completado", "El proceso ha finalizado exitosamente.")


if __name__ == "__main__":
    process_wind_data()

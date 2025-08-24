import pandas as pd
import matplotlib.pyplot as plt
import os
from glob import glob
import shutil

def graphe():
    # Dossier d'entrée contenant les fichiers Excel
    input_folder = r".\fusion_resultats"

    # Dossier de sortie
    output_base =  './graphe_par_station'
    if os.path.exists(output_base) and os.path.isdir(output_base):
        shutil.rmtree(output_base)
    os.makedirs(output_base, exist_ok=True)

    # Parcourir tous les fichiers Excel valides
    file_paths = [
        f for f in glob(os.path.join(input_folder, '*.xlsx'))
        if not (os.path.basename(f).startswith('rapp') or os.path.basename(f).startswith('résumé'))
    ]

    for file_path in file_paths:
        try:
            df = pd.read_excel(file_path)
            df['Datetime'] = pd.to_datetime(df[['Year', 'Month', 'Day', 'Hour']])
            station_name = os.path.splitext(os.path.basename(file_path))[0]

            df['TROTOT_mm'] = df['ztd_spotgins']
            df['Diff'] = df['diff_ztd']

            # Créer un dossier spécifique pour chaque station
            station_output = os.path.join(output_base, station_name)
            os.makedirs(station_output, exist_ok=True)

            # ----------- GRAPHE ZTD (GAMIT) ----------
            plt.figure(figsize=(15, 6))
            plt.scatter(df['Datetime'], df['ztd_gamit'], color='blue', s=8, alpha=0.6)
            plt.title(f'ZTD - Gamit ({station_name})')
            plt.xlabel('Temps')
            plt.ylabel('ZTD (mm)')
            plt.grid(True)
            plt.tight_layout()
            plt.savefig(os.path.join(station_output, f'{station_name}_ZTD_GAMIT.png'), dpi=300)
            plt.close()

            # ----------- GRAPHE TROTOT (SPOTGINS) ----------
            plt.figure(figsize=(15, 6))
            plt.scatter(df['Datetime'], df['TROTOT_mm'], color='orange', s=8, alpha=0.6)
            plt.title(f'ZTD - Spotgins ({station_name})')
            plt.xlabel('Temps')
            plt.ylabel('ZTD (mm)')
            plt.grid(True)
            plt.tight_layout()
            plt.savefig(os.path.join(station_output, f'{station_name}_ZTD_SPOTGINS.png'), dpi=300)
            plt.close()

            # ----------- GRAPHE DIFFERENCE ----------
            plt.figure(figsize=(15, 6))
            plt.scatter(df['Datetime'], df['Diff'], color='green', s=8, alpha=0.6)
            plt.title(f'Différence ZTD (Gamit - Spotgins) ({station_name})')
            plt.xlabel('Temps')
            plt.ylabel('Différence (mm)')
            plt.axhline(y=0, color='black', linestyle='--', linewidth=1)
            plt.ylim(-50, 50)
            plt.grid(True)
            plt.tight_layout()
            plt.savefig(os.path.join(station_output, f'{station_name}_DIFFERENCE.png'), dpi=300)
            plt.close()

        except Exception as e:
            print(f"Erreur avec le fichier {file_path} : {e}")

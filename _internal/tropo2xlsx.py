import os
import pandas as pd
from io import StringIO
def tr2xlsx():
    # Chemins des dossiers (à adapter)
    input_folder = "./Spotgins/Spotgins_Tropo"
    output_folder = "./Spotgins/Spotgins_xlsx"

    # Crée le dossier output si nécessaire
    os.makedirs(output_folder, exist_ok=True)

    # Parcours tous les fichiers .tropo dans le dossier input
    for filename in os.listdir(input_folder):
        if filename.endswith(".tropo"):
            file_path = os.path.join(input_folder, filename)

            with open(file_path, 'r') as file:
                lines = file.readlines()

            # Trouve l'en-tête des colonnes
            header_index = None
            for idx, line in enumerate(lines):
                if line.startswith('#jjjjj.jjjjjjjj'):
                    header_index = idx
                    break

            if header_index is not None:
                # Convertit les lignes en DataFrame
                data_lines = lines[header_index:]
                data_str = ''.join(data_lines).replace('#', '')
                df = pd.read_csv(StringIO(data_str), delim_whitespace=True)

                # Écrit le fichier Excel dans le dossier output
                output_filename = filename.replace(".tropo", ".xlsx")
                output_path = os.path.join(output_folder, output_filename)
                df.to_excel(output_path, index=False)
                print(f" {filename} → {output_filename}")
            else:
                print(f" En-tête introuvable dans : {filename}")

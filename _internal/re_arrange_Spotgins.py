import os
import pandas as pd
from io import StringIO

def arrng_spo(datee_start,datee_end):
    import os
    import pandas as pd
    from io import StringIO
    from datetime import datetime

    # Dossiers d'entrée/sortie
    input_folder = "./Spotgins/Spotgins_Tropo"
    output_folder = "./Spotgins/Spotgins_xlsx"
    os.makedirs(output_folder, exist_ok=True)
    print(datee_start,datee_end)
    # Définir la plage de dates à inclure
    date_start = datetime.strptime(datee_start, "%Y_%m_%d")
    date_end = datetime.strptime(datee_end, "%Y_%m_%d")

    for filename in os.listdir(input_folder):
        if filename.endswith(".tropo"):
            file_path = os.path.join(input_folder, filename)

            with open(file_path, 'r') as file:
                lines = file.readlines()

            header_index = None
            for idx, line in enumerate(lines):
                if line.startswith('#jjjjj.jjjjjjjj'):
                    header_index = idx
                    break

            if header_index is not None:
                data_lines = lines[header_index:]
                data_str = ''.join(data_lines).replace('#', '')
                df = pd.read_csv(StringIO(data_str), delim_whitespace=True)

                required_columns = ['yyyymmddHHMMSS', 'TROTOT', 'STDWET']
                if all(col in df.columns for col in required_columns):
                    df['datetime'] = pd.to_datetime(df['yyyymmddHHMMSS'], format='%Y%m%d%H%M%S')

                    # 🔍 Appliquer le filtre de date ici
                    df = df[(df['datetime'] >= date_start) & (df['datetime'] <= date_end)]

                    if not df.empty:
                        df['Year'] = df['datetime'].dt.year
                        df['Month'] = df['datetime'].dt.month
                        df['Day'] = df['datetime'].dt.day
                        df['Hr'] = df['datetime'].dt.hour
                        df['Mn'] = df['datetime'].dt.minute

                        df['TROTOT'] = df['TROTOT'] * 1000
                        df['STDWET'] = df['STDWET'] * 1000

                        export_df = df[['Year', 'Month', 'Day', 'Hr', 'Mn', 'TROTOT', 'STDWET']]

                        output_filename = filename.replace(".tropo", ".xlsx")
                        output_path = os.path.join(output_folder, output_filename)
                        export_df.to_excel(output_path, index=False)

                        print(f"{filename} → {output_filename}")
                    else:
                        print(f"{filename} : aucune donnée dans l’intervalle [{date_start.date()} → {date_end.date()}]")
                else:
                    print(f"{filename} : colonnes manquantes {required_columns}")
            else:
                print(f"En-tête introuvable dans : {filename}")

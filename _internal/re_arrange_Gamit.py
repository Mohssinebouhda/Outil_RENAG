import os
import pandas as pd

# Liste des stations
sitee = [
    "CHTL", "FCLZ", "GINA", "MICH", "MODA", "SAUV", "STJ9", "MTPL", "SJDV", "MANS", "JOUX", "CHIZ", "WLBH", "CHRN", "NICE",
    "SOPH", "RG00", "SVRN", "LROC", "RSTL", "AIGL", "FJCP", "BURE", "CLAP", "STEY", "BANN", "RABU", "CHAM", "TENC",
    "LFAZ", "LEBE", "TROP", "JANU", "PUYA", "ROSD", "LACA", "ALPE", "CLFD", "PLOE", "EOST", "MAKS", "AGDE", "PARD",
    "SLVT", "CHMX", "SETE", "LUCE", "BAUB", "GUIL", "BUAN", "PALI", "MAN2", "STMR", "ARGR", "AVR1", "AUBU", "ERCK",
    "BLIX", "ROTG", "CRAL", "SMTG", "GRJF", "PIMI", "RIXH", "OGAG", "PERX", "FLGY", "FIED", "MRON", "MTP2", "OPME",
    "PDOM", "ARUF", "FAJP", "FILF", "HOLA", "BLVR", "DUNQ", "DIPP", "CHTG", "GERL", "NATT", "SURF", "FJC2", "TLTG",
    "RUSB", "RUSA", "CREF", "CHA2", "AVR2", "LNE1", "IRAF", "CGRO", "GENF", "BOUF"
]
def arranger(site):
    # Répertoires
    input_directory = './Gamit/fichier_met'        # Dossier avec fichiers .met
    output_directory = './Gamit/fichier_xlsx'   # Dossier de sortie
    os.makedirs(output_directory, exist_ok=True)

    # Initialiser les listes de données
    data = {}
    s=site
    data[f"ZTD_list_{s}"] = []
    data[f"sigZTD_list_{s}"] = []
    data[f"Year_{s}"] = []
    data[f"Month_{s}"] = []
    data[f"Day_{s}"] = []
    data[f"Hr_{s}"] = []
    data[f"Mn_{s}"] = []

    # Lire et traiter les fichiers .met
    for filename in os.listdir(input_directory):
        if filename.endswith(".met"):
            try:
                # Charger fichier avec structure sans en-tête
                df = pd.read_csv(
                    os.path.join(input_directory, filename),
                    sep=r'\s+', header=None, skiprows=2
                )

                # Colonnes dynamiques + renommage
                df.columns = [f"Col{i}" for i in range(len(df.columns))]
                df["Site"] = df[df.columns[-2]]  # Avant-dernière colonne = station

                df.rename(columns={
                    "Col0": "Yr",
                    "Col1": "Doy",
                    "Col2": "Hr",
                    "Col3": "Mn",
                    "Col4": "Sec",
                    "Col5": "ZTD",
                    "Col7": "sigZTD"
                }, inplace=True)

                # Extraire Mois et Jour depuis DOY
                df["Month"] = pd.to_datetime(df["Yr"].astype(str) + df["Doy"].astype(str), format="%Y%j").dt.month
                df["Day"] = pd.to_datetime(df["Yr"].astype(str) + df["Doy"].astype(str), format="%Y%j").dt.day

                # Remplir les listes pour chaque station
                station=site
                df_station = df[df["Site"] == station]
                for _, row in df_station.iterrows():
                    data[f"ZTD_list_{station}"].append(row["ZTD"])
                    data[f"sigZTD_list_{station}"].append(row["sigZTD"])
                    data[f"Year_{station}"].append(int(row["Yr"]))
                    data[f"Month_{station}"].append(int(row["Month"]))
                    data[f"Day_{station}"].append(int(row["Day"]))
                    data[f"Hr_{station}"].append(int(row["Hr"]))
                    data[f"Mn_{station}"].append(int(row["Mn"]))

                print(f" {filename} traité.")
            except Exception as e:
                print(f" Erreur avec {filename}: {e}")

    # Exporter un Excel par station
    station=site
    station_df = pd.DataFrame({
        "Year": data[f"Year_{station}"],
        "Month": data[f"Month_{station}"],
        "Day": data[f"Day_{station}"],
        "Hr": data[f"Hr_{station}"],
        "Mn": data[f"Mn_{station}"],
        "ZTD": data[f"ZTD_list_{station}"],
        "sigZTD": data[f"sigZTD_list_{station}"]
    })

    if not station_df.empty:
        output_path = os.path.join(output_directory, f"{station}.xlsx")
        station_df.to_excel(output_path, index=False)

    print(f" les fichiers Excel de la station {site} est généré avec Hr et Mn.")

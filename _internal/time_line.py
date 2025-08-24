
print("**************************************")

folder_gamit = r'.\Gamit\Timetamp'  # Remplacez par le chemin de votre dossier de sortie
folder_spotgins = r'.\Spotgins\Timetamp'

import pandas as pd
import os

def timeline_gamit():
# Fonction pour traiter chaque fichier Excel
    def process_excel_file(input_file_path, output_file_path):
        # Charger le fichier Excel
        data = pd.read_excel(input_file_path)

        data = data.rename(columns={"Year": "year", "Month": "month", "Day": "day", "Hr": "hour"})
        # Créer une colonne datetime en combinant Year, Month, Day et Hour
        data['datetime'] = pd.to_datetime(data[['year', 'month', 'day', 'hour']])

        # Calculer la différence en minutes entre les lignes consécutives
        data['time_diff'] = data['datetime'].diff().dt.total_seconds() / 60  # Convertir en minutes

        # Identifier les lignes où la différence de temps est supérieure à 60 minutes
        data['new_series'] = data['time_diff'] > 60

        # Liste pour stocker les créneaux de début, fin, et le nombre de lignes
        sub_series_start_end_with_count = []

        # Initialisation de l'index de départ
        start_idx = 0

        # Parcours des données pour identifier les sous-séries
        for i in range(1, len(data)):
            if data['new_series'].iloc[i]:  # Si une nouvelle série commence
                sub_series_start_end_with_count.append(
                    (data['datetime'].iloc[start_idx], data['datetime'].iloc[i-1], i - start_idx)
                )
                start_idx = i

        # Ajouter la dernière sous-série
        sub_series_start_end_with_count.append(
            (data['datetime'].iloc[start_idx], data['datetime'].iloc[-1], len(data) - start_idx)
        )

        # Convertir la liste en DataFrame pour l'affichage et l'exportation
        sub_series_df = pd.DataFrame(sub_series_start_end_with_count, columns=['Start', 'End', 'Row_Count'])

        # Sauvegarder le DataFrame dans un fichier Excel
        sub_series_df.to_excel(output_file_path, index=False)

        print(f"Le fichier Excel a été sauvegardé sous : {output_file_path}")

    # Dossier contenant les fichiers Excel
    input_folder = r'C:\Users\Hp\Desktop\pythonProject\code assemble\Gamit\fichier_xlsx'  # Remplacez par le chemin de votre dossier

    # Vérifier que le dossier de sortie existe, sinon le créer
    if not os.path.exists(folder_gamit):
        os.makedirs(folder_gamit)

    # Parcourir tous les fichiers Excel dans le dossier
    for file_name in os.listdir(input_folder):
        if file_name.endswith('.xlsx') or file_name.endswith('.xls'):
            input_file_path = os.path.join(input_folder, file_name)
            output_file_path = os.path.join(folder_gamit, f"{file_name}")

            # Traiter chaque fichier
            process_excel_file(input_file_path, output_file_path)

    print("Traitement terminé pour tous les fichiers.")

#spotgins :


print("**************************************")


import pandas as pd
import os



def timeline_spotgins():
    # Fonction pour traiter chaque fichier Excel
    def process_excel_file(input_file_path, output_file_path):
        data = pd.read_excel(input_file_path)
        data = data.rename(columns={"Year": "year", "Month": "month", "Day": "day", "Hr": "hour"})

        # Créer la colonne datetime
        data['datetime'] = pd.to_datetime(data[['year', 'month', 'day', 'hour']])
        data['time_diff'] = data['datetime'].diff().dt.total_seconds() / 60
        data['new_series'] = data['time_diff'] > 60

        sub_series = []
        start_idx = 0

        for i in range(1, len(data)):
            if data['new_series'].iloc[i]:
                start = data['datetime'].iloc[start_idx]
                end = data['datetime'].iloc[i-1]
                if True:
                    sub_series.append((start, end, i - start_idx))
                start_idx = i

        # Dernière série
        start = data['datetime'].iloc[start_idx]
        end = data['datetime'].iloc[-1]
        if True:
            sub_series.append((start, end, len(data) - start_idx))

        # Sauvegarde
        sub_series_df = pd.DataFrame(sub_series, columns=['Start', 'End', 'Row_Count'])
        sub_series_df.to_excel(output_file_path, index=False)
        print(f" Fichier sauvegardé : {output_file_path}")

    # Dossiers d'entrée et de sortie
    input_folder = r'.\Spotgins\Spotgins_xlsx'

    os.makedirs(folder_spotgins, exist_ok=True)

    # Parcours des fichiers
    for file_name in os.listdir(input_folder):
        if file_name.endswith(('.xlsx', '.xls')):
            input_file_path = os.path.join(input_folder, file_name)
            output_file_path = os.path.join(folder_spotgins, f"{file_name[:4]}.xlsx")
            process_excel_file(input_file_path, output_file_path)

    print(" Traitement terminé pour tous les fichiers.")


def plot():
    timeline_gamit()
    timeline_spotgins()
    import pandas as pd
    import matplotlib.pyplot as plt
    import matplotlib.dates as mdates
    import os
    import matplotlib.patches as mpatches

    def plot_common_timelines(folder_gamit, folder_spogins, output_html):
        files_gamit = {os.path.splitext(f)[0]: os.path.join(folder_gamit, f)
                       for f in os.listdir(folder_gamit) if f.endswith('.xlsx')}
        files_spogins = {os.path.splitext(f)[0]: os.path.join(folder_spogins, f)
                         for f in os.listdir(folder_spogins) if f.endswith('.xlsx')}

        common_names = sorted(set(files_gamit.keys()).intersection(set(files_spogins.keys())))

        plt.figure(figsize=(14, max(5, len(common_names) * 0.4)))
        color_gam = "tab:blue"
        color_spo = "tab:orange"

        all_dates = []

        for idx, name in enumerate(common_names):
            # GAMIT
            df_gam = pd.read_excel(files_gamit[name])
            start_gam = pd.to_datetime(df_gam.iloc[:, 0], errors='coerce')
            end_gam = pd.to_datetime(df_gam.iloc[:, 1], errors='coerce')
            all_dates.extend(start_gam.dropna())
            all_dates.extend(end_gam.dropna())
            for s, e in zip(start_gam, end_gam):
                if pd.notnull(s) and pd.notnull(e):
                    plt.plot([s, e], [idx - 0.02, idx - 0.02], linewidth=5, color=color_gam)

            # SPOTGINS
            df_spo = pd.read_excel(files_spogins[name])
            start_spo = pd.to_datetime(df_spo.iloc[:, 0], errors='coerce')
            end_spo = pd.to_datetime(df_spo.iloc[:, 1], errors='coerce')
            all_dates.extend(start_spo.dropna())
            all_dates.extend(end_spo.dropna())
            for s, e in zip(start_spo, end_spo):
                if pd.notnull(s) and pd.notnull(e):
                    plt.plot([s, e], [idx + 0.02, idx + 0.02], linewidth=5, color=color_spo)

        plt.yticks(range(len(common_names)), common_names)
        plt.xlabel('Date')
        plt.title('Timeline des fichiers communs SPOTGINS & GAMIT')

        # Légende personnalisée
        patch_gam = mpatches.Patch(color=color_gam, label='GAMIT')
        patch_spo = mpatches.Patch(color=color_spo, label='SPOTGINS')
        plt.legend(handles=[patch_gam, patch_spo], loc='upper right')

        ax = plt.gca()

        #  Amélioration de l'échelle de temps sur l'axe X
        locator = mdates.AutoDateLocator()
        formatter = mdates.ConciseDateFormatter(locator)
        ax.xaxis.set_major_locator(locator)
        ax.xaxis.set_major_formatter(formatter)

        # Ajout de marges autour des dates pour ne pas couper les barres
        if all_dates:
            min_date = min(all_dates) - pd.Timedelta(days=30)
            max_date = max(all_dates) + pd.Timedelta(days=30)
            ax.set_xlim(min_date, max_date)

        plt.xticks(rotation=45)
        plt.grid(axis='x')
        plt.tight_layout()

        # Sauvegarde image
        image_path = os.path.splitext(output_html)[0] + '_timeline_new.png'
        plt.savefig(image_path, dpi=150, bbox_inches='tight')
        plt.close()

        # Génération HTML
        html_content = f"""
        <html>
        <head><title>Timeline avec légende</title></head>
        <body>
            <h1>Timeline des fichiers communs</h1>
            <img src="{os.path.basename(image_path)}" style="max-width: none;">
        </body>
        </html>
        """
        with open(output_html, 'w', encoding='utf-8') as f:
            f.write(html_content)

        print(f"Graphique avec légende sauvegardé dans : {output_html}")

    # Appel de la fonction (à personnaliser)
    plot_common_timelines(folder_gamit, folder_spotgins, output_html='timeline_communs_legende_new.html')

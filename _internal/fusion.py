import os
import pandas as pd
import numpy as np
import shutil

def fusion():
    folder1 = r".\Gamit\fichier_xlsx"
    folder2 = r".\Spotgins\Spotgins_xlsx"
    output_folder = r".\fusion_resultats"

    if os.path.exists(output_folder):
        shutil.rmtree(output_folder)
    os.makedirs(output_folder, exist_ok=True)

    files_gamit = [f for f in os.listdir(folder1) if f.endswith(".xlsx")]
    stats_list = []

    for file_gamit in files_gamit:
        prefix = file_gamit[:4].upper()
        file_path_gamit = os.path.join(folder1, file_gamit)
        matching_files_spot = [f for f in os.listdir(folder2) if f.upper().startswith(prefix) and f.endswith(".xlsx")]
        if not matching_files_spot:
            continue

        file_path_spot = os.path.join(folder2, matching_files_spot[0])

        try:
            df_gamit = pd.read_excel(file_path_gamit)
            df_spotgins = pd.read_excel(file_path_spot)

            df_gamit.columns = df_gamit.columns.str.lower()
            df_spotgins.columns = df_spotgins.columns.str.lower()

            df_gamit['datetime'] = pd.to_datetime(dict(year=df_gamit['year'], month=df_gamit['month'],
                                                       day=df_gamit['day'], hour=df_gamit['hr'], minute=df_gamit['mn']))
            df_spotgins['datetime'] = pd.to_datetime(dict(year=df_spotgins['year'], month=df_spotgins['month'],
                                                          day=df_spotgins['day'], hour=df_spotgins['hr'],
                                                          minute=df_spotgins['mn']))

            merged = pd.merge(df_gamit, df_spotgins, on="datetime", how="inner")

            merged['Year'] = merged['datetime'].dt.year
            merged['Month'] = merged['datetime'].dt.month
            merged['Day'] = merged['datetime'].dt.day
            merged['Hour'] = merged['datetime'].dt.hour

            final_df = merged[['Year', 'Month', 'Day', 'Hour',
                               'ztd', 'sigztd', 'trotot', 'stdwet']].copy()

            final_df = final_df.rename(columns={
                'ztd': 'ztd_gamit',
                'sigztd': 'sigztd_gamit',
                'trotot': 'ztd_spotgins',
                'stdwet': 'sigztd_spotgins'
            })

            final_df['diff_ztd'] = final_df['ztd_gamit'] - final_df['ztd_spotgins']
            final_df['diff_sigztd'] = final_df['sigztd_gamit'] - final_df['sigztd_spotgins']

            output_file = os.path.join(output_folder, f"fusion_{prefix}.xlsx")
            final_df.to_excel(output_file, index=False)

            # Statistiques
            diff_ztd = final_df['diff_ztd']
            diff_sigztd = final_df['diff_sigztd']
            stats_list.append({
                'Station': prefix,
                'BIAS_ZTD': diff_ztd.mean(),
                'MAE_ZTD': diff_ztd.abs().mean(),
                'RMS_ZTD': np.sqrt((diff_ztd ** 2).mean()),
                'BIAS_sigZTD': diff_sigztd.mean(),
                'MAE_sigZTD': diff_sigztd.abs().mean(),
                'RMS_sigZTD': np.sqrt((diff_sigztd ** 2).mean())
            })

        except Exception:
            pass

    if stats_list:
        df_stats = pd.DataFrame(stats_list)
        stats_txt = os.path.join(output_folder, "resume_stats.txt")

        with open(stats_txt, 'w') as f:
            f.write("Résumé statistique par station :\n\n")
            for stat in stats_list:
                f.write(f"{stat['Station']} :\n")
                f.write(f"  BIAS ZTD        = {stat['BIAS_ZTD']:.3f} mm\n")
                f.write(f"  MAE  ZTD        = {stat['MAE_ZTD']:.3f} mm\n")
                f.write(f"  RMS  ZTD        = {stat['RMS_ZTD']:.3f} mm\n")
                f.write(f"  BIAS sigZTD     = {stat['BIAS_sigZTD']:.3f} mm\n")
                f.write(f"  MAE  sigZTD     = {stat['MAE_sigZTD']:.3f} mm\n")
                f.write(f"  RMS  sigZTD     = {stat['RMS_sigZTD']:.3f} mm\n\n")

            f.write("Caractéristiques globales (moyennes et écarts-types) :\n\n")
            for col in ['BIAS_ZTD', 'MAE_ZTD', 'RMS_ZTD', 'BIAS_sigZTD', 'MAE_sigZTD', 'RMS_sigZTD']:
                mean = df_stats[col].mean()
                std = df_stats[col].std()
                f.write(f"  {col} : Moyenne = {mean:.3f} mm , Écart-type = {std:.3f} mm\n")

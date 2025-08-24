def val_aber():
    import os
    import shutil
    import pandas as pd

    def process_outliers(folder1, folder2, output_excel, output_txt, strategy_name):
        if os.path.exists(folder2):
            shutil.rmtree(folder2)
        os.makedirs(folder2)

        filtered_rows = []
        outlier_counts = {}
        cleaned_stats = {}

        for file_name in os.listdir(folder1):
            if file_name.endswith(('.xlsx', '.xls')):
                file_path = os.path.join(folder1, file_name)
                try:
                    df = pd.read_excel(file_path)
                    station = file_name[:4]
                    outlier_counts.setdefault(station, 0)

                    if {'ZTD', 'sigZTD'}.issubset(df.columns):
                        # Condition "ET" : sigZTD < 100 et ZTD entre 500 et 3000
                        valid = (df['sigZTD'] < 100) & (df['ZTD'] < 3000) & (df['ZTD'] > 500)
                        filtered = df[valid].copy()
                        cleaned = df[~valid].copy()

                        # Enregistrement des lignes valides
                        if not filtered.empty:
                            filtered.insert(0, 'Station', station)
                            filtered_rows.append(filtered)

                        # Compter les outliers
                        outlier_counts[station] += len(cleaned)

                        # Statistiques sur données d'origine
                        if not df.empty:
                            mean_ztd = df['ZTD'].mean()
                            std_ztd = df['ZTD'].std()
                            mean_sig = df['sigZTD'].mean()
                            std_sig = df['sigZTD'].std()
                            cleaned_stats[station] = {
                                'mean_ZTD': mean_ztd,
                                'std_ZTD': std_ztd,
                                'mean_sigZTD': mean_sig,
                                'std_sigZTD': std_sig
                            }
                        else:
                            cleaned_stats[station] = {
                                'mean_ZTD': None,
                                'std_ZTD': None,
                                'mean_sigZTD': None,
                                'std_sigZTD': None
                            }
                    elif {'TROTOT', 'STDWET'}.issubset(df.columns):
                        df['sigZTD']=df['STDWET']
                        df['ZTD']=df['TROTOT']
                        # Condition "ET" : sigZTD < 100 et ZTD entre 500 et 3000
                        valid = (df['sigZTD'] < 100) & (df['ZTD'] < 3000) & (df['ZTD'] > 500)
                        filtered = df[valid].copy()
                        cleaned = df[~valid].copy()

                        # Enregistrement des lignes valides
                        if not filtered.empty:
                            filtered.insert(0, 'Station', station)
                            filtered_rows.append(filtered)

                        # Compter les outliers
                        outlier_counts[station] += len(cleaned)

                        # Statistiques sur données d'origine
                        if not df.empty:
                            mean_ztd = df['ZTD'].mean()
                            std_ztd = df['ZTD'].std()
                            mean_sig = df['sigZTD'].mean()
                            std_sig = df['sigZTD'].std()
                            cleaned_stats[station] = {
                                'mean_ZTD': mean_ztd,
                                'std_ZTD': std_ztd,
                                'mean_sigZTD': mean_sig,
                                'std_sigZTD': std_sig
                            }
                        else:
                            cleaned_stats[station] = {
                                'mean_ZTD': None,
                                'std_ZTD': None,
                                'mean_sigZTD': None,
                                'std_sigZTD': None
                            }
                except Exception as e:
                    print(f"Erreur dans le fichier {file_name} : {e}")

        # Création DataFrame final
        if filtered_rows:
            result_df = pd.concat(filtered_rows, ignore_index=True)
        else:
            result_df = pd.DataFrame(columns=['Station', 'Year', 'Month', 'Day', 'Hr', 'Mn', 'ZTD', 'sigZTD'])

        result_path = os.path.join(folder2, output_excel)
        result_df.to_excel(result_path, index=False)

        # Fichier .txt
        txt_path = os.path.join(folder2, output_txt)
        with open(txt_path, 'w') as f:
            # Bloc 1 : Outliers
            f.write(f"Analyse des outliers pour stratégie {strategy_name} :\n\n")
            for station in sorted(outlier_counts):
                f.write(f"{station} : {outlier_counts[station]} outliers\n")

            # Bloc 2 : Stats sur données d’origine
            f.write("\nStatistiques des fichiers d'origine (toutes lignes, même avec outliers) :\n\n")
            for station in sorted(cleaned_stats):
                stats = cleaned_stats[station]
                f.write(f"{station} :\n")
                if stats['mean_ZTD'] is not None:
                    f.write(f"   Moyenne ZTD     = {stats['mean_ZTD']:.2f} mm\n")
                    f.write(f"   Écart-type ZTD  = {stats['std_ZTD']:.2f} mm\n")
                    f.write(f"   Moyenne sigZTD  = {stats['mean_sigZTD']:.2f} mm\n")
                    f.write(f"   Écart-type sigZTD = {stats['std_sigZTD']:.2f} mm\n")
                else:
                    f.write("   Pas de données valides.\n")
                f.write("\n")

            # Bloc 3 : Stats sur les données valides (fichier filtré)
            f.write("Statistiques sur les données valides (sans outliers) :\n\n")
            if filtered_rows:
                df_all = pd.concat(filtered_rows, ignore_index=True)
                for station in sorted(df_all['Station'].unique()):
                    df_station = df_all[df_all['Station'] == station]
                    mean_ztd = df_station['ZTD'].mean()
                    std_ztd = df_station['ZTD'].std()
                    mean_sig = df_station['sigZTD'].mean()
                    std_sig = df_station['sigZTD'].std()

                    f.write(f"{station} :\n")
                    f.write(f"   Moyenne ZTD     = {mean_ztd:.2f} mm\n")
                    f.write(f"   Écart-type ZTD  = {std_ztd:.2f} mm\n")
                    f.write(f"   Moyenne sigZTD  = {mean_sig:.2f} mm\n")
                    f.write(f"   Écart-type sigZTD = {std_sig:.2f} mm\n\n")
            else:
                f.write("   Aucune donnée valide après filtrage.\n")

        print(f"Résultats enregistrés dans : {folder2}\n")

    # === GAMIT ===
    process_outliers(
        folder1=r'.\Gamit\fichier_xlsx',
        folder2=r'.\Gamit\Sans_outlier',
        output_excel='outlier_Gamit.xlsx',
        output_txt='analyse_outliers_Gamit.txt',
        strategy_name='GAMIT'
    )

    # === SPOTGINS ===
    process_outliers(
        folder1=r'.\Spotgins\Spotgins_xlsx',
        folder2=r'.\Spotgins\Sans_outlier',
        output_excel='outlier_Spotgins.xlsx',
        output_txt='analyse_outliers_Spotgins.txt',
        strategy_name='SPOTGINS'
    )

#!/usr/bin/env python3# -*- coding: utf-8 -*-"""Created on Wed Dec  4 01:13:05 2024@author: noahwuhrmann"""import pandas as pdimport streamlit as stimport re# Streamlit Appst.title("Excel Sortier- und Exporttool")# Dateiuploaduploaded_file = st.file_uploader("Bitte eine Excel-Datei hochladen", type=["xlsx"])if uploaded_file:    # Mitteilung über erfolgreichen Upload    st.success("Die Datei wurde erfolgreich hochgeladen.")        # Datei in einen DataFrame laden    try:        df = pd.read_excel(uploaded_file, engine='openpyxl')                # Spaltennamen definieren        fixed_columns = ["Bestand", "Artikelname", "KatalogNr", "Artikel", "Netto"]                # Dynamische Spalten identifizieren        current_year = pd.Timestamp.now().year        last_year = current_year - 1                # Jahres-Spalten        year_columns = [str(current_year), str(last_year)]                # Monatsspalten erkennen (z. B. "Jan2024", "Dez2023")        month_pattern = re.compile(r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(\d{4})")        month_columns = [col for col in df.columns if month_pattern.match(col)]                # Monatsspalten sortieren (zuerst aktuelles Jahr, dann rückwärts)        month_columns.sort(key=lambda x: (int(x[-4:]), x[:3]), reverse=True)                # Zielspalten in der richtigen Reihenfolge        target_columns = fixed_columns + year_columns + month_columns                # Spalten im neuen DataFrame beibehalten        filtered_df = df[[col for col in target_columns if col in df.columns]]                # Nutzer auffordern, den Dateinamen anzugeben        output_filename = st.text_input("Bitte geben Sie einen Namen für die exportierte Datei ein (ohne .xlsx):", value="Exportierte_Datei")                if st.button("Excel-Datei exportieren"):            # Excel-Datei speichern            with pd.ExcelWriter(f"{output_filename}.xlsx", engine='openpyxl') as writer:                filtered_df.to_excel(writer, index=False)            st.success(f"Die Datei '{output_filename}.xlsx' wurde erfolgreich erstellt und ist bereit zum Herunterladen.")            st.download_button(                label="Herunterladen",                data=filtered_df.to_excel(index=False, engine='openpyxl'),                file_name=f"{output_filename}.xlsx",                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"            )    except Exception as e:        st.error(f"Ein Fehler ist aufgetreten: {e}")
#!/usr/bin/env python3# -*- coding: utf-8 -*-"""Created on Wed Dec  4 01:13:05 2024@author: noahwuhrmann"""# Verison 2import pandas as pdimport streamlit as stimport refrom io import BytesIO# Streamlit Appst.title("Excel Sortier- und Exporttool")# Dateiuploaduploaded_file = st.file_uploader("Bitte eine Excel-Datei hochladen", type=["xlsx"])if uploaded_file:    # Mitteilung über erfolgreichen Upload    st.success("Die Datei wurde erfolgreich hochgeladen.")        # Datei in einen DataFrame laden    try:        df = pd.read_excel(uploaded_file, engine='openpyxl')                # Spaltennamen definieren        fixed_columns = ["Bestand", "Artikelname", "Artikelgruppe", "KatalogNr", "Artikel", "Netto"]                # Dynamische Spalten identifizieren        current_year = pd.Timestamp.now().year        last_year = current_year - 1                year_columns = [str(current_year), str(last_year)]        month_columns = []                # Monatsspalten erkennen (z. B. "Jan 2023", "Dez 2022")        month_pattern = re.compile(r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s(\d{4})")        for col in df.columns:            if month_pattern.match(col):                month_columns.append(col)                # Monatsspalten sortieren (zuerst aktuelles Jahr, dann rückwärts)        month_columns.sort(key=lambda x: (int(x.split()[-1]), x[:3]), reverse=True)                # Zielspalten in der richtigen Reihenfolge        target_columns = fixed_columns + year_columns + month_columns                # Spalten im neuen DataFrame beibehalten        filtered_df = df[[col for col in target_columns if col in df.columns]]                # Nutzer auffordern, den Dateinamen anzugeben        output_filename = st.text_input("Bitte geben Sie einen Namen für die exportierte Datei ein (ohne .xlsx):", value="Exportierte_Datei")                if st.button("Excel-Datei exportieren"):            # Excel-Datei speichern            output = BytesIO()            with pd.ExcelWriter(output, engine='openpyxl') as writer:                filtered_df.to_excel(writer, index=False, sheet_name="Export")            st.success(f"Die Datei '{output_filename}.xlsx' wurde erfolgreich erstellt und ist bereit zum Herunterladen.")            st.download_button(                label="Herunterladen",                data=output.getvalue(),                file_name=f"{output_filename}.xlsx",                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"            )    except Exception as e:        st.error(f"Ein Fehler ist aufgetreten: {e}")
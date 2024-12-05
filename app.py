#!/usr/bin/env python3# -*- coding: utf-8 -*-"""Created on Thu Dec 5 2024@author: Noah Wuhrmann"""#v13import pandas as pdimport streamlit as stfrom openpyxl import load_workbookfrom openpyxl.styles import Font, Alignment, PatternFillfrom openpyxl.utils import get_column_letter# Benutzeroberflächest.title("Einkaufsanalyse Taurus Sports AG")st.image("absolut_bild.jpg", caption="Absolute Teamsport – Taurus Sports AG", use_column_width=True)# Dateiuploaduploaded_file = st.file_uploader("Bitte eine Excel-Datei hochladen", type=["xlsx"])if uploaded_file:    # Mitteilung über erfolgreichen Upload    st.success("Die Datei wurde erfolgreich hochgeladen.")        # Datei in einen DataFrame laden    try:        df = pd.read_excel(uploaded_file, engine="openpyxl")                # Fixe Spalten definieren        fixed_columns = ["Bestand", "Artikelname", "Artikelgruppe", "KatalogNr", "Artikel"]        # Spalten mit Jahreszahlen erkennen        year_pattern = r"^\d{4}$"        year_columns = [col for col in df.columns if pd.Series(col).str.match(year_pattern).any()]        # Monatsspalten erkennen (Deutsch)        month_pattern = r"^(Jan|Feb|Mär|Apr|Mai|Jun|Jul|Aug|Sep|Okt|Nov|Dez)\s\d{4}$"        month_columns = [col for col in df.columns if pd.Series(col).str.match(month_pattern).any()]        # Spalten sortieren: Fixe Spalten, Jahreszahlen, Monate        target_columns = fixed_columns + year_columns + month_columns        # Spalte "Netto" ans Ende verschieben, falls vorhanden        if "Netto" in df.columns:            target_columns.append("Netto")        target_columns = [col for col in target_columns if col in df.columns]        # Nicht relevante Zeilen filtern        relevant_df = df[target_columns].copy()        relevant_df = relevant_df[~((relevant_df["Bestand"] == 0) & (relevant_df[month_columns].sum(axis=1) == 0))]        # Exportierte Datei erstellen        output_filename = st.text_input("Bitte geben Sie einen Namen für die exportierte Datei ein (ohne .xlsx):", value="Exportierte_Datei")                if st.button("Excel-Datei exportieren"):            # DataFrame in eine Excel-Datei schreiben            with pd.ExcelWriter(f"{output_filename}.xlsx", engine="openpyxl") as writer:                relevant_df.to_excel(writer, index=False, sheet_name="Export")                # Workbook und Worksheet laden                workbook = writer.book                worksheet = writer.sheets["Export"]                # Schriftart und Schriftgröße für das gesamte Dokument setzen                for row in worksheet.iter_rows():                    for cell in row:                        cell.font = Font(name="Calibri", size=11)                        cell.alignment = Alignment(horizontal="right", vertical="center")                # Header formatieren                for cell in worksheet[1]:                    cell.font = Font(name="Calibri", size=11, bold=True)                    cell.alignment = Alignment(horizontal="center", vertical="center")                    cell.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")                # Spaltenausrichtung anpassen                for col in ["Artikelname", "Artikelgruppe", "KatalogNr", "Artikel"]:                    if col in relevant_df.columns:                        col_idx = relevant_df.columns.get_loc(col) + 1                        for cell in worksheet[get_column_letter(col_idx)]:                            cell.alignment = Alignment(horizontal="left", vertical="center")                # Spalte "Netto" formatieren                if "Netto" in relevant_df.columns:                    netto_col_idx = relevant_df.columns.get_loc("Netto") + 1                    for cell in worksheet[get_column_letter(netto_col_idx)][1:]:                        cell.number_format = "#,##0.00"                        if cell.value and cell.value < 0:                            cell.font = Font(color="FF0000")  # Rot für negative Werte                # Autofilter aktivieren                worksheet.auto_filter.ref = worksheet.dimensions                # Spaltenbreiten anpassen                for col_idx, column_cells in enumerate(worksheet.columns, start=1):                    max_length = max((len(str(cell.value)) for cell in column_cells if cell.value), default=0)                    adjusted_width = max_length + 2                    worksheet.column_dimensions[get_column_letter(col_idx)].width = adjusted_width            st.success(f"Die Datei '{output_filename}.xlsx' wurde erfolgreich erstellt und ist bereit zum Herunterladen.")            with open(f"{output_filename}.xlsx", "rb") as file:                st.download_button(                    label="Herunterladen",                    data=file,                    file_name=f"{output_filename}.xlsx",                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"                )    except Exception as e:        st.error(f"Ein Fehler ist aufgetreten: {e}")
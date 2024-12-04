#!/usr/bin/env python3# -*- coding: utf-8 -*-"""Created on Wed Dec  4 01:13:05 2024@author: noahwuhrmann"""''' Das wird das Programm für die Einkaufsanalyse für die Frima Taurus Sports AG'''import streamlit as stimport pandas as pd# Streamlit App Titelst.title("Einkauf Analyse - Excel Checker")# Datei-Uploaduploaded_file = st.file_uploader("Lade eine Excel-Datei hoch, um zu überprüfen, ob sie gefüllt oder leer ist:", type=["xlsx"])if uploaded_file:    try:        # Laden der Excel-Datei        df = pd.read_excel(uploaded_file)                # Überprüfen, ob die Datei leer ist        if df.empty:            st.warning("Die hochgeladene Datei ist leer.")        else:            st.success("Die hochgeladene Datei ist gefüllt.")                        # Anzeigen der Spaltennamen            st.write("Die Spaltennamen in der Datei sind:")            st.write(df.columns.tolist())    except Exception as e:        st.error(f"Fehler beim Verarbeiten der Datei: {e}")else:    st.info("Bitte lade eine Excel-Datei hoch.")
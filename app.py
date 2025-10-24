import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path
import tempfile
import io
from heat_analysis import run_pipeline


# ----------------------------
# Streamlit-Konfiguration
# ----------------------------
st.set_page_config(page_title="Heat Load Validator", page_icon="🔥", layout="wide")

# Custom CSS für besseres Design
st.markdown("""
<style>
    .main-header {
        font-size: 3rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        text-align: center;
        color: #666;
        margin-bottom: 2rem;
    }
    .metric-card {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        text-align: center;
    }
    .stDownloadButton button {
        width: 100%;
        background-color: #28a745;
        color: white;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<p class="main-header">🔥 Heat Load Validation Tool</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Analyse von Normheizlast-Berichten nach SIA 384-2</p>', unsafe_allow_html=True)

# ----------------------------
# Sidebar mit Info
# ----------------------------
with st.sidebar:
    st.header("ℹ️ Über dieses Tool")
    st.markdown("""
    **Heat Load Validation Tool** analysiert PDF-Berichte im SIA 384-2 Format und:
    
    ✅ Extrahiert Raumdaten automatisch  
    ✅ Kategorisiert Räume nach 7 Kategorien  
    ✅ Validiert gegen Benchmarks  
    ✅ Erstellt detaillierte Excel-Berichte  
    
    ---
    
    **Kategorien:**
    - Big Room (>25 m²)
    - Corner Room
    - Exposed Room
    - Internal Room
    - Small Room w/t Outside Wall
    - Wetcell
    - Wetcell Mit Aussenbezug
    
    ---
    
    **Excel-Output enthält:**
    - Rooms (alle Räume mit Validierung)
    - Summary (Zusammenfassung)
    - ProjectStats (Projekt-Statistiken)
    - All-Projects Benchmarks
    - Custom Intervals
    """)
    
    st.markdown("---")
    st.markdown("**Version:** 2.0")
    st.markdown("**Entwickelt mit:** Streamlit & pdfplumber")

# ----------------------------
# File Upload
# ----------------------------
st.markdown("### 📁 PDF hochladen")
uploaded_pdf = st.file_uploader(
    "Wähle eine SIA 384-2 PDF-Datei aus",
    type="pdf",
    help="Lade eine PDF-Datei mit Normheizlast-Berechnungen hoch"
)

if uploaded_pdf:
    col1, col2 = st.columns([3, 1])
    with col1:
        st.success(f"✅ Datei **{uploaded_pdf.name}** erfolgreich hochgeladen ({uploaded_pdf.size / 1024:.1f} KB)")
    with col2:
        if st.button("🔄 Neue Datei", use_container_width=True):
            st.rerun()

    st.markdown("---")

    # ----------------------------
    # Analyse starten
    # ----------------------------
    if st.button("▶️ Analyse starten", type="primary", use_container_width=True):
        with st.spinner("🔄 PDF wird verarbeitet... Dies kann einige Minuten dauern."):
            # Temporäre Datei erstellen
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                tmp_pdf.write(uploaded_pdf.getvalue())
                tmp_pdf_path = tmp_pdf.name

            df_results, out_xlsx = None, None
            try:
                df_results, out_xlsx = run_pipeline(tmp_pdf_path)
                # Ergebnisse in Session State speichern
                st.session_state['df_results'] = df_results
                st.session_state['out_xlsx'] = out_xlsx
                st.session_state['uploaded_filename'] = uploaded_pdf.name
            except Exception as e:
                st.error(f"❌ Fehler bei der Analyse: {e}")
                st.exception(e)
                st.stop()

    # ----------------------------
    # Ergebnisse anzeigen (aus Session State)
    # ----------------------------
    if 'df_results' in st.session_state and st.session_state['df_results'] is not None:
        df_results = st.session_state['df_results']
        out_xlsx = st.session_state['out_xlsx']
        
        if df_results is not None and not df_results.empty:
            st.success("✅ **Analyse erfolgreich abgeschlossen!**")
            
            # ----------------------------
            # KPI-Übersicht
            # ----------------------------
            st.markdown("### 📊 Projekt-Kennzahlen")
            
            col1, col2, col3, col4, col5 = st.columns(5)
            
            total_rooms = len(df_results)
            avg_heat = df_results['required_heat_per_m2'].mean()
            min_heat = df_results['required_heat_per_m2'].min()
            max_heat = df_results['required_heat_per_m2'].max()
            total_area = df_results['room_area'].sum()
            
            with col1:
                st.metric("Anzahl Räume", total_rooms)
            with col2:
                st.metric("Ø Heizlast", f"{avg_heat:.1f} W/m²")
            with col3:
                st.metric("Min Heizlast", f"{min_heat:.1f} W/m²")
            with col4:
                st.metric("Max Heizlast", f"{max_heat:.1f} W/m²")
            with col5:
                st.metric("Gesamt Fläche", f"{total_area:.1f} m²")

            st.markdown("---")

            # ----------------------------
            # Kategorieverteilung
            # ----------------------------
            st.markdown("### 📈 Kategorieverteilung")
            
            col1, col2 = st.columns([1, 1])
            
            with col1:
                # Balkendiagramm
                category_counts = df_results['category'].value_counts()
                fig1, ax1 = plt.subplots(figsize=(8, 5))
                category_counts.plot(kind='barh', ax=ax1, color='#1f77b4')
                ax1.set_xlabel('Anzahl Räume')
                ax1.set_ylabel('Kategorie')
                ax1.set_title('Räume pro Kategorie')
                plt.tight_layout()
                st.pyplot(fig1)
            
            with col2:
                # Pie Chart
                fig2, ax2 = plt.subplots(figsize=(8, 5))
                category_counts.plot(kind='pie', ax=ax2, autopct='%1.1f%%', startangle=90)
                ax2.set_ylabel('')
                ax2.set_title('Verteilung nach Kategorien')
                plt.tight_layout()
                st.pyplot(fig2)

            st.markdown("---")

            # ----------------------------
            # Heizlast-Analyse
            # ----------------------------
            st.markdown("### 🌡️ Heizlast-Analyse")
            
            col1, col2 = st.columns([1, 1])
            
            with col1:
                # Box-Plot pro Kategorie
                fig3, ax3 = plt.subplots(figsize=(10, 6))
                df_results.boxplot(column='required_heat_per_m2', by='category', ax=ax3)
                ax3.set_xlabel('Kategorie')
                ax3.set_ylabel('Spezifische Heizlast [W/m²]')
                ax3.set_title('Heizlast-Verteilung pro Kategorie')
                plt.xticks(rotation=45, ha='right')
                plt.suptitle('')
                plt.tight_layout()
                st.pyplot(fig3)
            
            with col2:
                # Violin Plot
                fig4, ax4 = plt.subplots(figsize=(10, 6))
                sns.violinplot(data=df_results, y='category', x='required_heat_per_m2', ax=ax4)
                ax4.set_xlabel('Spezifische Heizlast [W/m²]')
                ax4.set_ylabel('Kategorie')
                ax4.set_title('Heizlast-Verteilung (Violin Plot)')
                plt.tight_layout()
                st.pyplot(fig4)

            st.markdown("---")

            # ----------------------------
            # Detaillierte Tabelle
            # ----------------------------
            st.markdown("### 📋 Detaillierte Raumdaten")
            
            # Filter-Optionen
            col1, col2, col3 = st.columns(3)
            
            with col1:
                selected_categories = st.multiselect(
                    "Filter nach Kategorie",
                    options=sorted(df_results['category'].unique()),
                    default=list(df_results['category'].unique()),
                    key="category_filter"
                )
            
            with col2:
                selected_floors = st.multiselect(
                    "Filter nach Geschoss",
                    options=sorted(df_results['geschoss'].unique()),
                    default=list(sorted(df_results['geschoss'].unique())),
                    key="floor_filter"
                )
            
            with col3:
                min_heat = float(df_results['required_heat_per_m2'].min())
                max_heat = float(df_results['required_heat_per_m2'].max())
                
                heat_range = st.slider(
                    "Heizlast-Bereich [W/m²]",
                    min_value=min_heat,
                    max_value=max_heat,
                    value=(min_heat, max_heat),
                    key="heat_range_filter"
                )
            
            # Gefilterte Daten
            filtered_df = df_results[
                (df_results['category'].isin(selected_categories)) &
                (df_results['geschoss'].isin(selected_floors)) &
                (df_results['required_heat_per_m2'] >= heat_range[0]) &
                (df_results['required_heat_per_m2'] <= heat_range[1])
            ]
            
            # Anzeigeoptionen für Tabelle
            display_columns = [
                'geschoss', 'room_code', 'room_name', 'category',
                'room_area', 'raumtemperatur', 'normheizlast', 'required_heat_per_m2'
            ]
            
            # Spalten umbenennen für bessere Lesbarkeit
            display_df = filtered_df[display_columns].copy()
            display_df.columns = [
                'Geschoss', 'Code', 'Raumname', 'Kategorie',
                'Fläche [m²]', 'Temp [°C]', 'Normheizlast [W]', 'Spez. Heizlast [W/m²]'
            ]
            
            # Formatierung
            display_df['Fläche [m²]'] = display_df['Fläche [m²]'].round(1)
            display_df['Normheizlast [W]'] = display_df['Normheizlast [W]'].round(0)
            display_df['Spez. Heizlast [W/m²]'] = display_df['Spez. Heizlast [W/m²]'].round(1)
            
            st.dataframe(
                display_df,
                use_container_width=True,
                height=400
            )
            
            st.info(f"📊 Zeige {len(filtered_df)} von {len(df_results)} Räumen")

            st.markdown("---")

            # ----------------------------
            # Statistik-Tabelle
            # ----------------------------
            st.markdown("### 📊 Statistik pro Kategorie")
            
            stats_df = df_results.groupby('category').agg({
                'room_area': ['count', 'sum', 'mean'],
                'required_heat_per_m2': ['mean', 'min', 'max', 'std']
            }).round(1)
            
            stats_df.columns = [
                'Anzahl', 'Gesamt Fläche [m²]', 'Ø Fläche [m²]',
                'Ø Heizlast [W/m²]', 'Min [W/m²]', 'Max [W/m²]', 'StdAbw [W/m²]'
            ]
            
            st.dataframe(stats_df, use_container_width=True)

            st.markdown("---")

            # ----------------------------
            # Download-Bereich
            # ----------------------------
            st.markdown("### 💾 Excel-Bericht herunterladen")
            
            if out_xlsx and Path(out_xlsx).exists():
                with open(out_xlsx, "rb") as f:
                    xlsx_bytes = f.read()
                
                col1, col2, col3 = st.columns([2, 3, 2])
                with col2:
                    st.download_button(
                        label="📥 Excel-Bericht herunterladen",
                        data=xlsx_bytes,
                        file_name=Path(st.session_state['uploaded_filename']).stem + "_validation_results.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                
                st.success("""
                ✅ **Der Excel-Bericht enthält:**
                - **Rooms:** Alle Räume mit Validierungsstatus (Accepted/Low/High)
                - **Summary:** Zusammenfassung nach Kategorien
                - **ProjectStats:** Statistische Kennwerte für dieses Projekt
                - **All-Projects Benchmarks:** Referenzwerte aus allen Projekten
                - **Custom Intervals:** Möglichkeit für manuelle Grenzwert-Anpassungen
                """)
            else:
                st.error("⚠️ Excel-Datei konnte nicht erstellt werden.")
        
        else:
            st.warning("⚠️ Keine gültigen Daten im PDF gefunden. Bitte überprüfe das Format.")

else:
    # ----------------------------
    # Startseite ohne Upload
    # ----------------------------
    st.info("👆 Bitte lade eine PDF-Datei hoch, um die Analyse zu starten.")
    
    st.markdown("---")
    
    st.markdown("### 🚀 Schnellstart")
    st.markdown("""
    1. **PDF hochladen:** Wähle eine SIA 384-2 PDF-Datei aus
    2. **Analyse starten:** Klicke auf den Button "Analyse starten"
    3. **Ergebnisse prüfen:** Betrachte die Visualisierungen und Statistiken
    4. **Excel herunterladen:** Lade den detaillierten Bericht herunter
    """)
    
    st.markdown("### 🎯 Features")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("""
        **📊 Automatische Extraktion**
        - Raumcode & Name
        - Fläche & Geschoss
        - Normheizlast
        - Raumtemperatur
        - Bauteilinformationen
        """)
    
    with col2:
        st.markdown("""
        **🏷️ Intelligente Kategorisierung**
        - 7 vordefinierte Kategorien
        - Regelbasierte Klassifikation
        - Berücksichtigung von Orientierung
        - Erkennung von Nasszellen
        """)
    
    with col3:
        st.markdown("""
        **✅ Validierung & Benchmarking**
        - Vergleich mit Referenzwerten
        - Projekt-spezifische Statistik
        - Farbcodierte Bewertung
        - Manuelle Anpassungen möglich
        """)
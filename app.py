import streamlit as st
import pandas as pd
import io
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from pypdf import PdfReader, PdfWriter

st.set_page_config(page_title="Excel to PDF Template Mapper", layout="wide")

st.title("📄 Vpisování Excel dat do existujícího PDF")
st.write("Nahrajte své stávající PDF (leták) a Excel soubor s daty a umístěte texty na přesné pozice.")

# Inicializace stavu pro mapování pozic (sloupec: {x, y, size})
if "pdf_mapping" not in st.session_state:
    # Přednastavená výchozí pozice pro "Kód hry" na základě letáku Babylon na šířku
    st.session_state.pdf_mapping = {
        "Kód hry": {"x": 168, "y": 170, "size": 14}
    }

col1, col2 = st.columns([1, 1])

df = None
columns = []
template_pdf = None

with col1:
    st.subheader("1. Vstupní soubory")
    
    # Nahrání existujícího PDF
    uploaded_pdf = st.file_uploader("Nahrajte leták jako PDF šablonu (.pdf)", type=["pdf"])
    if uploaded_pdf is not None:
        template_pdf = uploaded_pdf
        st.success("PDF leták úspěšně nahrán.")
        
    # Nahrání Excelu
    uploaded_file = st.file_uploader("Nahrajte Excel soubor s daty (.xlsx, .xls)", type=["xlsx", "xls"])
    if uploaded_file is not None:
        try:
            xl = pd.ExcelFile(uploaded_file)
            sheet = st.selectbox("Vyberte list:", xl.sheet_names)
            df = xl.parse(sheet, header=None, dtype=str)
            df.columns = [f"Sloupec {i+1}" for i in range(df.shape[1])]
            columns = list(df.columns)
            st.success(f"Načteno {len(columns)} sloupců a {len(df)} řádků z Excelu.")
        except Exception as e:
            st.error(f"Chyba při čtení Excelu: {e}")

with col2:
    st.subheader("2. Nastavení pozic polí")
    if df is not None and template_pdf is not None:
        st.info("Zadejte souřadnice v bodech. Počátek [0,0] je vlevo dole. Výchozí pozice pro 'Kód hry' je přednastavena pro formát na šířku tak, aby odpovídala bílému políčku na letáku Babylon.")
        
        # Formulář pro přidání/úpravu souřadnic pro sloupec
        col_sel, col_x, col_y, col_size = st.columns([2, 1, 1, 1])
        
        with col_sel:
            selected_col = st.selectbox("Vyberte sloupec z Excelu:", options=columns, key="pdf_col_sel")
        with col_x:
            # Výchozí hodnota X=168 odpovídá pozici bílého obdélníku za "Kód hry" na šířku A4
            x_pos = st.number_input("Souřadnice X:", min_value=0, max_value=900, value=168, step=5)
        with col_y:
            # Výchozí hodnota Y=175 posouvá text přesně do spodního bílého políčka na šířku A4
            y_pos = st.number_input("Souřadnice Y:", min_value=0, max_value=600, value=175, step=5)
        with col_size:
            font_size = st.number_input("Velikost písma:", min_value=6, max_value=72, value=14, step=1)
            
        if st.button("Uložit pozici pro sloupec", use_container_width=True):
            st.session_state.pdf_mapping[selected_col] = {"x": x_pos, "y": y_pos, "size": font_size}
            st.toast(f"Uloženo pro {selected_col}: [{x_pos}, {y_pos}]")

        # Zobrazení aktuálního nastavení
        if st.session_state.pdf_mapping:
            st.write("🔍 **Aktuálně nastavené pozice:**")
            mapping_list = []
            for col, settings in st.session_state.pdf_mapping.items():
                mapping_list.append({
                    "Sloupec z Excelu": col,
                    "X (Vodorovně)": settings["x"],
                    "Y (Svisle)": settings["y"],
                    "Velikost písma": settings["size"]
                })
            st.table(mapping_list)
            
            if st.button("Vymazat všechna pravidla"):
                st.session_state.pdf_mapping = {}
                st.rerun()
    else:
        st.info("Pro konfiguraci mapování musíte nejprve nahrát PDF leták i Excel soubor.")

# Generování výstupu
if df is not None and template_pdf is not None and st.session_state.pdf_mapping:
    st.markdown("---")
    st.subheader("3. Vygenerovat popsané PDF")
    
    st.info("Aplikace automaticky spáruje dva po sobě jdoucí řádky z Excelu na jeden leták (první kód vlevo, druhý kód vpravo). Pokud máte lichý počet kódů, na posledním letáku zůstane pravá strana prázdná. Počítá se s orientací šablony na šířku.")
    
    if st.button("Přenést data do PDF šablony", type="primary"):
        try:
            template_bytes = template_pdf.read()
            output_writer = PdfWriter()
            
            # Procházíme tabulku po dvojicích řádků
            for i in range(0, len(df), 2):
                row_left = df.iloc[i]
                # Kontrola, zda existuje i druhý řádek pro pravou stranu
                row_right = df.iloc[i+1] if (i + 1) < len(df) else None
                
                overlay_buffer = io.BytesIO()
                # Výběr orientace na šířku (A4 landscape) pro zachování proporcí 842x595
                c = canvas.Canvas(overlay_buffer, pagesize=landscape(A4))
                
                for excel_col, config in st.session_state.pdf_mapping.items():
                    # Levá strana (1. kód ze dvojice) - použijeme zadané souřadnice X a Y
                    if excel_col in row_left:
                        val_left = str(row_left[excel_col]) if pd.notna(row_left[excel_col]) else ""
                        c.setFont("Helvetica-Bold", config["size"])
                        c.setFillColorRGB(0.1, 0.1, 0.1)
                        # Umístění na levou stranu
                        c.drawString(config["x"], config["y"], val_left)
                    
                    # Pravá strana (2. kód ze dvojice) - posuneme X souřadnici o 421 bodů doprava (přesná polovina A4 na šířku)
                    if row_right is not None and excel_col in row_right:
                        val_right = str(row_right[excel_col]) if pd.notna(row_right[excel_col]) else ""
                        c.setFont("Helvetica-Bold", config["size"])
                        c.setFillColorRGB(0.1, 0.1, 0.1)
                        # Posun o šířku poloviny A4 stránky na šířku (přesně o 421 bodů doprava)
                        c.drawString(config["x"] + 421, config["y"], val_right)
                
                c.save()
                overlay_bytes = overlay_buffer.getvalue()

                reader_template = PdfReader(io.BytesIO(template_bytes))
                reader_overlay = PdfReader(io.BytesIO(overlay_bytes))

                # 1. nejdřív přidat stránku šablony do writeru
                output_writer.add_page(reader_template.pages[0])
                # 2. potom na ni sloučit overlay
                output_writer.pages[-1].merge_page(reader_overlay.pages[0])

            final_pdf_buffer = io.BytesIO()
            output_writer.write(final_pdf_buffer)
            final_pdf_buffer.seek(0)

            
            st.success("Hotovo! Kódy byly úspěšně vepsány do obou částí letáku.")
            st.download_button(
                label="📥 Stáhnout vyplněné letáky (PDF)",
                data=final_pdf_buffer,
                file_name="vyplnene_letaky.pdf",
                mime="application/pdf"
            )
            
        except Exception as e:
            st.error(f"Chyba při generování PDF: {e}")
elif df is not None and template_pdf is not None:
    st.info("Před stažením musíte nastavit pozici alespoň pro jeden sloupec z Excelu.")

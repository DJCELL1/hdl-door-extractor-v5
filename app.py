import os, tempfile, streamlit as st, pandas as pd
import HDL_Door_Schedule_Extractor_v5 as extractor
from pathlib import Path

st.set_page_config(page_title='HDL Door Schedule Extractor v5', page_icon='🚪', layout='wide')
st.title('🚪 HDL Door Schedule Extractor v5')
st.write('Convert architectural door schedules from PDF to Excel.')

supplier = st.sidebar.selectbox('Supplier', ['ARA', 'ALLEGION', 'ASSA', 'DORMAKABA', 'JK', 'GENERIC'], index=5)
force_ocr = st.sidebar.checkbox('Force OCR', value=False)
files = st.file_uploader('Upload PDF files', type=['pdf'], accept_multiple_files=True)

def extract_to_df(file_bytes, supplier, force_ocr):
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
        tmp.write(file_bytes); path = tmp.name
    ext, parser = extractor.PDFExtractor(path), extractor.ScheduleParser(supplier)
    pages = ext.ocr_pages() if force_ocr else ext.text_pages()
    lines = parser.parse_lines(pages)
    data = [[l.area, l.door, l.code, l.description, l.colour, l.quantity] for l in lines]
    os.remove(path); return pd.DataFrame(data, columns=['Area','Door','Code','Description','Colour','Quantity'])

if files:
    for f in files:
        st.subheader(f'📄 {f.name}')
        try:
            df = extract_to_df(f.read(), supplier, force_ocr)
            if df.empty: st.warning('No data found.'); continue
            st.dataframe(df, use_container_width=True)
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmpx:
                extractor.write_excel([extractor.ExtractedLine(*r, source='app') for r in df.values.tolist()], tmpx.name)
                st.download_button('⬇️ Download Excel', data=open(tmpx.name,'rb').read(), file_name=f'{Path(f.name).stem}_extracted.xlsx')
                os.remove(tmpx.name)
        except Exception as e: st.error(f'Error: {e}')
else:
    st.info('👆 Upload PDF files to begin.')

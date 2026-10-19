import streamlit as st
from formatter import format_document

st.set_page_config(page_title='ITARS Smart Paper Formatter',page_icon='📄',layout='centered')
st.title('ITARS Smart Paper Formatter')
st.caption('International TVET Academic and Research Symposium (ITARS 6.0)')
st.info('Muat naik manuskrip Microsoft Word (.docx) untuk pemformatan asas dan laporan pematuhan berdasarkan format ITARS 5.0 yang digunakan semula untuk ITARS 6.0.')
st.warning('Versi percubaan: Pemformatan automatik tidak menjamin pematuhan penuh. Semak abstrak, rajah, jadual, persamaan dan rujukan APA 7 sebelum menghantar.')
with st.expander('Privasi dan had pemprosesan'):
    st.write('Fail diproses dalam sesi aplikasi untuk menghasilkan output dan tidak disimpan sengaja ke pangkalan data oleh kod ini. Perkhidmatan hosting mungkin mempunyai log atau dasar penyimpanan tersendiri. Elakkan memuat naik manuskrip sulit sehingga penganjur meluluskan hosting dan dasar privasi.')
st.download_button('Muat turun template asal (rujukan format)',data=open('ITARS_Original_Template.docx','rb').read(),file_name='ITARS_6_Reference_Template.docx',mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
uploaded=st.file_uploader('Pilih manuskrip Word',type=['docx'],help='Maksimum 20 MB')
if uploaded:
    if st.button('Format Manuskrip',type='primary'):
        try:
            out,report=format_document(uploaded.getvalue(),label='ITARS 6.0',year='2027')
            st.session_state['formatted']=out
            st.session_state['report']=report
            st.success('Pemformatan asas selesai. Sila periksa laporan sebelum menghantar manuskrip.')
        except Exception as e:st.error(f'Gagal memproses dokumen: {e}')
if 'formatted' in st.session_state:
    st.download_button('⬇️ Muat turun Word yang diformat',st.session_state['formatted'],file_name='ITARS6_Formatted_Paper.docx',mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    st.subheader('Laporan pematuhan')
    report=st.session_state['report']
    for c in report['checks']:
        st.write(('✅' if c['Status']=='Lulus' else '⚠️')+f" **{c['Semakan']}** — {c['Status']}")
        st.caption(c['Butiran'])
    lines=['ITARS 6.0 — LAPORAN PEMATUHAN','Nota: Pemeriksaan automatik terhad; bukan pengesahan penerimaan kertas.','']
    lines += [f"{c['Status']} | {c['Semakan']} | {c['Butiran']}" for c in report['checks']]
    st.download_button('⬇️ Muat turun laporan TXT','\n'.join(lines).encode(),file_name='ITARS6_Compliance_Report.txt',mime='text/plain')
st.divider()
st.caption('Aplikasi pemformatan berasingan daripada portal submission; muat turun dan semak Word sebelum penghantaran rasmi.')

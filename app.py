import json
from pathlib import Path
import streamlit as st
from engine import validate_profile, read_csv, normalize, process, excel_bytes

st.set_page_config(page_title='Perfilador de cursos', page_icon='🎯', layout='wide')
st.title('Perfilador de cursos')
st.caption('Convierte tu base de contactos en una lista priorizada para cada curso.')
with st.sidebar:
    st.header('1 · Perfil del curso')
    profiles = list((Path(__file__).parent / 'profiles').glob('*.json'))
    choice = st.selectbox('Perfil guardado', profiles, format_func=lambda p: p.stem)
    custom = st.file_uploader('O carga otro perfil JSON', type=['json'])
    try:
        profile = validate_profile(json.loads(custom.getvalue()) if custom else json.loads(choice.read_text()))
    except Exception as exc:
        st.error(f'Perfil inválido: {exc}')
        st.stop()
    st.success(profile['nombre_curso'])
    st.caption(f"Código: {profile['codigo_curso']} · Creado: {profile.get('fecha_creacion', '—')}")
    st.download_button('Descargar perfil JSON', json.dumps(profile, ensure_ascii=False, indent=2), f"{profile['perfil_id']}.json", 'application/json')
    with st.expander('Ver palabras clave y criterios'):
        st.json(profile)
    exclusions = bool(profile.get('keywords_excluir'))
    if exclusions:
        st.info('Las exclusiones se buscan en Cargo; los motivos que empiezan por SST también se buscan en Profesión. Una coincidencia excluye al contacto aunque tenga puntaje alto.')
    apply_exclusions = st.checkbox('Aplicar exclusiones del perfil', value=exclusions, disabled=not exclusions)

st.header('2 · Carga tu base')
upload = st.file_uploader('Archivo CSV con cargo, profesión y empresa; correo opcional', type=['csv'])
if not upload:
    st.info('Selecciona el perfil del curso y carga un CSV para empezar.')
    st.stop()
try:
    df = read_csv(upload.getvalue())
    if df.empty:
        raise ValueError('El archivo no contiene registros.')
except Exception as exc:
    st.error(f'No se pudo cargar la base: {exc}')
    st.stop()
st.caption(f'{len(df):,} registros cargados')
with st.expander('Vista previa de la base'):
    st.dataframe(df.head(20), hide_index=True)
columns = {}
for container, dim, aliases in zip(st.columns(4), ['Cargo','Profesion','Empresa','Email'], [('CARGO',),('PROFESION',),('EMPRESA',),('EMAIL','CORREO','CORREO ELECTRONICO','E MAIL','MAIL')]):
    options = ['No usar'] + list(df.columns) if dim == 'Email' else list(df.columns)
    default = next((col for col in df.columns if normalize(col) in aliases), options[0])
    selected = container.selectbox('Correo (opcional)' if dim == 'Email' else 'Profesión' if dim == 'Profesion' else dim, options, index=options.index(default))
    columns[dim] = None if dim == 'Email' and selected == 'No usar' else selected
signature = json.dumps([profile, columns, apply_exclusions], sort_keys=True) + str(upload.getvalue())
if st.button('Priorizar prospectos', type='primary', use_container_width=True):
    if len({columns[d] for d in ['Cargo','Profesion','Empresa']}) != 3:
        st.error('Selecciona una columna distinta para cargo, profesión y empresa.')
    else:
        try:
            with st.spinner('Calculando prioridades…'):
                result, duplicates = process(df, profile, columns, apply_exclusions)
                workbook = excel_bytes(result, columns['Empresa'])
                st.session_state['result'] = (signature, result, duplicates, workbook)
        except Exception as exc:
            st.error(f'No se pudo procesar la base: {exc}')
saved = st.session_state.get('result')
if saved and saved[0] == signature:
    _, result, duplicates, workbook = saved
    st.header('3 · Resultados')
    st.caption(f'{len(result):,} contactos únicos · {duplicates:,} duplicados por correo eliminados. Se conserva el de mayor puntaje; las filas sin correo se mantienen.')
    for container, seg in zip(st.columns(5), ['A','B','C','D','EXCLUIDO']):
        container.metric('Excluidos' if seg == 'EXCLUIDO' else f'Segmento {seg}', int((result.Segmento == seg).sum()))
    st.caption('A: 70 o más · B: 45–69 · C: 20–44 · D: menos de 20. Los puntajes expresan afinidad según las reglas del perfil.')
    segments = st.multiselect('Filtrar segmentos', ['A','B','C','D','EXCLUIDO'], default=['A','B','C'])
    search = st.text_input('Buscar en los resultados', placeholder='Empresa, cargo, nombre…')
    visible = result[result.Segmento.isin(segments)]
    if search:
        visible = visible[visible.astype(str).apply(lambda s: s.str.contains(search, case=False, regex=False)).any(axis=1)]
    st.dataframe(visible, hide_index=True, use_container_width=True)
    st.download_button('Descargar Excel completo', workbook, f"{profile['perfil_id']}_Resultados.xlsx", 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', type='primary')
    st.caption('El Excel incluye Base Priorizada (A–C), Resumen, Excluidos y Base Completa con Score. Los filtros de pantalla no limitan la descarga.')

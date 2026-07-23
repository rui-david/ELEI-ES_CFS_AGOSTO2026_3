import streamlit as st
import sqlite3
import pandas as pd
from datetime import datetime
import io
import os

st.set_page_config(page_title="Eleições CFS 2026", layout="centered")
ARQUIVO_SOCIOS = "dados.xlsx"
ARQUIVO_LOGO = "logo_cfs.png" # Coloca o logo com esse nome na pasta
DOC_ADMIN = "123548"

# CABEÇALHO COM LOGO
col1, col2 = st.columns([1, 4])
with col1:
    if os.path.exists(ARQUIVO_LOGO):
        st.image(ARQUIVO_LOGO, width=100)
    else:
        st.warning("Logo não encontrado")
with col2:
    st.markdown("<h1>Eleições CFS 2026</h1>", unsafe_allow_html=True)
    st.markdown("<h4>Clube Futebol os Sanjoanenses</h4>", unsafe_allow_html=True)

conn = sqlite3.connect("dados.db", check_same_thread=False, timeout=10)
cursor = conn.cursor()

cursor.execute("CREATE TABLE IF NOT EXISTS config (chave TEXT PRIMARY KEY, valor TEXT)")
cursor.execute("CREATE TABLE IF NOT EXISTS socios (numero_documento TEXT PRIMARY KEY, nome TEXT, votou INTEGER DEFAULT 0)")
cursor.execute("CREATE TABLE IF NOT EXISTS votos (numero_documento TEXT, voto TEXT, timestamp TEXT)")
conn.commit()

def carregar_socios_xlsx():
    if not os.path.exists(ARQUIVO_SOCIOS):
        st.error(f"Arquivo {ARQUIVO_SOCIOS} não encontrado!")
        return pd.DataFrame()
    df = pd.read_excel(ARQUIVO_SOCIOS)
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
    df.columns = df.columns.str.strip().str.upper()
    mapeamento = {'NOME': 'nome', 'NÚMERO_DOCUMENTO': 'numero_documento', 'NUMERO_DOCUMENTO': 'numero_documento'}
    df = df.rename(columns=mapeamento)
    df = df.drop_duplicates(subset=['numero_documento'], keep='first')
    df = df[df['numero_documento'].astype(str).str.strip()!= '']
    df['numero_documento'] = df['numero_documento'].astype(str).str.strip()
    df['nome'] = df['nome'].astype(str).str.strip()
    return df

def sincronizar_socios():
    df = carregar_socios_xlsx()
    if df.empty: return False

    novos = 0
    for _, row in df.iterrows():
        cursor.execute("INSERT OR IGNORE INTO socios (numero_documento, nome) VALUES (?,?)",
                       (row['numero_documento'], row['nome']))
        if cursor.rowcount > 0: novos += 1
    conn.commit()
    st.success(f"Sincronizado! {novos} sócios novos adicionados. Total: {len(df)}")
    return True

def get_config(chave, default):
    cursor.execute("SELECT valor FROM config WHERE chave=?", (chave,))
    res = cursor.fetchone()
    return res[0] if res else default

def set_config(chave, valor):
    cursor.execute("INSERT OR REPLACE INTO config VALUES (?,?)", (chave, valor))
    conn.commit()

def get_datas():
    inicio_str = get_config("inicio", "2026-08-01 08:00")
    fim_str = get_config("fim", "2026-08-02 18:00")
    return datetime.strptime(inicio_str, "%Y-%m-%d %H:%M"), datetime.strptime(fim_str, "%Y-%m-%d %H:%M")

def buscar_socio(doc):
    cursor.execute("SELECT * FROM socios WHERE numero_documento=?", (doc,))
    return cursor.fetchone()

def ja_votou(doc):
    socio = buscar_socio(doc)
    return socio and socio[2] == 1

def votar(doc, voto):
    cursor.execute("INSERT INTO votos VALUES (?,?,?)", (doc, voto, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    cursor.execute("UPDATE socios SET votou = 1 WHERE numero_documento =?", (doc,))
    conn.commit()

def get_resultados():
    df = pd.read_sql_query("SELECT voto, COUNT(*) as total FROM votos GROUP BY voto", conn)
    total_votos = pd.read_sql_query("SELECT COUNT(*) as total FROM votos", conn).iloc[0]['total']
    total_socios = pd.read_sql_query("SELECT COUNT(*) as total FROM socios", conn).iloc[0]['total']
    return df, total_votos, total_socios

def resetar_votacao():
    cursor.execute("DELETE FROM votos")
    cursor.execute("UPDATE socios SET votou = 0")
    conn.commit()

def exportar_excel():
    df_votos = pd.read_sql_query("SELECT * FROM votos", conn)
    df_socios = pd.read_sql_query("SELECT * FROM socios", conn)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_votos.to_excel(writer, sheet_name='Votos', index=False)
        df_socios.to_excel(writer, sheet_name='Socios', index=False)
    return output.getvalue()

sincronizar_socios()

tab1, tab2 = st.tabs(["Votação Sócios", "Área Administrador"])

with tab1:
    inicio, fim = get_datas()
    st.info(f"**Período Oficial:** {inicio.strftime('%d/%m %H:%M')} até {fim.strftime('%d/%m %H:%M')}")

    doc = st.text_input("Número de Documento")
    if st.button("Entrar", type="primary"):
        socio = buscar_socio(doc)
        if not socio:
            st.error("Documento não encontrado na lista de sócios!")
        elif ja_votou(doc):
            st.error("Este documento já votou!")
        else:
            st.session_state['logado'] = True
            st.session_state['doc'] = doc
            st.session_state['nome'] = socio[1]
            st.rerun()

    if st.session_state.get('logado'):
        st.success(f"Bem-vindo, {st.session_state['nome']}")
        voto = st.radio("Você vota:", ["SIM", "NÃO"], horizontal=True)
        if st.button("Confirmar Voto"):
            votar(st.session_state['doc'], voto)
            st.success("Voto registado com sucesso!")
            st.balloons()
            st.session_state['logado'] = False
            st.rerun()

with tab2:
    doc_admin = st.text_input("Número de Documento Administrador", type="password")

    if doc_admin == DOC_ADMIN:
        st.success("Acesso Administrador liberado")

        if st.button("🔄 Sincronizar Lista de Sócios"):
            sincronizar_socios()

        st.subheader("⚙️ Parametrizar Eleição")
        inicio, fim = get_datas()
        col1, col2 = st.columns(2)
        with col1:
            nova_data_inicio = st.date_input("Data Início", inicio.date())
            nova_hora_inicio = st.time_input("Hora Início", inicio.time())
        with col2:
            nova_data_fim = st.date_input("Data Fim", fim.date())
            nova_hora_fim = st.time_input("Hora Fim", fim.time())

        if st.button("Salvar Datas"):
            set_config("inicio", f"{nova_data_inicio} {nova_hora_inicio}")
            set_config("fim", f"{nova_data_fim} {nova_hora_fim}")
            st.success("Datas atualizadas!")
            st.rerun()

        st.markdown("---")
        if st.button("🔄 Reiniciar/Zerar Votação", type="primary"):
            resetar_votacao()
            st.warning("Todos os votos foram apagados!")

        excel_data = exportar_excel()
        st.download_button("📥 Exportar Resultados", data=excel_data, file_name=f"resultados_CFS_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")

        st.markdown("---")
        st.subheader("📊 Resultados em Tempo Real")
        df_resultados, total_votos, total_socios = get_resultados()
        c1, c2, c3 = st.columns(3)
        c1.metric("Total de Votos", total_votos)
        c2.metric("Total de Sócios", total_socios)
        c3.metric("Participação", f"{(total_votos/total_socios*100):.1f}%" if total_socios > 0 else "0%")
        if not df_resultados.empty:
            st.table(df_resultados)
            st.bar_chart(df_resultados.set_index('voto'))

    elif doc_admin:
        st.error("Documento de Administrador incorreto")

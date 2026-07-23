import streamlit as st
import sqlite3
import pandas as pd
from datetime import datetime
import io
import os

# -------------------------
# CONFIG
# -------------------------
st.set_page_config(page_title="Sistema de Votação", layout="centered")
ARQUIVO_SOCIOS = "dados.xlsx"

# CORES
st.markdown("""
<style>
.stButton>button {
    background-color: #004AAD;
    color: white;
    border-radius: 8px;
    height: 3em;
    width: 100%;
}
h1 {
    color: #004AAD;
}
</style>
""", unsafe_allow_html=True)

st.markdown("<h1>Sistema de Votação</h1>", unsafe_allow_html=True)

# -------------------------
# BASE DE DADOS
# -------------------------
conn = sqlite3.connect("dados.db", check_same_thread=False)
cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS config (
    chave TEXT PRIMARY KEY,
    valor TEXT
)
""")

cursor.execute("""
CREATE TABLE IF NOT EXISTS socios (
    numero_documento TEXT PRIMARY KEY,
    nome TEXT,
    votou INTEGER DEFAULT 0
)
""")

cursor.execute("""
CREATE TABLE IF NOT EXISTS votos (
    numero_documento TEXT,
    voto TEXT,
    timestamp TEXT
)
""")
conn.commit()

# -------------------------
# FUNÇÕES
# -------------------------
def carregar_socios_xlsx():
    if not os.path.exists(ARQUIVO_SOCIOS):
        return pd.DataFrame()
    df = pd.read_excel(ARQUIVO_SOCIOS)
    df['numero_documento'] = df['numero_documento'].astype(str)
    return df

def sincronizar_socios():
    df = carregar_socios_xlsx()
    if df.empty:
        return False
    for _, row in df.iterrows():
        cursor.execute("INSERT OR IGNORE INTO socios (numero_documento, nome) VALUES (?,?)",
                       (row['numero_documento'], row['nome']))
    conn.commit()
    return True

def get_config(chave, default):
    cursor.execute("SELECT valor FROM config WHERE chave=?", (chave,))
    res = cursor.fetchone()
    return res[0] if res else default

def set_config(chave, valor):
    cursor.execute("INSERT OR REPLACE INTO config VALUES (?,?)", (chave, valor))
    conn.commit()

def get_datas():
    inicio_str = get_config("inicio", "2026-04-25 00:00")
    fim_str = get_config("fim", "2026-04-26 12:00")
    return datetime.strptime(inicio_str, "%Y-%m-%d %H:%M"), datetime.strptime(fim_str, "%Y-%m-%d %H:%M")

def dentro_do_periodo():
    inicio, fim = get_datas()
    agora = datetime.now()
    return inicio <= agora <= fim

def tempo_restante():
    inicio, fim = get_datas()
    agora = datetime.now()
    if agora < inicio:
        return f"Votação começa em: {inicio.strftime('%d/%m %H:%M')}"
    elif agora > fim:
        return "Votação encerrada"
    else:
        restante = fim - agora
        horas, resto = divmod(restante.seconds, 3600)
        minutos, _ = divmod(resto, 60)
        return f"Tempo restante: {horas}h {minutos}min"

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

# Sincronizar socios ao iniciar
sincronizar_socios()

# -------------------------
# TABS
# -------------------------
tab1, tab2 = st.tabs(["Votação Sócios", "Área Administrador"])

# -------------------------
# TAB VOTAÇÃO
# -------------------------
with tab1:
    inicio, fim = get_datas()
    st.info(f"**Período:** {inicio.strftime('%d/%m %H:%M')} até {fim.strftime('%d/%m %H:%M')}")
    st.write(tempo_restante())

    if not dentro_do_periodo():
        st.warning("Fora do período de votação.")
    else:
        doc = st.text_input("Número de Documento", key="doc_voto")

        if st.button("Entrar"):
            socio = buscar_socio(doc)
            if not doc:
                st.error("Digite o número do documento")
            elif not socio:
                st.error("Documento não encontrado na lista de sócios!")
            elif ja_votou(doc):
                st.error("Este documento já votou!")
            else:
                st.session_state['logado'] = True
                st.session_state['doc'] = doc
                st.session_state['nome'] = socio[1]
                st.rerun()

        if st.session_state.get('logado'):
            st.markdown("---")
            st.success(f"Bem-vindo, {st.session_state['nome']}")
            st.subheader("Escolha seu voto")
            voto = st.radio("Você vota:", ["SIM", "NÃO"], horizontal=True)

            if st.button("Confirmar Voto"):
                votar(st.session_state['doc'], voto)
                st.success("Voto registado com sucesso!")
                st.balloons()
                st.session_state['logado'] = False
                st.session_state['doc'] = ""
                st.session_state['nome'] = ""
                st.rerun()

# -------------------------
# TAB ADMIN
# -------------------------
with tab2:
    st.caption("Acesso restrito: Rui Alberto da Cruz David - Doc: 123548")
    doc_admin = st.text_input("Número de Documento Admin", key="doc_admin")
    senha_admin = st.text_input("Senha Admin", type="password", key="senha_admin")

    if doc_admin == "123548" and senha_admin == "123548":
        st.success("Acesso Administrador liberado")

        if st.button("🔄 Sincronizar Lista de Sócios"):
            if sincronizar_socios():
                st.success("Lista de sócios atualizada do dados.xlsx")
            else:
                st.error("Arquivo dados.xlsx não encontrado")

        st.subheader("⚙️ Parametrizar Eleição")
        inicio, fim = get_datas()
        nova_data_inicio = st.date_input("Data Início", inicio.date())
        nova_hora_inicio = st.time_input("Hora Início", inicio.time())
        nova_data_fim = st.date_input("Data Fim", fim.date())
        nova_hora_fim = st.time_input("Hora Fim", fim.time())

        if st.button("Salvar Datas"):
            set_config("inicio", f"{nova_data_inicio} {nova_hora_inicio}")
            set_config("fim", f"{nova_data_fim} {nova_hora_fim}")
            st.success("Datas atualizadas!")
            st.rerun()

        st.markdown("---")
        col1, col2 = st.columns(2)
        with col1:
            if st.button("🔄 Reiniciar Votação"):
                resetar_votacao()
                st.warning("Todos os votos foram apagados! Lista de sócios mantida.")
                st.rerun()
        with col2:
            excel_data = exportar_excel()
            st.download_button(
                label="📥 Exportar para Excel",
                data=excel_data,
                file_name=f"resultados_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
            )

        st.markdown("---")
        st.subheader("📊 Resultados")
        df_resultados, total_votos, total_socios = get_resultados()

        c1, c2, c3 = st.columns(3)
        c1.metric("Total de Votos", total_votos)
        c2.metric("Total de Sócios", total_socios)
        c3.metric("Participação", f"{(total_votos/total_socios*100):.1f}%" if total_socios > 0 else "0%")

        if not df_resultados.empty:
            st.table(df_resultados)
            st.bar_chart(df_resultados.set_index('voto'))
        else:
            st.info("Nenhum voto ainda.")
    elif doc_admin or senha_admin:
        st.error("Documento ou senha incorretos")

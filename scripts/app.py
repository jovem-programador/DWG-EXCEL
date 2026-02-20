# app.py
import os
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from core_extracao import (
    preparar_pasta_temp,
    converter_dwg_para_dxf_oda,
    extrair_dados_completos_de_pasta_dxf,
    dataframe_to_excel_bytes,
)

st.set_page_config(page_title="Projeto Raio - Extração DWG/DXF", layout="wide")

#st.sidebar.image("Logo/Logotipo Projeta.png", use_container_width=True)

st.title("Projeto Raio — Extração de (.DWG/.DXF) → Excel")

with st.sidebar:
    st.image("../logo/Logotipo Projeta.png", width=220)
    st.divider()

with st.sidebar:
    st.header("Entrada")

    modo = st.radio("Modo", ["Pasta local (recomendado)", "Somente DXF (pasta)"], index=0)

    pasta_dwg = st.text_input("Pasta de DWG (origem)", value=r"C:\Users\anderson.marley\Documents\Projeta\7 - Desenvolvimento\DWG-EXCEL\dwg")
    pasta_dxf = st.text_input("Pasta de DXF (se já existir)", value=r"C:\caminho\para\DXF")

    st.divider()
    st.header("Conversão DWG → DXF (ODA)")

    converter = st.toggle("Converter DWG para DXF", value=True)
    path_oda = st.text_input("Caminho do ODAFileConverter.exe", value=r"C:\Program Files\ODA\ODAFileConverter 26.12.0\ODAFileConverter.exe")
    dxf_version = st.selectbox("Versão DXF", ["ACAD2018", "ACAD2013", "ACAD2010"], index=0)

    st.divider()
    st.header("Parâmetros")
    x_tol = st.slider("Janela X (fallback)", 50, 600, 420, step=10)
    y_tol = st.slider("Janela Y (fallback)", 1, 20, 6, step=1)

    st.divider()
    st.header("Saída")
    nome_saida = st.text_input("Nome do arquivo Excel", value="Extracao_DWG.xlsx")

processar = st.button("▶ Processar", type="primary")

col1, col2 = st.columns([3, 1], gap="large")
log_box = st.empty()

def log(msg: str):
    log_box.info(msg)

if processar:
    # validações simples
    if modo == "Pasta local (recomendado)" and (not pasta_dwg or not os.path.isdir(pasta_dwg)):
        st.error("Pasta de DWG inválida.")
        st.stop()

    if modo == "Somente DXF (pasta)" and (not pasta_dxf or not os.path.isdir(pasta_dxf)):
        st.error("Pasta de DXF inválida.")
        st.stop()

    with st.spinner("Executando..."):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)

            # 1) preparar temp de DXF
            pasta_dxf_temp = tmpdir / "DXF_Temporario"
            preparar_pasta_temp(pasta_dxf_temp)

            # 2) se precisar converter, converte; senão, usa a pasta de DXF fornecida
            if modo == "Pasta local (recomendado)":
                if converter:
                    log("Convertendo DWG → DXF via ODA...")
                    converter_dwg_para_dxf_oda(
                        path_oda=path_oda,
                        pasta_dwg_origem=Path(pasta_dwg),
                        pasta_dxf_destino=pasta_dxf_temp,
                        dxf_version=dxf_version,
                    )
                    pasta_dxf_processar = pasta_dxf_temp
                else:
                    st.error("No modo Pasta local, você marcou converter=False. Se você não quer converter, use o modo 'Somente DXF'.")
                    st.stop()
            else:
                pasta_dxf_processar = Path(pasta_dxf)

            # 3) extrair
            log("Extraindo carimbos dos DXFs...")
            dados = extrair_dados_completos_de_pasta_dxf(
                pasta_dxf=pasta_dxf_processar,
                x_tol=x_tol,
                y_tol=y_tol,
            )

            if not dados:
                st.warning("Nenhum dado extraído.")
                st.stop()

            df = pd.DataFrame(dados).drop_duplicates(subset=["Nome_Arquivo"])

            with col1:
                st.subheader("Preview (primeiras linhas)")
                st.dataframe(df.head(200), use_container_width=True)

            # 4) gerar excel em memória (bytes)
            excel_bytes = dataframe_to_excel_bytes(df)

            with col2:
                st.subheader("Resumo")
                st.metric("Arquivos processados", len(df))
                st.download_button(
                    label="⬇ Baixar Excel",
                    data=excel_bytes,
                    file_name=nome_saida,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

            log("Concluído ✅")
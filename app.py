import io
from pathlib import Path

import pandas as pd
import streamlit as st

from converter_registro import (
    agrupar,
    decodificar,
    extrair_batidas,
    gerar_csv_bytes,
    gerar_xlsx_bytes,
    montar_tabela,
)

st.set_page_config(page_title="Agrupador de Registro de Ponto", page_icon="🕒", layout="wide")

st.title("🕒 Agrupador de Registro de Ponto")
st.write(
    "Envie o arquivo bruto do relógio de ponto (uma linha por batida) e baixe o resultado "
    "com **uma linha por funcionário/dia**, no formato `Nome | Entrada | Saida | Entrada | Saida`."
)

# ------------------------------------------------------------------ opções
with st.sidebar:
    st.header("Opções")
    min_pares = st.number_input(
        "Mínimo de pares Entrada/Saida", min_value=1, max_value=6, value=2,
        help="Se algum dia tiver mais batidas, as colunas aumentam automaticamente.",
    )
    dedup = st.number_input(
        "Ignorar batidas repetidas (segundos)", min_value=0, max_value=3600, value=0,
        help="0 = desligado. Ex.: 60 descarta batidas do mesmo funcionário com menos de 60s de diferença.",
    )
    sep_label = st.selectbox(
        "Separador do CSV",
        ["Ponto e vírgula ( ; ) – Excel pt-BR", "Vírgula ( , )", "Tabulação"],
    )
    SEPS = {
        "Ponto e vírgula ( ; ) – Excel pt-BR": ";",
        "Vírgula ( , )": ",",
        "Tabulação": "\t",
    }
    sep = SEPS[sep_label]

# ------------------------------------------------------------------ upload
arquivo = st.file_uploader(
    "Arquivo de origem (Registro de atendimento)", type=["txt", "csv", "tsv"]
)

if arquivo is None:
    st.info("Aguardando o upload do arquivo.")
    st.stop()

try:
    texto = decodificar(arquivo.getvalue())
    batidas, ignoradas = extrair_batidas(texto)
except Exception as e:  # noqa: BLE001
    st.error(f"Não foi possível ler o arquivo: {e}")
    st.stop()

if not batidas:
    st.error("Nenhuma batida válida encontrada. Verifique se o arquivo tem o formato esperado.")
    st.stop()

linhas = agrupar(batidas, int(dedup))
cabecalho, tabela = montar_tabela(linhas, int(min_pares))

# ----------------------------------------------------------------- resumo
impares = [(n, h[0].strftime("%d/%m/%Y"), len(h)) for n, h in linhas if len(h) % 2]

c1, c2, c3, c4 = st.columns(4)
c1.metric("Batidas lidas", len(batidas))
c2.metric("Linhas (funcionário/dia)", len(tabela))
c3.metric("Funcionários", len({n for n, _ in linhas}))
c4.metric("Dias com batida ímpar", len(impares))

if ignoradas:
    st.caption(f"{ignoradas} linha(s) ignorada(s) (cabeçalho ou formato inválido).")

if impares:
    with st.expander(f"⚠️ {len(impares)} dia(s) com número ímpar de batidas – revise", expanded=True):
        st.write(
            "Nesses dias falta uma batida (ou há uma a mais). O último horário fica em uma "
            "coluna **Entrada** com a **Saida** vazia."
        )
        st.dataframe(
            pd.DataFrame(impares, columns=["Nome", "Data", "Qtd. de batidas"]),
            hide_index=True,
            use_container_width=True,
        )

# ------------------------------------------------------------- pré-visualização
# Nomes de coluna únicos só para exibição (o st.dataframe não aceita duplicados);
# os arquivos baixados mantêm Entrada/Saida repetidos.
cab_exibicao = ["Nome"] + [
    f"{'Entrada' if i % 2 == 0 else 'Saida'} {i // 2 + 1}" for i in range(len(cabecalho) - 1)
]
st.subheader("Pré-visualização")
st.dataframe(
    pd.DataFrame(tabela, columns=cab_exibicao),
    hide_index=True,
    use_container_width=True,
)

# ------------------------------------------------------------------ download
base = Path(arquivo.name).stem.replace(" ", "_") + "_agrupado"
st.subheader("Download")
d1, d2 = st.columns(2)

d1.download_button(
    "⬇️ Baixar CSV",
    data=gerar_csv_bytes(cabecalho, tabela, sep),
    file_name=f"{base}.csv",
    mime="text/csv",
    use_container_width=True,
)

try:
    d2.download_button(
        "⬇️ Baixar Excel (.xlsx)",
        data=gerar_xlsx_bytes(cabecalho, tabela),
        file_name=f"{base}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
except ImportError:
    d2.caption("Instale `openpyxl` para habilitar o download em Excel.")

#!/usr/bin/env python3
"""
Converte o "Registro de atendimento.txt" (uma linha por batida de ponto)
para o formato: uma linha por funcionário/dia, com as batidas agrupadas em
pares Entrada/Saida.

Entrada (bruto):
    ID  Nome  Depart.  Tempo                    Número da máquina
    3   alexsander  Not Set2   2026-09-01     07:38:23   1

Saída (agrupado):
    Nome  Entrada  Saida  Entrada  Saida
    alexsander  01/09/2026 07:38:23  01/09/2026 11:37:46  01/09/2026 13:13:00  01/09/2026 17:57:27

Uso:
    python converter_registro.py "Registro de atendimento.txt"
    python converter_registro.py entrada.txt -o saida.txt
    python converter_registro.py entrada.txt -o saida.csv
    python converter_registro.py entrada.txt -o saida.csv --sep ";"   (Excel pt-BR)
    python converter_registro.py entrada.txt -o saida.xlsx            (requer openpyxl)

Opções úteis:
    --pares N          número mínimo de pares Entrada/Saida (padrão 2).
                       Se algum dia tiver mais batidas, as colunas aumentam sozinhas.
    --dedup SEGUNDOS   ignora batidas repetidas do mesmo funcionário dentro
                       de N segundos (ex.: 60). Padrão 0 = desligado.
    --sep SEP          separador de colunas para .csv/.txt (use ";" ou "tab").
                       Padrão: "," para .csv e tabulação para .txt.
"""

import argparse
import csv
import io
import re
import sys
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path

# Aceita "2026-09-01     07:38:23" (com qualquer quantidade de espaços/tabs no meio)
RE_DATA_HORA = re.compile(r"(\d{4}-\d{2}-\d{2})\s+(\d{2}:\d{2}:\d{2})")
ENCODINGS = ("utf-8-sig", "utf-16", "cp1252", "latin-1")


# ---------------------------------------------------------------- leitura
def decodificar(dados: bytes) -> str:
    """Decodifica bytes tentando as codificações mais comuns de relógios de ponto."""
    for enc in ENCODINGS:
        try:
            return dados.decode(enc)
        except UnicodeError:
            continue
    return dados.decode("utf-8", errors="replace")


def ler_texto(caminho: Path) -> str:
    return decodificar(Path(caminho).read_bytes())


def extrair_batidas(texto: str):
    """Retorna ([(nome, datetime), ...], qtd_linhas_ignoradas) a partir do texto bruto."""
    batidas = []
    ignoradas = 0
    for linha in texto.splitlines():
        if not linha.strip():
            continue
        m = RE_DATA_HORA.search(linha)
        if not m:  # cabeçalho ou linha inválida
            ignoradas += 1
            continue
        campos = linha.split("\t")
        if len(campos) < 2 or not campos[1].strip():
            ignoradas += 1
            continue
        nome = campos[1].strip()
        momento = datetime.strptime(f"{m.group(1)} {m.group(2)}", "%Y-%m-%d %H:%M:%S")
        batidas.append((nome, momento))
    return batidas, ignoradas


def ler_batidas(caminho: Path):
    return extrair_batidas(ler_texto(caminho))


# ------------------------------------------------------------ processamento
def remover_duplicadas(horarios, segundos):
    """Descarta batidas a menos de `segundos` da última batida mantida."""
    if segundos <= 0:
        return horarios
    limite = timedelta(seconds=segundos)
    mantidas = []
    for h in horarios:
        if not mantidas or h - mantidas[-1] > limite:
            mantidas.append(h)
    return mantidas


def agrupar(batidas, dedup=0):
    """Agrupa por (nome, dia). Retorna [(nome, [datetime, ...]), ...] ordenado por data e 1ª batida."""
    por_dia = defaultdict(list)
    for nome, momento in batidas:
        por_dia[(nome, momento.date())].append(momento)

    linhas = []
    for (nome, dia), horarios in por_dia.items():
        horarios = remover_duplicadas(sorted(horarios), dedup)
        linhas.append((dia, horarios[0], nome, horarios))
    linhas.sort(key=lambda x: (x[0], x[1], x[2]))
    return [(nome, horarios) for _, _, nome, horarios in linhas]


def montar_tabela(linhas, min_pares=2):
    """Monta cabeçalho e linhas de texto no formato dd/mm/aaaa HH:MM:SS."""
    max_batidas = max((len(h) for _, h in linhas), default=0)
    pares = max(min_pares, (max_batidas + 1) // 2)

    cabecalho = ["Nome"] + ["Entrada", "Saida"] * pares
    tabela = []
    for nome, horarios in linhas:
        textos = [h.strftime("%d/%m/%Y %H:%M:%S") for h in horarios]
        textos += [""] * (pares * 2 - len(textos))  # completa com vazios
        tabela.append([nome] + textos)
    return cabecalho, tabela


# ------------------------------------------------------------------ saída
def gerar_csv_bytes(cabecalho, tabela, sep=","):
    """CSV/TXT em UTF-8 com BOM (acentos corretos no Excel)."""
    buf = io.StringIO(newline="")
    w = csv.writer(buf, delimiter=sep, lineterminator="\r\n")
    w.writerow(cabecalho)
    w.writerows(tabela)
    return buf.getvalue().encode("utf-8-sig")


def gerar_xlsx_bytes(cabecalho, tabela):
    """XLSX em memória (requer openpyxl). Valores mantidos como texto."""
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Registro"
    ws.append(cabecalho)
    for linha in tabela:
        ws.append(linha)
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 22
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def normalizar_sep(sep):
    if sep is None:
        return None
    return "\t" if sep.lower() in ("tab", "\\t", "tabulacao", "tabulação") else sep


def salvar(cabecalho, tabela, destino: Path, sep=None):
    ext = destino.suffix.lower()
    if ext == ".xlsx":
        try:
            destino.write_bytes(gerar_xlsx_bytes(cabecalho, tabela))
        except ImportError:
            sys.exit("Para gerar .xlsx instale: pip install openpyxl")
    else:
        sep = normalizar_sep(sep) or ("," if ext == ".csv" else "\t")
        destino.write_bytes(gerar_csv_bytes(cabecalho, tabela, sep))


# -------------------------------------------------------------------- CLI
def main():
    ap = argparse.ArgumentParser(description="Agrupa batidas de ponto por funcionário/dia.")
    ap.add_argument("entrada", type=Path, help="arquivo bruto (ex.: Registro de atendimento.txt)")
    ap.add_argument("-o", "--saida", type=Path, help="arquivo de saída (.txt, .csv ou .xlsx)")
    ap.add_argument("--pares", type=int, default=2, help="mínimo de pares Entrada/Saida (padrão 2)")
    ap.add_argument("--dedup", type=int, default=0, help="ignora batidas repetidas em N segundos")
    ap.add_argument("--sep", help='separador de colunas para .csv/.txt (ex.: ";" ou "tab")')
    args = ap.parse_args()

    if not args.entrada.exists():
        sys.exit(f"Arquivo não encontrado: {args.entrada}")

    destino = args.saida or args.entrada.with_name(args.entrada.stem + "_agrupado.txt")

    batidas, ignoradas = ler_batidas(args.entrada)
    if not batidas:
        sys.exit("Nenhuma batida válida encontrada no arquivo.")

    linhas = agrupar(batidas, args.dedup)
    cabecalho, tabela = montar_tabela(linhas, args.pares)
    salvar(cabecalho, tabela, destino, args.sep)

    impares = sum(1 for _, h in linhas if len(h) % 2)
    print(f"{len(batidas)} batidas -> {len(tabela)} linhas (funcionário/dia) em: {destino}")
    if ignoradas:
        print(f"  {ignoradas} linha(s) ignorada(s) (cabeçalho/inválidas).")
    if impares:
        print(f"  Atenção: {impares} dia(s) com número ímpar de batidas (última saída ficou vazia).")


if __name__ == "__main__":
    main()

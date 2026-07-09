#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Extrator de extrato Sicoob PDF -> CSV
======================================
Le um extrato de conta corrente do Sicoob em PDF e gera um arquivo CSV
com as transacoes no formato:

    DATA,TIPO,"VALOR",C/D,DESCRICAO

Exemplo de linha:
    01/06/2026,PIX REC.OUTRA IF MT,"20.171,17",C,Recebimento Pix BRASLOG T LTDA 53.983.231 0001-29 DOC.: Pix

Uso:
    python extract_sicoob_csv.py <arquivo.pdf> [saida.csv]

Se [saida.csv] nao for informado, o CSV sera salvo com o mesmo nome do PDF,
mas com extensao .csv, no mesmo diretorio.
"""

from __future__ import annotations

import csv
import os
import re
import sys
from typing import List, Dict, Optional, Tuple

import pdfplumber


# ---------------------------------------------------------------------------
# Configuracao
# ---------------------------------------------------------------------------
SKIP_TYPES = {"SALDO ANTERIOR", "SALDO BLOQ.ANTERIOR", "SALDO DO DIA"}

# Linhas de parada (fim do extrato)
STOP_MARKERS = {"RESUMO", "SAC:", "OUVIDORIA", "ENCARGOS VENCIDOS"}


# ---------------------------------------------------------------------------
# Funcoes auxiliares
# ---------------------------------------------------------------------------
def extract_year(text: str) -> Optional[int]:
    """Extrai o ano do cabecalho PERIODO ou da linha de data do extrato."""
    # Tenta PERIODO: 01/05/2026 - 31/05/2026
    m = re.search(r'PER[IÍ]ODO:\s*\d{2}/\d{2}/(\d{4})', text)
    if m:
        return int(m.group(1))
    # Tenta data no cabecalho: 30/06/2026 EXTRATO ...
    m = re.search(r'(\d{2}/\d{2}/(\d{4}))\s+EXTRATO', text)
    if m:
        return int(m.group(2))
    return None


def is_transaction_start(line: str) -> bool:
    """Verifica se a linha inicia uma transacao (comeca com DD/MM)."""
    return bool(re.match(r'^\d{2}/\d{2}\s', line))


def is_saldo_line(rest: str) -> bool:
    """Verifica se o restante da linha e um tipo SALDO que deve ser ignorado."""
    upper = rest.upper()
    for skip in SKIP_TYPES:
        if skip in upper:
            return True
    return False


def is_stop_line(line: str) -> bool:
    """Verifica se a linha e um marcador de parada (fim do extrato)."""
    upper = line.upper().strip()
    for marker in STOP_MARKERS:
        if upper.startswith(marker):
            return True
    return False


def extract_value_cd(rest: str) -> Tuple[str, str, str]:
    """
    Tenta extrair valor e C/D do restante da linha.

    Retorna (tipo, valor, cd) onde:
      - tipo: texto restante apos remover valor e C/D
      - valor: string do valor numerico (ex: "1.218,20")
      - cd: "C", "D" ou "" se nao encontrado nesta linha
    """
    # Padrao 1: valor com C/D colado (ex: "240,00C" ou "1.440,36D")
    m = re.search(r'([\d.]+,\d{2})\s*([CD])\s*$', rest)
    if m:
        valor = m.group(1)
        cd = m.group(2)
        tipo = rest[:m.start()].strip()
        return tipo, valor, cd

    # Padrao 2: valor sem C/D (C/D na proxima linha)
    m = re.search(r'([\d.]+,\d{2})\s*$', rest)
    if m:
        valor = m.group(1)
        tipo = rest[:m.start()].strip()
        return tipo, valor, ""

    # Padrao 3: valor com asterisco (ex: "0,00*")
    m = re.search(r'([\d.]+,\d{2})\*', rest)
    if m:
        valor = m.group(1)
        tipo = rest[:m.start()].strip()
        return tipo, valor, ""

    # Nenhum valor encontrado
    return rest.strip(), "", ""


# ---------------------------------------------------------------------------
# Parser principal
# ---------------------------------------------------------------------------
def parse_transactions(pdf_path: str) -> Tuple[List[Dict], Optional[int]]:
    """
    Faz o parse de todas as transacoes do PDF.

    Retorna (lista_de_transacoes, ano) onde cada transacao e um dict com:
      - date: str (DD/MM)
      - type: str (tipo da transacao)
      - value: str (valor no formato brasileiro)
      - cd: str ("C" ou "D")
      - desc: str (descricao completa)
    """
    transactions: List[Dict] = []
    year: Optional[int] = None

    with pdfplumber.open(pdf_path) as pdf:
        # Coletar texto de todas as paginas
        all_text = []
        for page in pdf.pages:
            text = page.extract_text() or ""
            if not year:
                year = extract_year(text)
            all_text.append(text)

        full_text = "\n".join(all_text)
        lines = full_text.split("\n")

    if not year:
        print("AVISO: Nao foi possivel determinar o ano. Usando 2025.", file=sys.stderr)
        year = 2025

    # Estado do parser
    i = 0
    pending_value = ""  # Valor solto que apareceu antes de uma linha DD/MM

    while i < len(lines):
        line = lines[i].strip()

        # Verificar marcadores de parada
        if is_stop_line(line):
            break

        # Verificar se e uma linha com valor solto (antes de SALDO DO DIA)
        if re.match(r'^[\d.]+,\d{2}\*?$', line) and not is_transaction_start(line):
            pending_value = line.rstrip('*')
            i += 1
            continue

        # Verificar se e inicio de transacao
        m = re.match(r'^(\d{2}/\d{2})\s+(.*)', line)
        if not m:
            # Linha C/D solta
            if line in ('C', 'D') and transactions:
                transactions[-1]['cd'] = line
            i += 1
            continue

        date_str = m.group(1)
        rest = m.group(2).strip()

        # Pular linhas de saldo
        if is_saldo_line(rest):
            pending_value = ""
            if i + 1 < len(lines) and lines[i + 1].strip() in ('C', 'D'):
                i += 2
            else:
                i += 1
            continue

        # Extrair tipo, valor e C/D
        tx_type, value, cd = extract_value_cd(rest)

        # Se nao encontrou valor na linha, usar valor pendente
        if not value and pending_value:
            value = pending_value
            pending_value = ""

        # Limpar valor pendente se nao foi usado
        if value:
            pending_value = ""

        # Verificar C/D na proxima linha
        if not cd and i + 1 < len(lines):
            next_line = lines[i + 1].strip()
            if next_line in ('C', 'D'):
                cd = next_line
                i += 1  # Consumir a linha do C/D

        # Coletar linhas de descricao
        desc_lines = []
        i += 1
        while i < len(lines):
            next_line = lines[i].strip()

            # Parar se for inicio de nova transacao
            if is_transaction_start(next_line):
                break

            # Parar se for marcador de parada
            if is_stop_line(next_line):
                break

            # Parar se for linha de saldo
            if is_saldo_line(next_line):
                break

            # Ignorar valores soltos (pertencem a saldo do dia)
            if re.match(r'^[\d.]+,\d{2}\*?$', next_line):
                pending_value = next_line.rstrip('*')
                i += 1
                continue

            # Ignorar C/D solto se ainda nao temos descricao
            if next_line in ('C', 'D') and not desc_lines:
                if not cd:
                    cd = next_line
                i += 1
                continue

            desc_lines.append(next_line)
            i += 1

        description = ' '.join(desc_lines).strip()

        # Montar transacao
        transactions.append({
            'date': date_str,
            'type': tx_type,
            'value': value,
            'cd': cd,
            'desc': description,
        })

    return transactions, year


# ---------------------------------------------------------------------------
# Geracao do CSV
# ---------------------------------------------------------------------------
def generate_csv(
    transactions: List[Dict],
    year: int,
    output_path: str,
) -> None:
    """
    Gera o arquivo CSV com as transacoes extraidas.

    Formato: DATA,TIPO,"VALOR",C/D,DESCRICAO
    """
    with open(output_path, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f, quoting=csv.QUOTE_MINIMAL)
        # Cabecalho
        writer.writerow(['DATA', 'TIPO', 'VALOR', 'C/D', 'DESCRICAO'])

        for tx in transactions:
            date_full = f"{tx['date']}/{year}"
            writer.writerow([
                date_full,
                tx['type'],
                tx['value'],
                tx['cd'],
                tx['desc'],
            ])


# ---------------------------------------------------------------------------
# Ponto de entrada
# ---------------------------------------------------------------------------
def main():
    if len(sys.argv) < 2:
        print("Uso: python extract_sicoob_csv.py <arquivo.pdf> [saida.csv]")
        print("Exemplo: python extract_sicoob_csv.py extrato.pdf saida.csv")
        sys.exit(1)

    pdf_path = sys.argv[1]

    if not os.path.isfile(pdf_path):
        print(f"ERRO: Arquivo nao encontrado: {pdf_path}", file=sys.stderr)
        sys.exit(1)

    # Determinar caminho de saida
    if len(sys.argv) >= 3:
        csv_path = sys.argv[2]
    else:
        base = os.path.splitext(pdf_path)[0]
        csv_path = base + '.csv'

    print(f"Processando: {pdf_path}")
    print("-" * 50)

    # Extrair transacoes
    transactions, year = parse_transactions(pdf_path)

    if not transactions:
        print("AVISO: Nenhuma transacao encontrada no PDF.", file=sys.stderr)
        sys.exit(1)

    print(f"Ano detectado: {year}")
    print(f"Transacoes encontradas: {len(transactions)}")
    print()

    # Mostrar primeiras transacoes como preview
    print("Preview (primeiras 10 transacoes):")
    print(f"{'DATA':<12} {'TIPO':<25} {'VALOR':>12} {'C/D':<4} {'DESCRICAO'}")
    print("-" * 100)
    for tx in transactions[:10]:
        date_full = f"{tx['date']}/{year}"
        desc_preview = (tx['desc'][:40] + '...') if len(tx['desc']) > 40 else tx['desc']
        print(f"{date_full:<12} {tx['type']:<25} {tx['value']:>12} {tx['cd']:<4} {desc_preview}")
    if len(transactions) > 10:
        print(f"... e mais {len(transactions) - 10} transacoes")
    print()

    # Gerar CSV
    generate_csv(transactions, year, csv_path)
    print(f"CSV salvo em: {csv_path}")

    # Estatisticas
    credit_count = sum(1 for tx in transactions if tx['cd'] == 'C')
    debit_count = sum(1 for tx in transactions if tx['cd'] == 'D')
    print(f"  Creditos: {credit_count}")
    print(f"  Debitos: {debit_count}")


if __name__ == '__main__':
    main()

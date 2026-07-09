#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Extrator de extrato Sicoob PDF -> CSV
======================================
Le um extrato de conta corrente do Sicoob em PDF e gera um arquivo CSV com as 
transacoes no formato: DATA,TIPO,"VALOR",C/D,DESCRICAO

Uso: python conversor.py <arquivo.pdf> [saida.csv]
"""
from __future__ import annotations

import csv
import os
import re
import sys
from datetime import datetime
from typing import List, Dict, Optional, Tuple

import pdfplumber

# ---------------------------------------------------------------------------
# Configuracao
# ---------------------------------------------------------------------------
SKIP_TYPES = {"SALDO ANTERIOR", "SALDO BLOQ.ANTERIOR", "SALDO DO DIA", "SALDO DISPONIVEL"}

# Linhas de parada (fim do extrato)
STOP_MARKERS = {"RESUMO", "SAC:", "OUVIDORIA", "ENCARGOS VENCIDOS", "TOTALIZADORES"}

# ---------------------------------------------------------------------------
# Funcoes auxiliares
# ---------------------------------------------------------------------------
def extract_year(text: str) -> Optional[int]:
    """Extrai o ano do cabecalho PERIODO ou da linha de data do extrato."""
    m = re.search(r'PER[IÍ]ODO:\s*\d{2}/\d{2}/(\d{4})', text)
    if m:
        return int(m.group(1))
    m = re.search(r'(\d{2}/\d{2}/(\d{4}))\s+EXTRATO', text)
    if m:
        return int(m.group(2))
    return None

def is_page_header(line: str) -> bool:
    """Verifica se a linha é um cabeçalho repetido de quebra de página."""
    upper = line.upper()
    if re.search(r'^DATA\s+HIST[OÓ]RICO', upper):
        return True
    if "SISTEMA DE COOPERATIVAS DE CR" in upper:
        return True
    if "EXTRATO DE CONTA CORRENTE" in upper:
        return True
    if re.match(r'^PER[IÍ]ODO:\s*\d{2}/\d{2}/\d{4}', upper):
        return True
    return False

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
    Retorna (tipo, valor, cd)
    """
    m = re.search(r'([\d.]+,\d{2})\s*([CD])\s*$', rest)
    if m:
        valor = m.group(1)
        cd = m.group(2)
        tipo = rest[:m.start()].strip()
        return tipo, valor, cd

    m = re.search(r'([\d.]+,\d{2})\s*$', rest)
    if m:
        valor = m.group(1)
        tipo = rest[:m.start()].strip()
        return tipo, valor, ""

    m = re.search(r'([\d.]+,\d{2})\*', rest)
    if m:
        valor = m.group(1)
        tipo = rest[:m.start()].strip()
        return tipo, valor, ""

    return rest.strip(), "", ""

# ---------------------------------------------------------------------------
# Parser principal
# ---------------------------------------------------------------------------
def parse_transactions(pdf_path: str) -> Tuple[List[Dict], Optional[int]]:
    transactions: List[Dict] = []
    year: Optional[int] = None
    all_lines = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            # layout=True preserva colunas, evitando que os valores se misturem na descrição
            text = page.extract_text(layout=True) 
            if not text:
                continue
                
            if not year:
                # Tenta extrair da página bruta antes de fatiar
                raw_text = page.extract_text() or ""
                year = extract_year(raw_text)

            for line in text.split("\n"):
                clean_line = line.strip()
                if not clean_line:
                    continue
                # Remove cabeçalhos duplicados ao longo das quebras de página
                if is_page_header(clean_line):
                    continue
                all_lines.append(clean_line)

    if not year:
        year = datetime.now().year
        print(f"AVISO: Nao foi possivel determinar o ano no arquivo. Usando ano atual ({year}).", file=sys.stderr)

    lines = all_lines
    i = 0
    pending_value = ""

    while i < len(lines):
        line = lines[i].strip()

        if is_stop_line(line):
            break

        if re.match(r'^[\d.]+,\d{2}\*?$', line) and not is_transaction_start(line):
            pending_value = line.rstrip('*')
            i += 1
            continue

        m = re.match(r'^(\d{2}/\d{2})\s+(.*)', line)
        if not m:
            if line in ('C', 'D') and transactions:
                transactions[-1]['cd'] = line
            i += 1
            continue

        date_str = m.group(1)
        rest = m.group(2).strip()

        if is_saldo_line(rest):
            pending_value = ""
            if i + 1 < len(lines) and lines[i + 1].strip() in ('C', 'D'):
                i += 2
            else:
                i += 1
            continue

        tx_type, value, cd = extract_value_cd(rest)

        if not value and pending_value:
            value = pending_value
            pending_value = ""

        if value:
            pending_value = ""

        if not cd and i + 1 < len(lines):
            next_line = lines[i + 1].strip()
            if next_line in ('C', 'D'):
                cd = next_line
                i += 1 

        desc_lines = []
        i += 1
        
        while i < len(lines):
            next_line = lines[i].strip()

            if is_transaction_start(next_line) or is_stop_line(next_line) or is_saldo_line(next_line):
                break

            if re.match(r'^[\d.]+,\d{2}\*?$', next_line):
                pending_value = next_line.rstrip('*')
                i += 1
                continue

            if next_line in ('C', 'D') and not desc_lines:
                if not cd:
                    cd = next_line
                i += 1
                continue

            desc_lines.append(next_line)
            i += 1

        # Limpa múltiplos espaços extras gerados pelo layout=True
        raw_description = ' '.join(desc_lines).strip()
        description = re.sub(r'\s+', ' ', raw_description)

        transactions.append({
            'date': date_str,
            'type': re.sub(r'\s+', ' ', tx_type), # Limpa espaços no tipo também
            'value': value,
            'cd': cd,
            'desc': description,
        })

    return transactions, year

# ---------------------------------------------------------------------------
# Geracao do CSV
# ---------------------------------------------------------------------------
def generate_csv(transactions: List[Dict], year: int, output_path: str) -> None:
    with open(output_path, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f, quoting=csv.QUOTE_MINIMAL)
        writer.writerow(['DATA', 'TIPO', 'VALOR', 'C/D', 'DESCRICAO'])
        
        for tx in transactions:
            date_full = f"{tx['date']}/{year}"
            writer.writerow([
                date_full, 
                tx['type'], 
                tx['value'], 
                tx['cd'], 
                tx['desc']
            ])

# ---------------------------------------------------------------------------
# Ponto de entrada
# ---------------------------------------------------------------------------
def main():
    if len(sys.argv) < 2:
        print("Uso: python conversor.py <arquivo.pdf> [saida.csv]")
        sys.exit(1)

    pdf_path = sys.argv[1]
    if not os.path.isfile(pdf_path):
        print(f"ERRO: Arquivo nao encontrado: {pdf_path}", file=sys.stderr)
        sys.exit(1)

    if len(sys.argv) >= 3:
        csv_path = sys.argv[2]
    else:
        base = os.path.splitext(pdf_path)[0]
        csv_path = base + '.csv'

    print(f"Processando: {pdf_path}")
    print("-" * 50)

    transactions, year = parse_transactions(pdf_path)

    if not transactions:
        print("AVISO: Nenhuma transacao encontrada no PDF.", file=sys.stderr)
        sys.exit(1)

    print(f"Ano detectado: {year}")
    print(f"Transacoes encontradas: {len(transactions)}\n")

    print("Preview (primeiras 10 transacoes):")
    print(f"{'DATA':<12} {'TIPO':<25} {'VALOR':>12} {'C/D':<4} {'DESCRICAO'}")
    print("-" * 100)
    for tx in transactions[:10]:
        date_full = f"{tx['date']}/{year}"
        desc_preview = (tx['desc'][:40] + '...') if len(tx['desc']) > 40 else tx['desc']
        # Truncar o tipo para o visual do preview ficar alinhado
        type_preview = tx['type'][:23] + '..' if len(tx['type']) > 25 else tx['type'] 
        print(f"{date_full:<12} {type_preview:<25} {tx['value']:>12} {tx['cd']:<4} {desc_preview}")

    if len(transactions) > 10:
        print(f"... e mais {len(transactions) - 10} transacoes\n")

    generate_csv(transactions, year, csv_path)
    print(f"CSV salvo em: {csv_path}")

    credit_count = sum(1 for tx in transactions if tx['cd'] == 'C')
    debit_count = sum(1 for tx in transactions if tx['cd'] == 'D')
    print(f"  Creditos: {credit_count}")
    print(f"  Debitos: {debit_count}")

if __name__ == '__main__':
    main()

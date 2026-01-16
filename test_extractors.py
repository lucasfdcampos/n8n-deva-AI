#!/usr/bin/env python3
"""Testa os extractors antes de usar no n8n"""

import pandas as pd
from datetime import datetime

# Simula a lógica do Clean Dashboard
def clean_dashboard(file_path):
    df = pd.read_excel(file_path, sheet_name="Dashboard", header=None)
    
    resultado = {
        "caixa": None,
        "janeiro": {
            "receber": None,
            "recebido": None,
            "pagar": None,
            "pago": None
        },
        "saldo": {},
        "caixaProjetado": {}
    }
    
    for i in range(len(df)):
        row = df.iloc[i]
        
        # Caixa (linha 5, col 1 -> valor linha 7, col 1)
        if pd.notna(row[1]) and row[1] == 'Caixa':
            if i + 2 < len(df) and pd.notna(df.iloc[i + 2, 1]):
                resultado["caixa"] = df.iloc[i + 2, 1]
        
        # Receber (Janeiro) - col 9
        if pd.notna(row[9]) and row[9] == 'Receber (Janeiro)':
            if i + 2 < len(df) and pd.notna(df.iloc[i + 2, 9]):
                resultado["janeiro"]["receber"] = df.iloc[i + 2, 9]
        
        # Recebido (Janeiro) - col 9
        if pd.notna(row[9]) and row[9] == 'Recebido (Janeiro)':
            if i + 2 < len(df) and pd.notna(df.iloc[i + 2, 9]):
                resultado["janeiro"]["recebido"] = df.iloc[i + 2, 9]
        
        # Pagar (Janeiro) - col 9
        if pd.notna(row[9]) and row[9] == 'Pagar (Janeiro)':
            if i + 2 < len(df) and pd.notna(df.iloc[i + 2, 9]):
                resultado["janeiro"]["pagar"] = df.iloc[i + 2, 9]
        
        # Pago (Janeiro) - col 9
        if pd.notna(row[9]) and row[9] == 'Pago (Janeiro)':
            if i + 2 < len(df) and pd.notna(df.iloc[i + 2, 9]):
                resultado["janeiro"]["pago"] = df.iloc[i + 2, 9]
        
        # Caixa Projetado (Janeiro) - col 9
        if pd.notna(row[9]) and row[9] == 'Caixa Projetado (Janeiro)':
            if i + 2 < len(df) and pd.notna(df.iloc[i + 2, 9]):
                resultado["caixaProjetado"]["Janeiro"] = df.iloc[i + 2, 9]
        
        # Janeiro na col 4 (linha 57)
        if pd.notna(row[4]) and row[4] == 'Janeiro':
            if i + 1 < len(df) and i + 2 < len(df):
                entrada = df.iloc[i + 1, 4] if pd.notna(df.iloc[i + 1, 4]) else 0
                saida = df.iloc[i + 2, 4] if pd.notna(df.iloc[i + 2, 4]) else 0
                resultado["saldo"]["Janeiro"] = entrada - saida
    
    return resultado


# Simula a lógica do Clean Fluxo
def clean_fluxo(file_path):
    df = pd.read_excel(file_path, sheet_name="Fluxo de Caixa - Mensal", header=None)
    
    meses_nomes = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
                   'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
    
    mes_atual = datetime.now().month - 1
    nome_mes_atual = meses_nomes[mes_atual]
    
    resultado = {
        "realizado": {
            "saldoInicial": {},
            "entradas": {},
            "saidas": {},
            "caixaFinal": {}
        },
        "emAndamento": {
            "entradasAReceber": {},
            "saidasAPagar": {},
            "movimentacaoTotal": {},
            "saldo": {}
        }
    }
    
    for i in range(len(df)):
        row = df.iloc[i]
        primeira_col = row[0]
        
        if pd.isna(primeira_col):
            continue
        
        # Saldo Inicial
        if primeira_col == 'Saldo Inicial':
            for idx, mes in enumerate(meses_nomes):
                valor = row[idx + 1]
                if pd.notna(valor) and mes == nome_mes_atual:
                    resultado["realizado"]["saldoInicial"][mes] = float(valor)
        
        # Entradas
        if primeira_col == 'Entradas':
            for idx, mes in enumerate(meses_nomes):
                valor = row[idx + 1]
                if pd.notna(valor) and valor != 0:
                    resultado["realizado"]["entradas"][mes] = {"total": float(valor)}
        
        # Saídas
        if primeira_col == 'Saídas':
            for idx, mes in enumerate(meses_nomes):
                valor = row[idx + 1]
                if pd.notna(valor) and valor != 0:
                    resultado["realizado"]["saidas"][mes] = {"total": float(valor)}
        
        # Caixa
        if primeira_col == 'Caixa':
            for idx, mes in enumerate(meses_nomes):
                valor = row[idx + 1]
                if pd.notna(valor) and mes == nome_mes_atual:
                    resultado["realizado"]["caixaFinal"][mes] = float(valor)
        
        # Entradas (A receber)
        if primeira_col == 'Entradas (A receber)':
            for idx in range(5):
                valor = row[idx + 1]
                if pd.notna(valor) and valor != 0:
                    resultado["emAndamento"]["entradasAReceber"][meses_nomes[idx]] = {"total": float(valor)}
        
        # Saídas (A pagar)
        if primeira_col == 'Saídas (A pagar)':
            for idx in range(5):
                valor = row[idx + 1]
                if pd.notna(valor) and valor != 0:
                    resultado["emAndamento"]["saidasAPagar"][meses_nomes[idx]] = {"total": float(valor)}
        
        # Movimentação Total
        if primeira_col == 'Movimentação Total':
            for idx in range(5):
                valor = row[idx + 1]
                if pd.notna(valor):
                    resultado["emAndamento"]["movimentacaoTotal"][meses_nomes[idx]] = float(valor)
        
        # Saldo
        if primeira_col == 'Saldo':
            for idx in range(5):
                valor = row[idx + 1]
                if pd.notna(valor) and valor != 0:
                    resultado["emAndamento"]["saldo"][meses_nomes[idx]] = float(valor)
    
    return resultado


if __name__ == "__main__":
    import json
    
    print("🧪 Testando extractors...\n")
    
    # Testa Dashboard
    print("📊 Dashboard:")
    dashboard = clean_dashboard("Controle - 2026.xlsx")
    print(json.dumps(dashboard, indent=2, ensure_ascii=False))
    
    print("\n" + "="*60 + "\n")
    
    # Testa Fluxo
    print("💰 Fluxo de Caixa:")
    fluxo = clean_fluxo("Controle - 2026.xlsx")
    print(json.dumps(fluxo, indent=2, ensure_ascii=False))
    
    # Valida
    print("\n" + "="*60)
    print("✅ VALIDAÇÃO:")
    print(f"   Caixa: R$ {dashboard['caixa']:,.2f}" if dashboard['caixa'] else "   ❌ Caixa não encontrado")
    print(f"   Receber Jan: R$ {dashboard['janeiro']['receber']:,.2f}" if dashboard['janeiro']['receber'] else "   ❌ Receber não encontrado")
    print(f"   Entradas Jan (realizado): {fluxo['realizado']['entradas'].get('Janeiro', {})}")
    print(f"   Entradas Jan (a receber): {fluxo['emAndamento']['entradasAReceber'].get('Janeiro', {})}")
    print(f"   Saldo projetado: {fluxo['emAndamento']['saldo'].get('Janeiro', 'N/A')}")

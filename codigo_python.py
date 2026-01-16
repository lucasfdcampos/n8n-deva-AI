import io
import base64
import zipfile
import xml.etree.ElementTree as ET
import unicodedata
import re

def normalize(text):
    if not text: return ""
    return "".join(c for c in unicodedata.normalize('NFD', str(text)) if unicodedata.category(c) != 'Mn').lower().strip()

def to_float(val):
    if val is None or val == "": return None
    try:
        s = str(val).replace('R$', '').replace(' ', '').replace('.', '').replace(',', '.')
        # Caso o valor venha com múltiplos pontos de milhar, mantém apenas o último como decimal
        if s.count('.') > 1:
            partes = s.split('.')
            s = "".join(partes[:-1]) + "." + partes[-1]
        return float(s)
    except:
        return None

def criar_resumo_de_texto(sheet_name, rows_data):
    resumo_texto = f"### Aba: {sheet_name}\n"
    metricas = {}
    sheet_norm = normalize(sheet_name)
    
    # Criar matriz de busca rápida
    matrix = {}
    max_row = 0
    cols_letras = []
    for r_idx, row in enumerate(rows_data):
        max_row = max(max_row, r_idx)
        for col_letter, val in row.items():
            if col_letter not in cols_letras: cols_letras.append(col_letter)
            matrix[(r_idx, col_letter)] = val
    
    sorted_cols = sorted(cols_letras, key=lambda x: (len(x), x))

    for (r, c), val in matrix.items():
        txt = normalize(str(val))
        
        # --- FUNÇÃO SCANNER (Procura o número na linha ou abaixo) ---
        def scanner_universal(label_busca, chave_saida, modo='linha'):
            if label_busca in txt and chave_saida not in metricas:
                if modo == 'linha':
                    # Procura o primeiro número da esquerda para a direita na mesma linha
                    for col in sorted_cols:
                        v = to_float(matrix.get((r, col)))
                        if v is not None and v != 0:
                            metricas[chave_saida] = v
                else: # modo vertical (para Dashboard)
                    for offset in range(1, 4):
                        v = to_float(matrix.get((r + offset, c)))
                        if v is not None:
                            metricas[chave_saida] = v
                            break

        # Filtros por Aba
        if 'dashboard' in sheet_norm:
            scanner_universal('caixa', 'Saldo em Caixa Atual', 'vertical')
            scanner_universal('receber (dezembro)', 'Total Receber (Dez)', 'vertical')
            scanner_universal('pagar (dezembro)', 'Total Pagar (Dez)', 'vertical')
            scanner_universal('saldo (dezembro)', 'Saldo Projetado Dez', 'vertical')

        elif 'fluxo' in sheet_norm:
            if 'receitas de vendas' in txt: scanner_universal(txt, 'Faturamento Total')
            if 'resultado liquido' in txt or 'movimentacao total' in txt: scanner_universal(txt, 'Resultado Liquido')
            if 'retiradas' in txt: scanner_universal(txt, 'Total Retiradas')

        elif 'metas' in sheet_norm:
            if 'meta' == txt: scanner_universal(txt, 'Meta Anual')
            if 'realizado' == txt: scanner_universal(txt, 'Realizado Anual')

    # Montagem do Resumo
    resumo_texto += "DADOS ESTRATÉGICOS:\n"
    if metricas:
        for k, v in metricas.items():
            resumo_texto += f"- {k}: R$ {v:,.2f}\n"
    else:
        resumo_texto += "- Dados consolidados não identificados nesta aba.\n"
    
    return resumo_texto

# --- MANTÉM A ESTRUTURA DE EXTRAÇÃO DO XLSX ---
output = []
for input_item in _input.all():
    try:
        binary_file = input_item.binary['data']['data']
        file_buffer = base64.b64decode(binary_file)
        processed_sheets = {}
        with zipfile.ZipFile(io.BytesIO(file_buffer)) as xlsx:
            ss_root = ET.fromstring(xlsx.read('xl/sharedStrings.xml'))
            shared_strings = [si.text for si in ss_root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')]
            rels_root = ET.fromstring(xlsx.read('xl/_rels/workbook.xml.rels'))
            rels = {r.get('Id'): r.get('Target') for r in rels_root.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship')}
            wb_root = ET.fromstring(xlsx.read('xl/workbook.xml'))
            
            for s in wb_root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet'):
                nome_aba = s.get('name')
                rId = s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                path = f"xl/{rels[rId]}" if not rels[rId].startswith('xl/') else rels[rId]
                root = ET.fromstring(xlsx.read(path))
                rows = []
                for r_elem in root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row'):
                    row_data = {}
                    for c in r_elem.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
                        ref = c.get('r')
                        col = "".join(re.findall("[a-zA-Z]+", ref))
                        v_tag = c.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
                        if v_tag is not None:
                            val = v_tag.text
                            if c.get('t') == 's': val = shared_strings[int(val)]
                            row_data[col] = val
                    rows.append(row_data)
                processed_sheets[nome_aba] = {"resumo": criar_resumo_de_texto(nome_aba, rows)}
                    
        output.append({"json": {"analysis": {"status": "success", "sheets": processed_sheets}}})
    except Exception as e:
        output.append({"json": {"status": "error", "error": str(e)}})
return output
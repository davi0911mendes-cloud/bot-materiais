"""
Bot Telegram – Controle de Gastos com Materiais
================================================
Armazenamento: Google Sheets (via gspread)
Deploy:        Railway (nuvem)
"""

import os
import json
import base64
import logging
import time
from datetime import datetime

import gspread
from telegram import Update, ReplyKeyboardMarkup, ReplyKeyboardRemove
from telegram.ext import Application, CommandHandler, MessageHandler, ContextTypes, filters, PicklePersistence

# ── Configuração ─────────────────────────────────────────────────────────────
BOT_TOKEN = os.environ.get("BOT_TOKEN", "")

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

# ── Chaves de estado manual ───────────────────────────────────────────────────
STEP      = "step"
DADOS     = "dados"

S_IDLE        = None
S_FORNECEDOR  = "fornecedor"
S_MATERIAL    = "material"
S_QUANTIDADE  = "quantidade"
S_UNIDADE     = "unidade"
S_PRECO       = "preco"
S_NOTA_FISCAL = "nota_fiscal"
S_PAGAMENTO   = "pagamento"
S_CONFIRMAR   = "confirmar"
S_LIMPAR      = "limpar"

# ── Teclados ──────────────────────────────────────────────────────────────────
OUTRO_FORN = "Outro fornecedor"
OUTRO_MAT  = "Outro material"

TECLADO_FORNECEDORES = [
    ["LISBOA", "MADECENTER"],
    ["LEO MADEIRAS", "VERDMADE"],
    ["CENCOMAL", "MADEREIRAS EXTRAS"],
    ["FGV", "HARDT"],
    ["HD FERRAGENS", "HAYD FERRAGENS"],
    ["ALTAPE FILMES E FITAS", "KILDERY THINNER"],
    ["PEQUENOS FORNECEDORES VARIAVEIS"],
    [OUTRO_FORN],
]

TECLADO_MATERIAIS = [
    ["CHAPAS UNICOLOR 18MM", "CHAPAS MADEIRADO 18MM"],
    ["CHAPAS UNICOLOR 15MM", "CHAPAS MADEIRADO 15MM"],
    ["CHAPAS BRANCO 18MM", "CHAPAS BRANCO 15MM"],
    ["CHAPAS BRANCO 6MM", "CHAPAS UNICOLOR 6MM"],
    ["CHAPAS MADEIRADO 6MM", "FITA DE BORDA BRANCA 0,45"],
    ["FITA DE BORDA COLORIDA 0,45", "FITA DE BORDA BRANCA 1MM"],
    ["FITA DE BORDA COLORIDA 1MM", "CORREDICA INVISIVEL"],
    ["CORREDICA TELESCOPIA", "DOBRADICA CURVA"],
    ["DOBRADICA RETA", "COLA FORMICA"],
    ["COLA EXPANSIVA", "COLA PUR COLADEIRA"],
    ["COLA INSTANTANEA", "PARAFUSOS"],
    ["MINIFIX", "CAVILHA"],
    ["TAMBOR", "PIVO DE PORTA"],
    ["THINNER", "ALCOOL E VASELINA"],
    ["ESTOPA", OUTRO_MAT],
]

TECLADO_UNIDADES = [
    ["kg", "g"],
    ["L", "mL"],
    ["m", "m2"],
    ["un", "cx"],
    ["sacos", "pecas"],
]

TECLADO_PAGAMENTO = [
    ["PIX", "DINHEIRO"],
    ["CARTAO", "CARTEIRA"],
]


# ── Google Sheets ─────────────────────────────────────────────────────────────
def _get_credentials_dict() -> dict:
    """Lê credenciais do Google — tenta base64 primeiro, depois JSON raw."""
    # Tenta variável base64 (mais confiável no Railway)
    b64 = os.environ.get("GOOGLE_CREDENTIALS_B64", "").strip()
    logger.info(f"[CREDS] GOOGLE_CREDENTIALS_B64 presente: {bool(b64)} (len={len(b64)})")
    if b64:
        try:
            decoded = base64.b64decode(b64).decode("utf-8-sig")
            creds = json.loads(decoded.strip())
            logger.info("[CREDS] B64 decodificado com sucesso!")
            return creds
        except Exception as e:
            logger.error(f"[CREDS] Falha ao decodificar B64: {e}")

    # Fallback: JSON raw
    raw = os.environ.get("GOOGLE_CREDENTIALS", "").strip()
    # Remove BOM e caracteres invisíveis do início
    raw = raw.lstrip("﻿\x00\r\n ").strip()
    logger.info(f"[CREDS] GOOGLE_CREDENTIALS presente: {bool(raw)} (len={len(raw)}, inicio={repr(raw[:10]) if raw else 'vazio'})")
    if not raw:
        raise ValueError("Nenhuma credencial Google configurada!")
    return json.loads(raw)


def _get_sheet():
    spreadsheet_id = os.environ.get("SPREADSHEET_ID", "").strip()
    gc          = gspread.service_account_from_dict(_get_credentials_dict())
    spreadsheet = gc.open_by_key(spreadsheet_id)
    try:
        sheet = spreadsheet.worksheet("Registro de Compras")
    except gspread.WorksheetNotFound:
        sheet = spreadsheet.add_worksheet("Registro de Compras", rows=1000, cols=10)
        sheet.append_row(
            ["#", "Data", "Fornecedor", "Material / Produto",
             "Qtd", "Unidade", "Preco Unit. (R$)", "Total (R$)", "Nota Fiscal", "Pagamento"],
            value_input_option="USER_ENTERED",
        )
    return sheet


def salvar_registro(dados: dict) -> int:
    sheet    = _get_sheet()
    all_vals = sheet.get_all_values()

    # Remove qualquer linha TOTAL antes de salvar (evita conflito com mesclagem)
    for r in sorted(range(2, len(all_vals) + 1), reverse=True):
        idx = r - 1
        if idx < len(all_vals) and all_vals[idx] and "TOTAL" in str(all_vals[idx][0]).upper():
            sheet.delete_rows(r)
            logger.info(f"[SALVAR] Removeu linha TOTAL em {r}")

    # Re-lê após limpeza
    all_vals  = sheet.get_all_values()
    data_rows = [r for r in all_vals[1:] if len(r) > 1 and str(r[1]).strip()]
    numero    = len(data_rows) + 1
    total     = dados["quantidade"] * dados["preco"]
    linha     = [
        numero,
        dados["data"].strftime("%d/%m/%Y"),
        dados["fornecedor"],
        dados["material"],
        dados["quantidade"],
        dados["unidade"],
        dados["preco"],
        total,
        dados.get("nota_fiscal", "-"),
        dados.get("pagamento", "-"),
    ]
    sheet.append_row(linha, value_input_option="USER_ENTERED")
    return numero


def gerar_resumo() -> str:
    sheet = _get_sheet()
    rows  = sheet.get_all_values()[1:]
    rows  = [r for r in rows if any(r)]
    if not rows:
        return "Nenhum registro ainda."
    total_geral    = 0.0
    por_fornecedor = {}
    for r in rows:
        try:
            total = float(str(r[7]).replace(",", ".").replace("R$", "").strip())
        except (ValueError, IndexError):
            total = 0.0
        forn = r[2] if len(r) > 2 and r[2] else "-"
        total_geral += total
        por_fornecedor[forn] = por_fornecedor.get(forn, 0.0) + total
    linhas = [
        f"Resumo Geral — {len(rows)} compra(s)\n",
        f"Total Gasto: R$ {total_geral:,.2f}\n",
        "─────────────────────",
        "Por Fornecedor:",
    ]
    for forn, val in sorted(por_fornecedor.items(), key=lambda x: -x[1]):
        linhas.append(f"  {forn}: R$ {val:,.2f}")
    return "\n".join(linhas)


def ultimos_registros(n: int = 5) -> str:
    sheet  = _get_sheet()
    rows   = sheet.get_all_values()[1:]
    rows   = [r for r in rows if any(r)]
    if not rows:
        return "Nenhum registro ainda."
    ultimas = rows[-n:][::-1]
    linhas  = [f"Ultimas {len(ultimas)} compra(s):\n"]
    for r in ultimas:
        try:
            total = float(str(r[7]).replace(",", ".").replace("R$", "").strip())
        except (ValueError, IndexError):
            total = 0.0
        linhas.append(f"- {r[1]} | {r[3]} | {r[4]} {r[5]} | R$ {total:,.2f}")
    return "\n".join(linhas)


def resetar_estrutura():
    """Preserva os dados reais e reconstrói a planilha do zero (sem TOTAL, sem mesclagens)."""
    spreadsheet_id = os.environ.get("SPREADSHEET_ID", "").strip()
    gc             = gspread.service_account_from_dict(_get_credentials_dict())
    spreadsheet    = gc.open_by_key(spreadsheet_id)

    CABECALHO = ["#", "Data", "Fornecedor", "Material / Produto",
                 "Qtd", "Unidade", "Preco Unit. (R$)", "Total (R$)", "Nota Fiscal", "Pagamento"]

    try:
        ws = spreadsheet.worksheet("Registro de Compras")
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet("Registro de Compras", rows=600, cols=10)
        ws.append_row(CABECALHO, value_input_option="USER_ENTERED")
        return 0

    # Lê todos os dados antes de apagar
    all_vals = ws.get_all_values()
    # Filtra apenas linhas com data real (ignora TOTAL, linhas vazias e cabeçalho)
    dados_reais = [
        r for r in all_vals[1:]
        if len(r) > 1
        and str(r[1]).strip()
        and "TOTAL" not in str(r[0]).upper()
        and any(str(v).strip() for v in r)
    ]

    # Renumera os registros em sequência
    for i, r in enumerate(dados_reais):
        if len(r) > 0:
            r[0] = str(i + 1)

    # Limpa tudo (conteúdo + formatação + mesclagens)
    ws.clear()

    # Reescreve cabeçalho + dados limpos
    ws.append_row(CABECALHO, value_input_option="USER_ENTERED")
    if dados_reais:
        ws.update(f"A2:J{1 + len(dados_reais)}", dados_reais, value_input_option="USER_ENTERED")

    logger.info(f"[RESETAR] {len(dados_reais)} registros preservados")
    return len(dados_reais)


def limpar_planilha():
    spreadsheet_id = os.environ.get("SPREADSHEET_ID", "").strip()
    gc             = gspread.service_account_from_dict(_get_credentials_dict())
    spreadsheet    = gc.open_by_key(spreadsheet_id)
    try:
        ws = spreadsheet.worksheet("Registro de Compras")
        ws.clear()  # apaga TODO conteúdo e formatação
        ws.append_row(
            ["#", "Data", "Fornecedor", "Material / Produto",
             "Qtd", "Unidade", "Preco Unit. (R$)", "Total (R$)", "Nota Fiscal", "Pagamento"],
            value_input_option="USER_ENTERED",
        )
    except gspread.WorksheetNotFound:
        pass
    for nome in ("Por Fornecedor", "Por Material"):
        try:
            spreadsheet.del_worksheet(spreadsheet.worksheet(nome))
        except Exception:
            pass


def aplicar_formatacao():
    MAX_DATA = 500   # pré-formata 500 linhas (suficiente para anos de uso)

    spreadsheet_id = os.environ.get("SPREADSHEET_ID", "").strip()
    gc             = gspread.service_account_from_dict(_get_credentials_dict())
    spreadsheet    = gc.open_by_key(spreadsheet_id)

    try:
        ws = spreadsheet.worksheet("Registro de Compras")
    except gspread.WorksheetNotFound:
        raise ValueError("Aba 'Registro de Compras' nao encontrada.")

    # Garante linhas suficientes
    ws.resize(rows=MAX_DATA + 50, cols=10)

    # ── 1. Lê todos os dados e remove linhas TOTAL antigas ────────────────────
    all_rows = ws.get_all_values()
    for r in sorted(range(2, len(all_rows) + 1), reverse=True):
        idx = r - 1
        if idx < len(all_rows) and all_rows[idx] and "TOTAL" in str(all_rows[idx][0]).upper():
            ws.delete_rows(r)

    # Re-lê após remoção do TOTAL antigo
    all_rows = ws.get_all_values()
    dados = [
        r for r in all_rows[1:]
        if len(r) > 1 and str(r[1]).strip()
    ]
    last_row = 1 + len(dados)  # linha do último dado real
    logger.info(f"[FORMATAR] {len(dados)} registros encontrados, last_row={last_row}")

    # ── 2. Cabeçalho ──────────────────────────────────────────────────────────
    ws.update("A1:J1", [["#", "Data", "Fornecedor", "Material / Produto",
                          "Qtd", "Unidade", "Preco Unit. (R$)", "Total (R$)", "Nota Fiscal", "Pagamento"]])
    ws.format("A1:J1", {
        "backgroundColor": {"red": 0.122, "green": 0.220, "blue": 0.392},
        "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}, "fontSize": 11},
        "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
    })

    # ── 3. Pré-formata 500 linhas (novas compras ficam formatadas automaticamente)
    ws.format(f"A2:J{MAX_DATA}", {
        "backgroundColor": {"red": 0.839, "green": 0.894, "blue": 0.941},
        "textFormat": {"fontSize": 10}, "verticalAlignment": "MIDDLE",
    })
    ws.format(f"H2:H{MAX_DATA}", {
        "backgroundColor": {"red": 0.886, "green": 0.937, "blue": 0.855},
        "textFormat": {"bold": True},
        "numberFormat": {"type": "CURRENCY", "pattern": "R$ #,##0.00"},
        "horizontalAlignment": "RIGHT",
    })
    ws.format(f"G2:G{MAX_DATA}", {
        "numberFormat": {"type": "CURRENCY", "pattern": "R$ #,##0.00"},
        "horizontalAlignment": "RIGHT",
    })
    ws.format(f"B2:B{MAX_DATA}", {
        "numberFormat": {"type": "DATE", "pattern": "dd/mm/yyyy"},
        "horizontalAlignment": "CENTER",
    })
    ws.format(f"A2:A{MAX_DATA}", {"horizontalAlignment": "CENTER"})
    ws.format(f"E2:F{MAX_DATA}", {"horizontalAlignment": "CENTER"})
    ws.format(f"I2:J{MAX_DATA}", {"horizontalAlignment": "CENTER"})

    ws.freeze(rows=1)

    # ── 4. Desmescla toda a área de dados + largura das colunas ──────────────
    sid        = ws.id
    col_widths = [45, 105, 185, 230, 55, 85, 140, 140, 125, 110]
    req = [
        # Desmescla qualquer célula mesclada na área de dados (resolve TOTAL antigo)
        {"unmergeCells": {
            "range": {
                "sheetId": sid,
                "startRowIndex": 1,       # linha 2 (0-based)
                "endRowIndex": MAX_DATA,  # até linha 500
                "startColumnIndex": 0,
                "endColumnIndex": 10,
            }
        }},
    ]
    req += [{"updateDimensionProperties": {
        "range": {"sheetId": sid, "dimension": "COLUMNS", "startIndex": i, "endIndex": i + 1},
        "properties": {"pixelSize": w}, "fields": "pixelSize",
    }} for i, w in enumerate(col_widths)]
    req.append({"updateDimensionProperties": {
        "range": {"sheetId": sid, "dimension": "ROWS", "startIndex": 0, "endIndex": 1},
        "properties": {"pixelSize": 42}, "fields": "pixelSize",
    }})
    spreadsheet.batch_update({"requests": req})

    # Nota: TOTAL GERAL não fica na aba principal para não conflitar com novos
    # dados. O total aparece nas abas "Por Fornecedor" e "Por Material".

    # ── 6. Abas de resumo ─────────────────────────────────────────────────────
    def _parse_float(v):
        try:
            s = str(v).strip().replace("R$", "").replace(" ", "")
            if "," in s and "." in s:
                s = s.replace(".", "").replace(",", ".")
            elif "," in s:
                s = s.replace(",", ".")
            return float(s)
        except (ValueError, TypeError):
            return 0.0

    def _cor(r, g, b):
        return {"red": r, "green": g, "blue": b}

    def _cell_fmt(sid, r1, c1, r2, c2, fmt):
        """Gera um repeatCell request (0-based, end exclusive)."""
        return {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": r1, "endRowIndex": r2,
                      "startColumnIndex": c1, "endColumnIndex": c2},
            "cell": {"userEnteredFormat": fmt},
            "fields": "userEnteredFormat(" + ",".join(fmt.keys()) + ")",
        }}

    def criar_resumo(nome, col_idx, titulo_col):
        try:
            spreadsheet.del_worksheet(spreadsheet.worksheet(nome))
        except Exception:
            pass
        s = spreadsheet.add_worksheet(nome, rows=200, cols=5)
        sid2 = s.id

        # Coleta e agrupa
        unicos = sorted(set(
            r[col_idx] for r in dados
            if len(r) > col_idx and str(r[col_idx]).strip()
        ))
        logger.info(f"[RESUMO] {nome}: {len(unicos)} itens únicos")

        rows_data  = []
        total_soma = 0.0
        for u in unicos:
            regs = [r for r in dados if len(r) > col_idx and r[col_idx] == u]
            tot  = sum(_parse_float(r[7]) for r in regs if len(r) > 7)
            total_soma += tot
            rows_data.append([u, len(regs), round(tot, 2)])

        rows_data.sort(key=lambda x: x[2], reverse=True)
        for row in rows_data:
            pct = (row[2] / total_soma * 100) if total_soma > 0 else 0
            row.append(round(pct, 1))

        n       = len(rows_data)
        end_row = 2 + n   # 1-based last data row
        t2      = end_row + 2

        # ── Escreve dados ────────────────────────────────────────────────────
        s.update("A1:E1", [[titulo_col, "", "", "", ""]])
        s.update("A2:E2", [[titulo_col, "Qtd. Compras", "Total Gasto (R$)", "Ranking", "% do Total"]])
        if rows_data:
            final_rows = [[r[0], r[1], r[2], f"{i+1}º", r[3]] for i, r in enumerate(rows_data)]
            s.update(f"A3:E{end_row}", final_rows, value_input_option="USER_ENTERED")
            s.update(f"A{t2}:E{t2}",
                     [["TOTAL GERAL", len(dados), round(total_soma, 2), "", "100%"]],
                     value_input_option="USER_ENTERED")

        # ── Toda formatação numa única chamada batch ──────────────────────────
        reqs = []

        # Título (linha 1)
        reqs.append({"mergeCells": {
            "range": {"sheetId": sid2, "startRowIndex": 0, "endRowIndex": 1,
                      "startColumnIndex": 0, "endColumnIndex": 5},
            "mergeType": "MERGE_ALL"
        }})
        reqs.append(_cell_fmt(sid2, 0, 0, 1, 5, {
            "backgroundColor": _cor(0.122, 0.220, 0.392),
            "textFormat": {"bold": True, "foregroundColor": _cor(1,1,1), "fontSize": 13},
            "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
        }))

        # Cabeçalho colunas (linha 2)
        reqs.append(_cell_fmt(sid2, 1, 0, 2, 5, {
            "backgroundColor": _cor(0.180, 0.459, 0.710),
            "textFormat": {"bold": True, "foregroundColor": _cor(1,1,1), "fontSize": 11},
            "horizontalAlignment": "CENTER",
        }))

        if rows_data:
            # Todas as linhas de dados — fundo claro
            reqs.append(_cell_fmt(sid2, 2, 0, end_row, 5, {
                "backgroundColor": _cor(0.839, 0.894, 0.941),
                "textFormat": {"fontSize": 10},
                "verticalAlignment": "MIDDLE",
            }))
            # Linhas pares — fundo levemente diferente
            for i in range(1, n, 2):
                reqs.append(_cell_fmt(sid2, 2+i, 0, 3+i, 5, {
                    "backgroundColor": _cor(0.922, 0.949, 0.973),
                }))
            # Coluna C (Total) — verde + moeda
            reqs.append(_cell_fmt(sid2, 2, 2, end_row, 3, {
                "backgroundColor": _cor(0.851, 0.918, 0.827),
                "textFormat": {"bold": True},
                "numberFormat": {"type": "CURRENCY", "pattern": "R$ #,##0.00"},
                "horizontalAlignment": "RIGHT",
            }))
            # Coluna B (Qtd) centralizada
            reqs.append(_cell_fmt(sid2, 2, 1, end_row, 2, {"horizontalAlignment": "CENTER"}))
            # Coluna D (Ranking) centralizada
            reqs.append(_cell_fmt(sid2, 2, 3, end_row, 4, {"horizontalAlignment": "CENTER"}))
            # Coluna E (%) centralizada
            reqs.append(_cell_fmt(sid2, 2, 4, end_row, 5, {
                "horizontalAlignment": "CENTER",
                "numberFormat": {"type": "NUMBER", "pattern": "0.0\"%\""},
            }))

            # TOTAL GERAL
            reqs.append({"mergeCells": {
                "range": {"sheetId": sid2, "startRowIndex": t2-1, "endRowIndex": t2,
                          "startColumnIndex": 0, "endColumnIndex": 2},
                "mergeType": "MERGE_ALL"
            }})
            reqs.append(_cell_fmt(sid2, t2-1, 0, t2, 5, {
                "backgroundColor": _cor(0.122, 0.220, 0.392),
                "textFormat": {"bold": True, "foregroundColor": _cor(1,1,1), "fontSize": 11},
                "horizontalAlignment": "CENTER",
            }))
            reqs.append(_cell_fmt(sid2, t2-1, 2, t2, 3, {
                "numberFormat": {"type": "CURRENCY", "pattern": "R$ #,##0.00"},
                "horizontalAlignment": "RIGHT",
            }))

        # Largura colunas + altura cabeçalho
        for i, w in enumerate([230, 110, 150, 80, 90]):
            reqs.append({"updateDimensionProperties": {
                "range": {"sheetId": sid2, "dimension": "COLUMNS",
                          "startIndex": i, "endIndex": i+1},
                "properties": {"pixelSize": w}, "fields": "pixelSize",
            }})
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": sid2, "dimension": "ROWS", "startIndex": 0, "endIndex": 1},
            "properties": {"pixelSize": 36}, "fields": "pixelSize",
        }})

        spreadsheet.batch_update({"requests": reqs})
        s.freeze(rows=2)

    # ── Aba Por Fornecedor (com breakdown por mês) ───────────────────────────
    def criar_resumo_fornecedor():
        from collections import defaultdict

        try:
            spreadsheet.del_worksheet(spreadsheet.worksheet("Por Fornecedor"))
        except Exception:
            pass
        s = spreadsheet.add_worksheet("Por Fornecedor", rows=300, cols=20)
        sid2 = s.id

        NOMES_MES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho",
                     "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]

        # Coleta fornecedores únicos (ordenados por gasto total)
        forn_totais = defaultdict(float)
        forn_qtd    = defaultdict(int)
        for r in dados:
            if len(r) > 7 and len(r) > 2 and str(r[2]).strip():
                forn_totais[r[2]] += _parse_float(r[7])
                forn_qtd[r[2]]    += 1
        fornecedores = sorted(forn_totais, key=lambda x: forn_totais[x], reverse=True)
        total_geral  = sum(forn_totais.values())
        n_forn       = len(fornecedores)

        # ── SEÇÃO 1: Resumo por fornecedor ───────────────────────────────────
        rows_res = []
        for i, f in enumerate(fornecedores):
            pct = round(forn_totais[f] / total_geral * 100, 1) if total_geral else 0
            rows_res.append([f, forn_qtd[f], round(forn_totais[f], 2), f"{i+1}º", pct])

        end_res = 2 + n_forn   # última linha de dados (1-based)
        tot_res = end_res + 2  # linha TOTAL GERAL (1-based)

        s.update("A1:E1", [["Fornecedor"]])
        s.update("A2:E2", [["Fornecedor", "Qtd. Compras", "Total Gasto (R$)", "Ranking", "% do Total"]])
        if rows_res:
            s.update(f"A3:E{end_res}", rows_res, value_input_option="USER_ENTERED")
            s.update(f"A{tot_res}:E{tot_res}",
                     [["TOTAL GERAL", sum(forn_qtd.values()), round(total_geral, 2), "", "100%"]],
                     value_input_option="USER_ENTERED")

        # ── SEÇÃO 2: Pivot mês × fornecedor ──────────────────────────────────
        # Descobre meses presentes (ordenados cronologicamente)
        import datetime as dt
        anos = set([dt.datetime.now().year])
        for r in dados:
            if len(r) > 1 and str(r[1]).strip():
                try:
                    anos.add(int(str(r[1]).strip().split("/")[2]))
                except Exception:
                    pass

        todos_meses = []
        for ano in sorted(anos):
            for mm in range(1, 13):
                todos_meses.append((ano, mm))

        # Agrega: pivot[mes_key][fornecedor] = total
        pivot = defaultdict(lambda: defaultdict(float))
        for r in dados:
            if len(r) > 7 and str(r[1]).strip() and str(r[2]).strip():
                try:
                    partes = str(r[1]).strip().split("/")
                    chave  = (int(partes[2]), int(partes[1]))
                    pivot[chave][r[2]] += _parse_float(r[7])
                except Exception:
                    pass

        # Linha de início da seção pivot (2 linhas após TOTAL GERAL)
        pivot_start = tot_res + 2   # 1-based

        n_mes   = len(todos_meses)
        n_cols  = 1 + n_forn + 1   # Mês + N fornecedores + Total Mês
        end_piv = pivot_start + 1 + n_mes  # última linha de dados pivot (1-based)
        tot_piv = end_piv + 1              # linha totais pivot (1-based)

        # Cabeçalho pivot
        cab_pivot = ["Mês"] + fornecedores + ["TOTAL MÊS"]
        s.update(f"A{pivot_start}:A{pivot_start}", [["GASTOS POR MÊS E FORNECEDOR"]])
        s.update(f"A{pivot_start+1}:{chr(64+n_cols)}{pivot_start+1}", [cab_pivot])

        # Dados pivot
        pivot_rows = []
        for (ano, mm) in todos_meses:
            nome_mes = f"{NOMES_MES[mm-1]}/{ano}"
            vals     = [pivot[(ano, mm)].get(f, 0.0) for f in fornecedores]
            total_m  = sum(vals)
            pivot_rows.append([nome_mes] + [round(v, 2) for v in vals] + [round(total_m, 2)])

        if pivot_rows:
            data_start = pivot_start + 2
            s.update(f"A{data_start}:{chr(64+n_cols)}{data_start + n_mes - 1}",
                     pivot_rows, value_input_option="USER_ENTERED")

        # Linha totais por fornecedor
        tot_cols = [round(forn_totais.get(f, 0), 2) for f in fornecedores]
        s.update(f"A{tot_piv}:{chr(64+n_cols)}{tot_piv}",
                 [["TOTAL GERAL"] + tot_cols + [round(total_geral, 2)]],
                 value_input_option="USER_ENTERED")

        # ── Formatação batch ──────────────────────────────────────────────────
        reqs = []
        ps   = pivot_start - 1   # 0-based pivot_start
        ds   = ps + 1             # 0-based cabeçalho pivot
        dd   = ds + 1             # 0-based início dados pivot
        tp   = tot_piv - 1        # 0-based linha totais pivot

        # --- Seção resumo ---
        reqs.append({"mergeCells": {
            "range": {"sheetId": sid2, "startRowIndex": 0, "endRowIndex": 1,
                      "startColumnIndex": 0, "endColumnIndex": 5},
            "mergeType": "MERGE_ALL"
        }})
        reqs.append(_cell_fmt(sid2, 0, 0, 1, 5, {
            "backgroundColor": _cor(0.122, 0.220, 0.392),
            "textFormat": {"bold": True, "foregroundColor": _cor(1,1,1), "fontSize": 13},
            "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
        }))
        reqs.append(_cell_fmt(sid2, 1, 0, 2, 5, {
            "backgroundColor": _cor(0.180, 0.459, 0.710),
            "textFormat": {"bold": True, "foregroundColor": _cor(1,1,1), "fontSize": 11},
            "horizontalAlignment": "CENTER",
        }))
        if rows_res:
            reqs.append(_cell_fmt(sid2, 2, 0, end_res, 5, {
                "backgroundColor": _cor(0.839, 0.894, 0.941),
                "textFormat": {"fontSize": 10}, "verticalAlignment": "MIDDLE",
            }))
            for i in range(1, n_forn, 2):
                reqs.append(_cell_fmt(sid2, 2+i, 0, 3+i, 5, {"backgroundColor": _cor(0.922, 0.949, 0.973)}))
            reqs.append(_cell_fmt(sid2, 2, 2, end_res, 3, {
                "backgroundColor": _cor(0.851, 0.918, 0.827),
                "textFormat": {"bold": True},
                "numberFormat": {"type": "CURRENCY", "pattern": "R$ #,##0.00"},
                "horizontalAlignment": "RIGHT",
            }))
            reqs.append(_cell_fmt(sid2, 2, 1, end_res, 2, {"horizontalAlignment": "CENTER"}))
            reqs.append(_cell_fmt(sid2, 2, 3, end_res, 4, {"horizontalAlignment": "CENTER"}))
            reqs.append(_cell_fmt(sid2, 2, 4, end_res, 5, {
                "horizontalAlignment": "CENTER",
                "numberFormat": {"type": "NUMBER", "pattern": "0.0\"%\""},
            }))
            reqs.append({"mergeCells": {
                "range": {"sheetId": sid2, "startRowIndex": tot_res-1, "endRowIndex": tot_res,
                          "startColumnIndex": 0, "endColumnIndex": 2},
                "mergeType": "MERGE_ALL"
            }})
            reqs.append(_cell_fmt(sid2, tot_res-1, 0, tot_res, 5, {
                "backgroundColor": _cor(0.122, 0.220, 0.392),
                "textFormat": {"bold": True, "foregroundColor": _cor(1,1,1), "fontSize": 11},
                "horizontalAlignment": "CENTER",
            }))
            reqs.append(_cell_fmt(sid2, tot_res-1, 2, tot_res, 3, {
                "numberFormat": {"type": "CURRENCY", "pattern": "R$ #,##0.00"},
                "horizontalAlignment": "RIGHT",
            }))

        # --- Seção pivot ---
        reqs.append({"mergeCells": {
            "range": {"sheetId": sid2, "startRowIndex": ps, "endRowIndex": ps+1,
                      "startColumnIndex": 0, "endColumnIndex": n_cols},
            "mergeType": "MERGE_ALL"
        }})
        reqs.append(_cell_fmt(sid2, ps, 0, ps+1, n_cols, {
            "backgroundColor": _cor(0.122, 0.220, 0.392),
            "textFormat": {"bold": True, "foregroundColor": _cor(1,1,1), "fontSize": 13},
            "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
        }))
        reqs.append(_cell_fmt(sid2, ds, 0, ds+1, n_cols, {
            "backgroundColor": _cor(0.180, 0.459, 0.710),
            "textFormat": {"bold": True, "foregroundColor": _cor(1,1,1), "fontSize": 11},
            "horizontalAlignment": "CENTER",
        }))
        if pivot_rows:
            reqs.append(_cell_fmt(sid2, dd, 0, dd+n_mes, n_cols, {
                "backgroundColor": _cor(0.839, 0.894, 0.941),
                "textFormat": {"fontSize": 10}, "verticalAlignment": "MIDDLE",
            }))
            for i in range(1, n_mes, 2):
                reqs.append(_cell_fmt(sid2, dd+i, 0, dd+i+1, n_cols, {"backgroundColor": _cor(0.922, 0.949, 0.973)}))
            # Valores numéricos — colunas 1..n_forn+1 (fornecedores + total mês)
            reqs.append(_cell_fmt(sid2, dd, 1, dd+n_mes, n_cols, {
                "numberFormat": {"type": "CURRENCY", "pattern": "R$ #,##0.00"},
                "horizontalAlignment": "RIGHT",
            }))
            # Coluna TOTAL MÊS em verde
            reqs.append(_cell_fmt(sid2, dd, n_cols-1, dd+n_mes, n_cols, {
                "backgroundColor": _cor(0.851, 0.918, 0.827),
                "textFormat": {"bold": True},
            }))
            # Linha TOTAL GERAL pivot
            reqs.append(_cell_fmt(sid2, tp, 0, tp+1, n_cols, {
                "backgroundColor": _cor(0.122, 0.220, 0.392),
                "textFormat": {"bold": True, "foregroundColor": _cor(1,1,1), "fontSize": 11},
                "horizontalAlignment": "CENTER",
            }))
            reqs.append(_cell_fmt(sid2, tp, 1, tp+1, n_cols, {
                "numberFormat": {"type": "CURRENCY", "pattern": "R$ #,##0.00"},
                "horizontalAlignment": "RIGHT",
            }))

        # Largura colunas: col A mais larga, demais por igual
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": sid2, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1},
            "properties": {"pixelSize": 200}, "fields": "pixelSize",
        }})
        for i in range(1, n_cols):
            reqs.append({"updateDimensionProperties": {
                "range": {"sheetId": sid2, "dimension": "COLUMNS", "startIndex": i, "endIndex": i+1},
                "properties": {"pixelSize": 140}, "fields": "pixelSize",
            }})
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": sid2, "dimension": "ROWS", "startIndex": 0, "endIndex": 1},
            "properties": {"pixelSize": 36}, "fields": "pixelSize",
        }})
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": sid2, "dimension": "ROWS", "startIndex": ps, "endIndex": ps+1},
            "properties": {"pixelSize": 36}, "fields": "pixelSize",
        }})

        spreadsheet.batch_update({"requests": reqs})
        s.freeze(rows=2)
        logger.info(f"[RESUMO] Por Fornecedor: {n_forn} fornecedores, {n_mes} meses")

    criar_resumo_fornecedor()
    time.sleep(5)
    criar_resumo("Por Material", 3, "Material / Produto")
    time.sleep(5)

    # ── Aba Pivot Fornecedor × Mês ────────────────────────────────────────────
    def criar_pivot_forn_mes():
        from collections import defaultdict
        import datetime as dt

        try:
            spreadsheet.del_worksheet(spreadsheet.worksheet("Forn. x Mes"))
        except Exception:
            pass
        s    = spreadsheet.add_worksheet("Forn. x Mes", rows=200, cols=20)
        sid2 = s.id

        NOMES_MES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho",
                     "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]

        # Fornecedores ordenados por gasto total
        forn_totais = defaultdict(float)
        for r in dados:
            if len(r) > 7 and len(r) > 2 and str(r[2]).strip():
                forn_totais[r[2]] += _parse_float(r[7])
        fornecedores = sorted(forn_totais, key=lambda x: forn_totais[x], reverse=True)
        n_forn       = len(fornecedores)

        # Meses (todos do(s) ano(s) presentes)
        anos = {dt.datetime.now().year}
        for r in dados:
            if len(r) > 1 and str(r[1]).strip():
                try:
                    anos.add(int(str(r[1]).strip().split("/")[2]))
                except Exception:
                    pass
        todos_meses = [(ano, mm) for ano in sorted(anos) for mm in range(1, 13)]
        n_mes       = len(todos_meses)

        # Pivot
        pivot = defaultdict(lambda: defaultdict(float))
        for r in dados:
            if len(r) > 7 and str(r[1]).strip() and str(r[2]).strip():
                try:
                    partes = str(r[1]).strip().split("/")
                    chave  = (int(partes[2]), int(partes[1]))
                    pivot[chave][r[2]] += _parse_float(r[7])
                except Exception:
                    pass

        n_cols = 1 + n_forn + 1   # Mês + fornecedores + TOTAL MÊS

        # ── Escreve dados ─────────────────────────────────────────────────────
        cab = ["Mês"] + fornecedores + ["TOTAL MÊS"]
        s.update("A1:A1", [["GASTOS POR MES E FORNECEDOR"]])
        s.update(f"A2:{chr(64+n_cols)}2", [cab])

        pivot_rows = []
        for (ano, mm) in todos_meses:
            nome = f"{NOMES_MES[mm-1]}/{ano}"
            vals = [round(pivot[(ano, mm)].get(f, 0.0), 2) for f in fornecedores]
            pivot_rows.append([nome] + vals + [round(sum(vals), 2)])

        if pivot_rows:
            s.update(f"A3:{chr(64+n_cols)}{2+n_mes}", pivot_rows, value_input_option="USER_ENTERED")

        # Linha TOTAL GERAL
        tot_row = 3 + n_mes + 1
        tot_vals = [round(forn_totais.get(f, 0), 2) for f in fornecedores]
        s.update(f"A{tot_row}:{chr(64+n_cols)}{tot_row}",
                 [["TOTAL GERAL"] + tot_vals + [round(sum(tot_vals), 2)]],
                 value_input_option="USER_ENTERED")

        # ── Formatação batch ──────────────────────────────────────────────────
        reqs = []

        # Título linha 1
        reqs.append({"mergeCells": {
            "range": {"sheetId": sid2, "startRowIndex": 0, "endRowIndex": 1,
                      "startColumnIndex": 0, "endColumnIndex": n_cols},
            "mergeType": "MERGE_ALL"
        }})
        reqs.append(_cell_fmt(sid2, 0, 0, 1, n_cols, {
            "backgroundColor": _cor(0.122, 0.220, 0.392),
            "textFormat": {"bold": True, "foregroundColor": _cor(1,1,1), "fontSize": 13},
            "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
        }))
        # Cabeçalho linha 2
        reqs.append(_cell_fmt(sid2, 1, 0, 2, n_cols, {
            "backgroundColor": _cor(0.180, 0.459, 0.710),
            "textFormat": {"bold": True, "foregroundColor": _cor(1,1,1), "fontSize": 11},
            "horizontalAlignment": "CENTER",
        }))
        if pivot_rows:
            # Dados — fundo claro com zebra
            reqs.append(_cell_fmt(sid2, 2, 0, 2+n_mes, n_cols, {
                "backgroundColor": _cor(0.839, 0.894, 0.941),
                "textFormat": {"fontSize": 10}, "verticalAlignment": "MIDDLE",
            }))
            for i in range(1, n_mes, 2):
                reqs.append(_cell_fmt(sid2, 2+i, 0, 3+i, n_cols, {"backgroundColor": _cor(0.922, 0.949, 0.973)}))
            # Colunas numéricas (B em diante)
            reqs.append(_cell_fmt(sid2, 2, 1, 2+n_mes, n_cols, {
                "numberFormat": {"type": "CURRENCY", "pattern": "R$ #,##0.00"},
                "horizontalAlignment": "RIGHT",
            }))
            # Última coluna (TOTAL MÊS) em verde
            reqs.append(_cell_fmt(sid2, 2, n_cols-1, 2+n_mes, n_cols, {
                "backgroundColor": _cor(0.851, 0.918, 0.827),
                "textFormat": {"bold": True},
                "numberFormat": {"type": "CURRENCY", "pattern": "R$ #,##0.00"},
                "horizontalAlignment": "RIGHT",
            }))
            # Linha TOTAL GERAL
            tp = tot_row - 1
            reqs.append(_cell_fmt(sid2, tp, 0, tp+1, n_cols, {
                "backgroundColor": _cor(0.122, 0.220, 0.392),
                "textFormat": {"bold": True, "foregroundColor": _cor(1,1,1), "fontSize": 11},
                "horizontalAlignment": "CENTER",
            }))
            reqs.append(_cell_fmt(sid2, tp, 1, tp+1, n_cols, {
                "numberFormat": {"type": "CURRENCY", "pattern": "R$ #,##0.00"},
                "horizontalAlignment": "RIGHT",
            }))

        # Largura: col A = 160px, demais = 140px
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": sid2, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1},
            "properties": {"pixelSize": 160}, "fields": "pixelSize",
        }})
        for i in range(1, n_cols):
            reqs.append({"updateDimensionProperties": {
                "range": {"sheetId": sid2, "dimension": "COLUMNS", "startIndex": i, "endIndex": i+1},
                "properties": {"pixelSize": 140}, "fields": "pixelSize",
            }})
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": sid2, "dimension": "ROWS", "startIndex": 0, "endIndex": 1},
            "properties": {"pixelSize": 40}, "fields": "pixelSize",
        }})

        spreadsheet.batch_update({"requests": reqs})
        s.freeze(rows=2)
        logger.info(f"[RESUMO] Forn. x Mes: {n_forn} fornecedores × {n_mes} meses")

    time.sleep(5)
    criar_pivot_forn_mes()

    # ── Aba Por Mês ───────────────────────────────────────────────────────────
    def criar_resumo_mes():
        from collections import defaultdict
        import datetime as dt

        try:
            spreadsheet.del_worksheet(spreadsheet.worksheet("Por Mes"))
        except Exception:
            pass
        s = spreadsheet.add_worksheet("Por Mes", rows=100, cols=3)

        NOMES_MES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho",
                     "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]

        # Descobre anos presentes nos dados (+ ano atual)
        anos = set()
        anos.add(dt.datetime.now().year)
        for r in dados:
            if len(r) > 1 and str(r[1]).strip():
                try:
                    anos.add(int(str(r[1]).strip().split("/")[2]))
                except Exception:
                    pass

        # Agrupa por mm/yyyy
        meses_data = defaultdict(lambda: {"qtd": 0, "total": 0.0})
        for r in dados:
            if len(r) > 7 and str(r[1]).strip():
                try:
                    partes = str(r[1]).strip().split("/")
                    chave  = (int(partes[2]), int(partes[1]))  # (yyyy, mm)
                    meses_data[chave]["qtd"]   += 1
                    meses_data[chave]["total"] += _parse_float(r[7])
                except Exception:
                    pass

        # Monta lista completa: todos os meses de cada ano presente
        todas_linhas = []
        for ano in sorted(anos):
            for mm in range(1, 13):
                chave  = (ano, mm)
                v      = meses_data.get(chave, {"qtd": 0, "total": 0.0})
                nome   = f"{NOMES_MES[mm-1]}/{ano}"
                todas_linhas.append([nome, round(v["total"], 2), v["qtd"]])

        n        = len(todas_linhas)
        last_data = 2 + n   # 0-based last data row (exclusive)
        total_row = last_data + 1  # 0-based total row

        # Escreve dados
        s.update("A1:C1", [["RESUMO POR MES"]])
        s.update("A2:C2", [["Mês", "Total (R$)", "Nº Lançamentos"]])
        if todas_linhas:
            s.update(f"A3:C{2 + n}", todas_linhas, value_input_option="USER_ENTERED")
            total_soma = sum(r[1] for r in todas_linhas)
            total_qtd  = sum(r[2] for r in todas_linhas)
            s.update(f"A{total_row + 1}:C{total_row + 1}",
                     [["TOTAL GERAL", round(total_soma, 2), total_qtd]],
                     value_input_option="USER_ENTERED")

        # Toda formatação numa única chamada batch
        sid2 = s.id
        reqs = []

        # Título
        reqs.append({"mergeCells": {
            "range": {"sheetId": sid2, "startRowIndex": 0, "endRowIndex": 1,
                      "startColumnIndex": 0, "endColumnIndex": 3},
            "mergeType": "MERGE_ALL"
        }})
        reqs.append(_cell_fmt(sid2, 0, 0, 1, 3, {
            "backgroundColor": _cor(0.122, 0.220, 0.392),
            "textFormat": {"bold": True, "foregroundColor": _cor(1,1,1), "fontSize": 13},
            "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
        }))
        # Cabeçalho colunas
        reqs.append(_cell_fmt(sid2, 1, 0, 2, 3, {
            "backgroundColor": _cor(0.180, 0.459, 0.710),
            "textFormat": {"bold": True, "foregroundColor": _cor(1,1,1), "fontSize": 11},
            "horizontalAlignment": "CENTER",
        }))

        if todas_linhas:
            # Todas linhas de dados — fundo claro
            reqs.append(_cell_fmt(sid2, 2, 0, last_data, 3, {
                "backgroundColor": _cor(0.839, 0.894, 0.941),
                "textFormat": {"fontSize": 10}, "verticalAlignment": "MIDDLE",
            }))
            # Linhas pares — fundo alternado
            for i in range(1, n, 2):
                reqs.append(_cell_fmt(sid2, 2+i, 0, 3+i, 3, {
                    "backgroundColor": _cor(0.922, 0.949, 0.973),
                }))
            # Coluna B (Total) — moeda + negrito
            reqs.append(_cell_fmt(sid2, 2, 1, last_data, 2, {
                "numberFormat": {"type": "CURRENCY", "pattern": "R$ #,##0.00"},
                "horizontalAlignment": "RIGHT", "textFormat": {"bold": True},
            }))
            # Coluna C (Qtd) — centralizado
            reqs.append(_cell_fmt(sid2, 2, 2, last_data, 3, {"horizontalAlignment": "CENTER"}))

            # TOTAL GERAL
            reqs.append(_cell_fmt(sid2, total_row, 0, total_row+1, 3, {
                "backgroundColor": _cor(0.122, 0.220, 0.392),
                "textFormat": {"bold": True, "foregroundColor": _cor(1,1,1), "fontSize": 11},
                "horizontalAlignment": "CENTER",
            }))
            reqs.append(_cell_fmt(sid2, total_row, 1, total_row+1, 2, {
                "numberFormat": {"type": "CURRENCY", "pattern": "R$ #,##0.00"},
                "horizontalAlignment": "RIGHT",
            }))

        # Largura colunas + altura cabeçalho
        for i, w in enumerate([160, 160, 140]):
            reqs.append({"updateDimensionProperties": {
                "range": {"sheetId": sid2, "dimension": "COLUMNS",
                          "startIndex": i, "endIndex": i+1},
                "properties": {"pixelSize": w}, "fields": "pixelSize",
            }})
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": sid2, "dimension": "ROWS", "startIndex": 0, "endIndex": 1},
            "properties": {"pixelSize": 40}, "fields": "pixelSize",
        }})

        spreadsheet.batch_update({"requests": reqs})
        s.freeze(rows=2)
        logger.info(f"[RESUMO] Por Mes: {n} meses")

    criar_resumo_mes()


# ── Handlers de comandos ──────────────────────────────────────────────────────
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data[STEP] = S_IDLE
    await update.message.reply_text(
        "Ola! Sou seu bot de controle de materiais.\n\n"
        "Comandos:\n"
        "  /adicionar - Registrar nova compra\n"
        "  /resumo    - Ver total gasto\n"
        "  /ultimas   - Ver ultimas 5 compras\n"
        "  /formatar  - Formatar planilha e criar resumos\n"
        "  /limpar    - Apagar todos os dados\n"
        "  /cancelar  - Cancelar operacao atual",
        reply_markup=ReplyKeyboardRemove(),
    )


async def cancelar(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data[STEP]  = S_IDLE
    context.user_data[DADOS] = {}
    await update.message.reply_text("Operacao cancelada.", reply_markup=ReplyKeyboardRemove())


async def resumo_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Buscando dados...")
    try:
        await update.message.reply_text(gerar_resumo())
    except Exception as e:
        logger.error(e)
        await update.message.reply_text("Erro ao acessar a planilha.")


async def ultimas_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Buscando dados...")
    try:
        await update.message.reply_text(ultimos_registros(5))
    except Exception as e:
        logger.error(e)
        await update.message.reply_text("Erro ao acessar a planilha.")


async def resetar_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Reconstruindo estrutura da planilha... (pode levar 30 segundos)")
    try:
        n = resetar_estrutura()
        await update.message.reply_text(
            f"Estrutura reconstruída!\n"
            f"{n} registro(s) preservado(s).\n\n"
            "Use /formatar para aplicar a formatação visual."
        )
    except Exception as e:
        logger.error(f"Erro ao resetar: {e}", exc_info=True)
        await update.message.reply_text(f"Erro: {e}")


async def formatar_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Formatando planilha... (pode levar 30 segundos)")
    try:
        aplicar_formatacao()
        await update.message.reply_text(
            "Pronto!\n\nAbas criadas:\n  - Por Fornecedor\n  - Por Material\n\nAbra o Google Sheets para ver!"
        )
    except Exception as e:
        logger.error(f"Erro ao formatar: {e}", exc_info=True)
        await update.message.reply_text(f"Erro: {e}")


async def adicionar_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data[STEP]  = S_FORNECEDOR
    context.user_data[DADOS] = {}
    await update.message.reply_text(
        "Selecione o fornecedor (ou /cancelar para sair):",
        reply_markup=ReplyKeyboardMarkup(TECLADO_FORNECEDORES, one_time_keyboard=True, resize_keyboard=True),
    )


async def limpar_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data[STEP] = S_LIMPAR
    await update.message.reply_text(
        "ATENCAO: isso vai apagar TODOS os registros.\n\nDigite CONFIRMAR para continuar ou /cancelar para sair.",
        reply_markup=ReplyKeyboardRemove(),
    )


# ── Handler principal de texto ────────────────────────────────────────────────
async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    step = context.user_data.get(STEP, S_IDLE)
    text = update.message.text.strip()
    d    = context.user_data.setdefault(DADOS, {})

    logger.info(f"[STATE] step={step!r} text={text!r}")

    # ── LIMPAR ────────────────────────────────────────────────────────────────
    if step == S_LIMPAR:
        if text.upper() == "CONFIRMAR":
            await update.message.reply_text("Limpando planilha...")
            try:
                limpar_planilha()
                await update.message.reply_text("Planilha limpa! Todos os dados foram apagados.")
            except Exception as e:
                logger.error(f"Erro ao limpar: {e}", exc_info=True)
                await update.message.reply_text(f"Erro ao limpar: {e}")
        else:
            await update.message.reply_text("Operacao cancelada.")
        context.user_data[STEP] = S_IDLE
        return

    # ── FORNECEDOR ────────────────────────────────────────────────────────────
    if step == S_FORNECEDOR:
        if text == OUTRO_FORN:
            await update.message.reply_text("Digite o nome do fornecedor:", reply_markup=ReplyKeyboardRemove())
            return
        d["fornecedor"] = text
        context.user_data[STEP] = S_MATERIAL
        await update.message.reply_text(
            "Selecione o material / produto:",
            reply_markup=ReplyKeyboardMarkup(TECLADO_MATERIAIS, one_time_keyboard=True, resize_keyboard=True),
        )
        return

    # ── MATERIAL ──────────────────────────────────────────────────────────────
    if step == S_MATERIAL:
        if text == OUTRO_MAT:
            await update.message.reply_text("Digite o nome do material:", reply_markup=ReplyKeyboardRemove())
            return
        d["material"] = text
        context.user_data[STEP] = S_QUANTIDADE
        await update.message.reply_text(
            "Qual a quantidade? (use ponto para decimais, ex: 10.5)",
            reply_markup=ReplyKeyboardRemove(),
        )
        return

    # ── QUANTIDADE ────────────────────────────────────────────────────────────
    if step == S_QUANTIDADE:
        try:
            d["quantidade"] = float(text.replace(",", "."))
        except ValueError:
            await update.message.reply_text("Numero invalido. Digite novamente (ex: 10 ou 10.5):")
            return
        context.user_data[STEP] = S_UNIDADE
        await update.message.reply_text(
            "Qual a unidade?",
            reply_markup=ReplyKeyboardMarkup(TECLADO_UNIDADES, one_time_keyboard=True, resize_keyboard=True),
        )
        return

    # ── UNIDADE ───────────────────────────────────────────────────────────────
    if step == S_UNIDADE:
        d["unidade"] = text
        context.user_data[STEP] = S_PRECO
        await update.message.reply_text(
            "Qual o preco unitario? (R$) Ex: 42.50",
            reply_markup=ReplyKeyboardRemove(),
        )
        return

    # ── PRECO ─────────────────────────────────────────────────────────────────
    if step == S_PRECO:
        try:
            d["preco"] = float(text.replace(",", ".").replace("R$", "").replace(" ", ""))
        except ValueError:
            await update.message.reply_text("Preco invalido. Digite novamente (ex: 42.50):")
            return
        context.user_data[STEP] = S_NOTA_FISCAL
        await update.message.reply_text(
            "Numero da Nota Fiscal? (ou toque em 'Sem NF')",
            reply_markup=ReplyKeyboardMarkup([["Sem NF"]], one_time_keyboard=True, resize_keyboard=True),
        )
        return

    # ── NOTA FISCAL ───────────────────────────────────────────────────────────
    if step == S_NOTA_FISCAL:
        d["nota_fiscal"] = "-" if text == "Sem NF" else text
        context.user_data[STEP] = S_PAGAMENTO
        await update.message.reply_text(
            "Metodo de pagamento?",
            reply_markup=ReplyKeyboardMarkup(TECLADO_PAGAMENTO, one_time_keyboard=True, resize_keyboard=True),
        )
        return

    # ── PAGAMENTO ─────────────────────────────────────────────────────────────
    if step == S_PAGAMENTO:
        d["pagamento"] = text
        d["data"]      = datetime.now()
        context.user_data[STEP] = S_CONFIRMAR
        total = d["quantidade"] * d["preco"]
        msg   = (
            "Confirme o registro:\n\n"
            f"  Fornecedor:  {d['fornecedor']}\n"
            f"  Material:    {d['material']}\n"
            f"  Quantidade:  {d['quantidade']} {d.get('unidade','')}\n"
            f"  Preco unit:  R$ {d['preco']:,.2f}\n"
            f"  TOTAL:       R$ {total:,.2f}\n"
            f"  Nota Fiscal: {d['nota_fiscal']}\n"
            f"  Pagamento:   {d['pagamento']}\n\n"
            "Digite SIM para salvar ou NAO para cancelar."
        )
        await update.message.reply_text(
            msg,
            reply_markup=ReplyKeyboardMarkup([["SIM", "NAO"]], one_time_keyboard=True, resize_keyboard=True),
        )
        return

    # ── CONFIRMAR ─────────────────────────────────────────────────────────────
    if step == S_CONFIRMAR:
        if text.upper() in ("SIM", "S"):
            try:
                num   = salvar_registro(d)
                total = d["quantidade"] * d["preco"]
                await update.message.reply_text(
                    f"Compra #{num} salva!\nTotal: R$ {total:,.2f}",
                    reply_markup=ReplyKeyboardRemove(),
                )
            except Exception as e:
                logger.error(f"Erro ao salvar: {e}", exc_info=True)
                await update.message.reply_text(f"Erro ao salvar: {e}", reply_markup=ReplyKeyboardRemove())
        else:
            await update.message.reply_text("Registro cancelado.", reply_markup=ReplyKeyboardRemove())
        context.user_data[STEP]  = S_IDLE
        context.user_data[DADOS] = {}
        return

    # ── IDLE (nenhum passo ativo) ─────────────────────────────────────────────
    await update.message.reply_text(
        "Use /adicionar para registrar uma compra ou /ajuda para ver os comandos."
    )


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    if not BOT_TOKEN:
        raise RuntimeError("BOT_TOKEN nao configurado!")

    persistence = PicklePersistence(filepath="/tmp/bot_state")
    app = Application.builder().token(BOT_TOKEN).persistence(persistence).build()

    WEBHOOK_DOMAIN = os.environ.get("RAILWAY_PUBLIC_DOMAIN", "").strip()

    app.add_handler(CommandHandler("start",     start))
    app.add_handler(CommandHandler("ajuda",     start))
    app.add_handler(CommandHandler("cancelar",  cancelar))
    app.add_handler(CommandHandler("adicionar", adicionar_cmd))
    app.add_handler(CommandHandler("limpar",    limpar_cmd))
    app.add_handler(CommandHandler("resumo",    resumo_cmd))
    app.add_handler(CommandHandler("ultimas",   ultimas_cmd))
    app.add_handler(CommandHandler("formatar",  formatar_cmd))
    app.add_handler(CommandHandler("resetar",   resetar_cmd))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))

    async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE):
        logger.error("Erro:", exc_info=context.error)
        if isinstance(update, Update) and update.effective_message:
            await update.effective_message.reply_text(f"Erro interno: {context.error}")

    app.add_error_handler(error_handler)

    logger.info("Bot rodando...")
    if WEBHOOK_DOMAIN:
        PORT = int(os.environ.get("PORT", 8080))
        logger.info(f"Modo WEBHOOK: {WEBHOOK_DOMAIN} porta {PORT}")
        app.run_webhook(
            listen="0.0.0.0",
            port=PORT,
            url_path=BOT_TOKEN,
            webhook_url=f"https://{WEBHOOK_DOMAIN}/{BOT_TOKEN}",
            drop_pending_updates=True,
        )
    else:
        logger.info("Modo POLLING (local)")
        app.run_polling(allowed_updates=Update.ALL_TYPES, drop_pending_updates=True)


if __name__ == "__main__":
    main()

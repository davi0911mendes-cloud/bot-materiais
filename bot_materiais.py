"""
Bot Telegram – Controle de Gastos com Materiais
================================================
Armazenamento: Google Sheets (via gspread)
Deploy:        Railway (nuvem)
"""

import os
import json
import logging
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
def _get_sheet():
    creds_str      = os.environ.get("GOOGLE_CREDENTIALS", "").strip()
    spreadsheet_id = os.environ.get("SPREADSHEET_ID", "").strip()
    if not creds_str:
        raise ValueError("GOOGLE_CREDENTIALS vazia!")
    gc          = gspread.service_account_from_dict(json.loads(creds_str))
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
    sheet  = _get_sheet()
    total  = dados["quantidade"] * dados["preco"]
    numero = max(len(sheet.get_all_values()) - 1, 0) + 1
    linha  = [
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


def limpar_planilha():
    creds_str      = os.environ.get("GOOGLE_CREDENTIALS", "").strip()
    spreadsheet_id = os.environ.get("SPREADSHEET_ID", "").strip()
    gc             = gspread.service_account_from_dict(json.loads(creds_str))
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
    MAX_DATA = 1000   # suporta até 1000 compras
    T_ROW    = 1002   # linha fixa do TOTAL GERAL

    creds_str      = os.environ.get("GOOGLE_CREDENTIALS", "").strip()
    spreadsheet_id = os.environ.get("SPREADSHEET_ID", "").strip()
    gc             = gspread.service_account_from_dict(json.loads(creds_str))
    spreadsheet    = gc.open_by_key(spreadsheet_id)

    try:
        ws = spreadsheet.worksheet("Registro de Compras")
    except gspread.WorksheetNotFound:
        raise ValueError("Aba 'Registro de Compras' nao encontrada.")

    # ── 1. Lê dados reais (linhas com data na coluna B, sem TOTAL) ────────────
    all_rows = ws.get_all_values()
    dados = [
        r for r in all_rows[1:]
        if len(r) > 1 and str(r[1]).strip() and "TOTAL" not in str(r[0]).upper()
    ]
    logger.info(f"[FORMATAR] {len(dados)} registros encontrados")

    # ── 2. Cabeçalho ──────────────────────────────────────────────────────────
    ws.update("A1:J1", [["#", "Data", "Fornecedor", "Material / Produto",
                          "Qtd", "Unidade", "Preco Unit. (R$)", "Total (R$)", "Nota Fiscal", "Pagamento"]])
    ws.format("A1:J1", {
        "backgroundColor": {"red": 0.122, "green": 0.220, "blue": 0.392},
        "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}, "fontSize": 11},
        "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
    })

    # ── 3. Pré-formata 1000 linhas de dados de uma vez ────────────────────────
    #       (novas compras adicionadas depois já ficam formatadas automaticamente)
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

    # ── 4. Largura das colunas e altura do cabeçalho ──────────────────────────
    sid        = ws.id
    col_widths = [45, 105, 185, 230, 55, 85, 140, 140, 125, 110]
    req = [{"updateDimensionProperties": {
        "range": {"sheetId": sid, "dimension": "COLUMNS", "startIndex": i, "endIndex": i + 1},
        "properties": {"pixelSize": w}, "fields": "pixelSize",
    }} for i, w in enumerate(col_widths)]
    req.append({"updateDimensionProperties": {
        "range": {"sheetId": sid, "dimension": "ROWS", "startIndex": 0, "endIndex": 1},
        "properties": {"pixelSize": 42}, "fields": "pixelSize",
    }})
    spreadsheet.batch_update({"requests": req})

    # ── 5. TOTAL GERAL na linha fixa T_ROW (1002) ─────────────────────────────
    #       Soma calculada em Python → sem depender de locale ou fórmulas
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

    total_geral = sum(_parse_float(r[7]) for r in dados if len(r) > 7)

    # Limpa a linha inteira antes de escrever
    ws.update(f"A{T_ROW}:J{T_ROW}", [["", "", "", "", "", "", "", "", "", ""]])
    ws.update(f"A{T_ROW}:A{T_ROW}", [["TOTAL GERAL"]])
    ws.merge_cells(f"A{T_ROW}:G{T_ROW}")
    ws.format(f"A{T_ROW}:G{T_ROW}", {
        "backgroundColor": {"red": 0.122, "green": 0.220, "blue": 0.392},
        "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}, "fontSize": 12},
        "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
    })
    ws.update(f"H{T_ROW}:H{T_ROW}", [[round(total_geral, 2)]], value_input_option="USER_ENTERED")
    ws.format(f"H{T_ROW}:J{T_ROW}", {
        "backgroundColor": {"red": 0.180, "green": 0.459, "blue": 0.710},
    })
    ws.format(f"H{T_ROW}:H{T_ROW}", {
        "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}, "fontSize": 12},
        "numberFormat": {"type": "CURRENCY", "pattern": "R$ #,##0.00"},
        "horizontalAlignment": "RIGHT",
    })

    # ── 6. Abas de resumo ─────────────────────────────────────────────────────
    def criar_resumo(nome, col_idx, titulo_col):
        try:
            spreadsheet.del_worksheet(spreadsheet.worksheet(nome))
        except Exception:
            pass
        s = spreadsheet.add_worksheet(nome, rows=500, cols=3)

        unicos = sorted(set(
            r[col_idx] for r in dados
            if len(r) > col_idx and str(r[col_idx]).strip()
        ))
        logger.info(f"[RESUMO] {nome}: {len(unicos)} itens únicos")

        # Título e cabeçalhos
        s.update("A1:A1", [[titulo_col]])
        s.merge_cells("A1:C1")
        s.format("A1:C1", {
            "backgroundColor": {"red": 0.122, "green": 0.220, "blue": 0.392},
            "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}, "fontSize": 13},
            "horizontalAlignment": "CENTER",
        })
        s.update("A2:C2", [[titulo_col, "Qtd. Compras", "Total Gasto (R$)"]])
        s.format("A2:C2", {
            "backgroundColor": {"red": 0.180, "green": 0.459, "blue": 0.710},
            "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}, "fontSize": 11},
            "horizontalAlignment": "CENTER",
        })

        # Calcula linhas de dados
        rows_data   = []
        total_soma  = 0.0
        for u in unicos:
            regs = [r for r in dados if len(r) > col_idx and r[col_idx] == u]
            tot  = sum(_parse_float(r[7]) for r in regs if len(r) > 7)
            total_soma += tot
            rows_data.append([u, len(regs), round(tot, 2)])

        if rows_data:
            n       = len(rows_data)
            end_row = 2 + n          # linha final dos dados (ex: n=2 → end_row=4)
            s.update(f"A3:C{end_row}", rows_data, value_input_option="USER_ENTERED")
            s.format(f"A3:C{end_row}", {"textFormat": {"fontSize": 10}, "verticalAlignment": "MIDDLE"})
            s.format(f"C3:C{end_row}", {
                "backgroundColor": {"red": 0.886, "green": 0.937, "blue": 0.855},
                "textFormat": {"bold": True},
                "numberFormat": {"type": "CURRENCY", "pattern": "R$ #,##0.00"},
                "horizontalAlignment": "RIGHT",
            })
            # TOTAL na linha end_row + 2
            t2 = end_row + 2
            s.update(f"A{t2}:C{t2}", [["TOTAL", "", round(total_soma, 2)]], value_input_option="USER_ENTERED")
            s.merge_cells(f"A{t2}:B{t2}")
            s.format(f"A{t2}:C{t2}", {
                "backgroundColor": {"red": 0.122, "green": 0.220, "blue": 0.392},
                "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}, "fontSize": 12},
                "numberFormat": {"type": "CURRENCY", "pattern": "R$ #,##0.00"},
                "horizontalAlignment": "RIGHT",
            })

        # Largura das colunas do resumo
        sid2 = s.id
        spreadsheet.batch_update({"requests": [
            {"updateDimensionProperties": {"range": {"sheetId": sid2, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1}, "properties": {"pixelSize": 230}, "fields": "pixelSize"}},
            {"updateDimensionProperties": {"range": {"sheetId": sid2, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 2}, "properties": {"pixelSize": 120}, "fields": "pixelSize"}},
            {"updateDimensionProperties": {"range": {"sheetId": sid2, "dimension": "COLUMNS", "startIndex": 2, "endIndex": 3}, "properties": {"pixelSize": 160}, "fields": "pixelSize"}},
        ]})
        s.freeze(rows=2)

    criar_resumo("Por Fornecedor", 2, "Fornecedor")
    criar_resumo("Por Material",   3, "Material / Produto")


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

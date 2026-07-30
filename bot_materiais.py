"""
Bot Telegram – Controle de Gastos com Materiais
================================================
Armazenamento: Google Sheets (via gspread)
Deploy:        Railway (nuvem) ou local

Dependências:
    pip install python-telegram-bot gspread google-auth

Variáveis de ambiente necessárias:
    BOT_TOKEN           – token do BotFather
    GOOGLE_CREDENTIALS  – conteúdo JSON da chave da conta de serviço (Google Cloud)
    SPREADSHEET_ID      – ID da planilha Google (da URL)
"""

import os
import json
import logging
from datetime import datetime

import gspread
from telegram import Update, ReplyKeyboardMarkup, ReplyKeyboardRemove
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    ConversationHandler, ContextTypes, filters,
)

# ── Configuração ────────────────────────────────────────────────────────────
BOT_TOKEN      = os.environ.get("BOT_TOKEN", "")
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "")
GOOGLE_CREDS   = os.environ.get("GOOGLE_CREDENTIALS", "")


logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

# Estados do fluxo de conversa
FORNECEDOR, MATERIAL, QUANTIDADE, UNIDADE, PRECO, NOTA_FISCAL = range(6)

# ── Listas padrão ────────────────────────────────────────────────────────────
OUTRO_FORN = "✏️ Outro fornecedor"
OUTRO_MAT  = "✏️ Outro material"

TECLADO_FORNECEDORES = [
    ["LISBOA", "MADECENTER"],
    ["LEO MADEIRAS", "VERDMADE"],
    ["CENCOMAL", "MADEREIRAS EXTRAS"],
    ["FGV", "HARDT"],
    ["HD FERRAGENS", "HAYD FERRAGENS"],
    ["ALTAPE FILMES E FITAS", "KILDERY THINNER"],
    ["PEQUENOS FORNECEDORES VARIÁVEIS"],
    [OUTRO_FORN],
]

TECLADO_MATERIAIS = [
    ["CHAPAS UNICOLOR 18MM", "CHAPAS MADEIRADO 18MM"],
    ["CHAPAS UNICOLOR 15MM", "CHAPAS MADEIRADO 15MM"],
    ["CHAPAS BRANCO 18MM", "CHAPAS BRANCO 15MM"],
    ["CHAPAS BRANCO 6MM", "CHAPAS UNICOLOR 6MM"],
    ["CHAPAS MADEIRADO 6MM", "FITA DE BORDA BRANCA 0,45"],
    ["FITA DE BORDA COLORIDA 0,45", "FITA DE BORDA BRANCA 1MM"],
    ["FITA DE BORDA COLORIDA 1MM", "CORREDIÇA INVISIVEL"],
    ["CORREDIÇA TELESCOPIA", "DOBRADIÇA CURVA"],
    ["DOBRADIÇA RETA", "COLA FORMICA"],
    ["COLA EXPANSIVA", "COLA PUR COLADEIRA"],
    ["COLA INSTANTANEA", "PARAFUSOS"],
    ["MINIFIX", "CAVILHA"],
    ["TAMBOR", "PIVO DE PORTA"],
    ["THINNER", "ALCOOL E VASELINA"],
    ["ESTOPA", OUTRO_MAT],
]


# ── Google Sheets ────────────────────────────────────────────────────────────
def _get_sheet():
    """Conecta ao Google Sheets e retorna a aba 'Registro de Compras'."""
    # Lê a variável fresca a cada chamada (evita cache de startup)
    creds_str = os.environ.get("GOOGLE_CREDENTIALS", "").strip()
    logger.info(f"[DIAG] GOOGLE_CREDENTIALS: {len(creds_str)} chars | início: {creds_str[:30]!r}")
    if not creds_str:
        raise ValueError("GOOGLE_CREDENTIALS está vazia no Railway!")
    creds_dict = json.loads(creds_str)
    spreadsheet_id = os.environ.get("SPREADSHEET_ID", "").strip()
    logger.info(f"[DIAG] SPREADSHEET_ID: {spreadsheet_id!r}")
    # Usa service_account_from_dict (sem cache local) em vez de authorize()
    gc = gspread.service_account_from_dict(creds_dict)
    spreadsheet = gc.open_by_key(spreadsheet_id)

    # Cria a aba se não existir
    try:
        sheet = spreadsheet.worksheet("Registro de Compras")
    except gspread.WorksheetNotFound:
        sheet = spreadsheet.add_worksheet("Registro de Compras", rows=1000, cols=9)
        sheet.append_row(
            ["#", "Data", "Fornecedor", "Material / Produto",
             "Qtd", "Unidade", "Preço Unit. (R$)", "Total (R$)", "Nota Fiscal"],
            value_input_option="USER_ENTERED",
        )
    return sheet


def salvar_registro(dados: dict) -> int:
    """Salva uma nova linha na planilha Google Sheets."""
    sheet = _get_sheet()
    total = dados["quantidade"] * dados["preco"]
    numero = max(len(sheet.get_all_values()) - 1, 0) + 1  # ignora cabeçalho

    linha = [
        numero,
        dados["data"].strftime("%d/%m/%Y"),
        dados["fornecedor"],
        dados["material"],
        dados["quantidade"],
        dados["unidade"],
        dados["preco"],
        total,
        dados.get("nota_fiscal", "–"),
    ]
    sheet.append_row(linha, value_input_option="USER_ENTERED")
    return numero


def gerar_resumo() -> str:
    """Lê a planilha e monta texto de resumo."""
    sheet = _get_sheet()
    rows  = sheet.get_all_values()[1:]  # pula cabeçalho
    rows  = [r for r in rows if any(r)]

    if not rows:
        return "Nenhum registro encontrado ainda."

    total_geral    = 0.0
    por_fornecedor = {}
    registros      = 0

    for r in rows:
        try:
            total = float(str(r[7]).replace("R$", "").replace(",", ".").strip())
        except (ValueError, IndexError):
            total = 0.0
        fornecedor = r[2] if len(r) > 2 and r[2] else "–"
        total_geral += total
        por_fornecedor[fornecedor] = por_fornecedor.get(fornecedor, 0.0) + total
        registros += 1

    linhas = [f"📊 *Resumo Geral* — {registros} compra(s)\n",
              f"💰 *Total Gasto: R$ {total_geral:,.2f}*\n",
              "─────────────────────",
              "*Por Fornecedor:*"]
    for forn, val in sorted(por_fornecedor.items(), key=lambda x: -x[1]):
        linhas.append(f"  • {forn}: R$ {val:,.2f}")

    return "\n".join(linhas)


def aplicar_formatacao():
    """Formata a planilha e cria abas de resumo."""
    creds_str      = os.environ.get("GOOGLE_CREDENTIALS", "").strip()
    spreadsheet_id = os.environ.get("SPREADSHEET_ID", "").strip()
    gc             = gspread.service_account_from_dict(json.loads(creds_str))
    spreadsheet    = gc.open_by_key(spreadsheet_id)

    # ── Aba principal ────────────────────────────────────────────────────────
    try:
        ws = spreadsheet.worksheet("Registro de Compras")
    except gspread.WorksheetNotFound:
        raise ValueError("Aba 'Registro de Compras' não encontrada.")

    all_rows = ws.get_all_values()
    last_row = max(len(all_rows), 1)

    # Limpa linhas de TOTAL antigas (contém "TOTAL" na col A ou D)
    for r in sorted(range(2, last_row + 1), reverse=True):
        vals = ws.row_values(r)
        if vals and "TOTAL" in str(vals[0]).upper():
            ws.delete_rows(r)
    all_rows = ws.get_all_values()
    last_row = max(len(all_rows), 1)

    # Cabeçalho
    ws.update('A1:I1', [['#','Data','Fornecedor','Material / Produto',
                          'Qtd','Unidade','Preço Unit. (R$)','Total (R$)','Nota Fiscal']])
    ws.format('A1:I1', {
        'backgroundColor': {'red': 0.122, 'green': 0.220, 'blue': 0.392},
        'textFormat': {'bold': True, 'foregroundColor': {'red': 1, 'green': 1, 'blue': 1}, 'fontSize': 11},
        'horizontalAlignment': 'CENTER', 'verticalAlignment': 'MIDDLE'
    })

    # Dados
    if last_row > 1:
        ws.format(f'A2:I{last_row}', {
            'backgroundColor': {'red': 0.839, 'green': 0.894, 'blue': 0.941},
            'textFormat': {'fontSize': 10}, 'verticalAlignment': 'MIDDLE'
        })
        ws.format(f'H2:H{last_row}', {
            'backgroundColor': {'red': 0.886, 'green': 0.937, 'blue': 0.855},
            'textFormat': {'bold': True},
            'numberFormat': {'type': 'CURRENCY', 'pattern': 'R$ #,##0.00'},
            'horizontalAlignment': 'RIGHT'
        })
        ws.format(f'G2:G{last_row}', {
            'numberFormat': {'type': 'CURRENCY', 'pattern': 'R$ #,##0.00'},
            'horizontalAlignment': 'RIGHT'
        })
        ws.format(f'B2:B{last_row}', {'numberFormat': {'type': 'DATE', 'pattern': 'dd/mm/yyyy'}, 'horizontalAlignment': 'CENTER'})
        ws.format(f'A2:A{last_row}', {'horizontalAlignment': 'CENTER'})
        ws.format(f'E2:F{last_row}', {'horizontalAlignment': 'CENTER'})
        ws.format(f'I2:I{last_row}', {'horizontalAlignment': 'CENTER'})

    # Congela cabeçalho
    ws.freeze(rows=1)

    # Larguras e altura via batch_update
    sid = ws.id
    col_widths = [45, 105, 185, 230, 55, 85, 140, 140, 125]
    requests = [{'updateDimensionProperties': {
        'range': {'sheetId': sid, 'dimension': 'COLUMNS', 'startIndex': i, 'endIndex': i+1},
        'properties': {'pixelSize': w}, 'fields': 'pixelSize'
    }} for i, w in enumerate(col_widths)]
    requests.append({'updateDimensionProperties': {
        'range': {'sheetId': sid, 'dimension': 'ROWS', 'startIndex': 0, 'endIndex': 1},
        'properties': {'pixelSize': 42}, 'fields': 'pixelSize'
    }})

    # Linha de Total Geral
    if last_row > 1:
        t = last_row + 2
        ws.update(f'A{t}', [['💰 TOTAL GERAL']])
        ws.merge_cells(f'A{t}:G{t}')
        ws.format(f'A{t}:G{t}', {
            'backgroundColor': {'red': 0.122, 'green': 0.220, 'blue': 0.392},
            'textFormat': {'bold': True, 'foregroundColor': {'red': 1, 'green': 1, 'blue': 1}, 'fontSize': 12},
            'horizontalAlignment': 'CENTER', 'verticalAlignment': 'MIDDLE'
        })
        ws.update(f'H{t}', [[f'=SUM(H2:H{last_row})']], value_input_option='USER_ENTERED')
        ws.format(f'H{t}', {
            'backgroundColor': {'red': 0.180, 'green': 0.459, 'blue': 0.710},
            'textFormat': {'bold': True, 'foregroundColor': {'red': 1, 'green': 1, 'blue': 1}, 'fontSize': 12},
            'numberFormat': {'type': 'CURRENCY', 'pattern': 'R$ #,##0.00'},
            'horizontalAlignment': 'RIGHT'
        })
        ws.format(f'I{t}', {'backgroundColor': {'red': 0.180, 'green': 0.459, 'blue': 0.710}})

    spreadsheet.batch_update({'requests': requests})

    # ── Resumos ──────────────────────────────────────────────────────────────
    dados = ws.get_all_values()[1:]  # sem cabeçalho
    dados = [r for r in dados if r and r[1]]  # linhas com data

    def criar_resumo(nome, col_idx, titulo_col):
        try:
            old = spreadsheet.worksheet(nome)
            spreadsheet.del_worksheet(old)
        except Exception:
            pass
        s = spreadsheet.add_worksheet(nome, rows=200, cols=3)

        unicos = sorted(set(r[col_idx] for r in dados if r[col_idx]))

        s.update('A1', [[titulo_col]])
        s.format('A1:C1', {
            'backgroundColor': {'red': 0.122, 'green': 0.220, 'blue': 0.392},
            'textFormat': {'bold': True, 'foregroundColor': {'red': 1, 'green': 1, 'blue': 1}, 'fontSize': 13},
            'horizontalAlignment': 'CENTER'
        })
        s.merge_cells('A1:C1')

        s.update('A2:C2', [[titulo_col, 'Qtd. Compras', 'Total Gasto (R$)']])
        s.format('A2:C2', {
            'backgroundColor': {'red': 0.180, 'green': 0.459, 'blue': 0.710},
            'textFormat': {'bold': True, 'foregroundColor': {'red': 1, 'green': 1, 'blue': 1}, 'fontSize': 11},
            'horizontalAlignment': 'CENTER'
        })

        # Calcula totais em Python (evita problemas de locale com fórmulas pt-BR)
        def _parse_float(v):
            try:
                s_val = str(v).strip().replace('R$', '').replace(' ', '')
                if ',' in s_val and '.' in s_val:
                    s_val = s_val.replace('.', '').replace(',', '.')
                elif ',' in s_val:
                    s_val = s_val.replace(',', '.')
                return float(s_val)
            except (ValueError, TypeError):
                return 0.0

        rows_data = []
        total_geral = 0.0
        for u in unicos:
            regs = [r for r in dados if r[col_idx] == u]
            qtd  = len(regs)
            tot  = sum(_parse_float(r[7]) for r in regs)
            total_geral += tot
            rows_data.append([u, qtd, tot])

        if rows_data:
            s.update('A3', rows_data)
            n = len(rows_data)
            s.format(f'A3:C{n+2}', {'textFormat': {'fontSize': 10}, 'verticalAlignment': 'MIDDLE'})
            s.format(f'C3:C{n+2}', {
                'backgroundColor': {'red': 0.886, 'green': 0.937, 'blue': 0.855},
                'textFormat': {'bold': True},
                'numberFormat': {'type': 'CURRENCY', 'pattern': 'R$ #,##0.00'},
                'horizontalAlignment': 'RIGHT'
            })
            # Total
            t2 = n + 4
            s.update(f'A{t2}', [['TOTAL', '', total_geral]])
            s.merge_cells(f'A{t2}:B{t2}')
            s.format(f'A{t2}:C{t2}', {
                'backgroundColor': {'red': 0.122, 'green': 0.220, 'blue': 0.392},
                'textFormat': {'bold': True, 'foregroundColor': {'red': 1, 'green': 1, 'blue': 1}, 'fontSize': 12},
                'numberFormat': {'type': 'CURRENCY', 'pattern': 'R$ #,##0.00'},
                'horizontalAlignment': 'RIGHT'
            })

        sid2 = s.id
        spreadsheet.batch_update({'requests': [
            {'updateDimensionProperties': {'range': {'sheetId': sid2, 'dimension': 'COLUMNS', 'startIndex': 0, 'endIndex': 1}, 'properties': {'pixelSize': 230}, 'fields': 'pixelSize'}},
            {'updateDimensionProperties': {'range': {'sheetId': sid2, 'dimension': 'COLUMNS', 'startIndex': 1, 'endIndex': 2}, 'properties': {'pixelSize': 120}, 'fields': 'pixelSize'}},
            {'updateDimensionProperties': {'range': {'sheetId': sid2, 'dimension': 'COLUMNS', 'startIndex': 2, 'endIndex': 3}, 'properties': {'pixelSize': 160}, 'fields': 'pixelSize'}},
        ]})
        s.freeze(rows=2)

    try:
        criar_resumo('Por Fornecedor', 2, 'Fornecedor')
        logger.info("Aba Por Fornecedor criada.")
    except Exception as e:
        logger.error(f"Erro ao criar resumo fornecedor: {e}", exc_info=True)
        raise

    try:
        criar_resumo('Por Material', 3, 'Material / Produto')
        logger.info("Aba Por Material criada.")
    except Exception as e:
        logger.error(f"Erro ao criar resumo material: {e}", exc_info=True)
        raise


def limpar_planilha():
    """Apaga todos os dados da planilha, mantendo apenas o cabeçalho."""
    creds_str      = os.environ.get("GOOGLE_CREDENTIALS", "").strip()
    spreadsheet_id = os.environ.get("SPREADSHEET_ID", "").strip()
    gc             = gspread.service_account_from_dict(json.loads(creds_str))
    spreadsheet    = gc.open_by_key(spreadsheet_id)

    # Limpa aba principal (mantém só o cabeçalho)
    try:
        ws = spreadsheet.worksheet("Registro de Compras")
        last_row = ws.row_count
        if last_row > 1:
            ws.delete_rows(2, last_row)
    except gspread.WorksheetNotFound:
        pass

    # Remove abas de resumo
    for nome in ("Por Fornecedor", "Por Material"):
        try:
            spreadsheet.del_worksheet(spreadsheet.worksheet(nome))
        except Exception:
            pass


def ultimos_registros(n: int = 5) -> str:
    """Retorna os últimos N registros."""
    sheet = _get_sheet()
    rows  = sheet.get_all_values()[1:]
    rows  = [r for r in rows if any(r)]

    if not rows:
        return "Nenhum registro ainda."

    ultimas = rows[-n:][::-1]
    linhas  = [f"📋 *Últimas {len(ultimas)} compra(s):*\n"]
    for r in ultimas:
        try:
            total = float(str(r[7]).replace("R$", "").replace(",", ".").strip())
        except (ValueError, IndexError):
            total = 0.0
        linhas.append(f"• {r[1]} | {r[3]} | {r[4]} {r[5]} | R$ {total:,.2f}")

    return "\n".join(linhas)


# ── Handlers ─────────────────────────────────────────────────────────────────
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    texto = (
        "👋 *Olá! Sou seu bot de controle de materiais.*\n\n"
        "Comandos disponíveis:\n"
        "  /adicionar – Registrar nova compra\n"
        "  /resumo    – Ver total gasto\n"
        "  /ultimas   – Ver últimas 5 compras\n"
        "  /formatar  – Formatar planilha e criar resumos\n"
        "  /limpar    – Apagar todos os dados da planilha\n"
        "  /ajuda     – Mostrar esta mensagem"
    )
    await update.message.reply_text(texto, parse_mode="Markdown")


async def ajuda(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await start(update, context)


# ── Fluxo /adicionar ─────────────────────────────────────────────────────────
async def adicionar_inicio(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "🏭 *Selecione o fornecedor:*\n_(ou /cancelar para sair)_",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup(
            TECLADO_FORNECEDORES, one_time_keyboard=True, resize_keyboard=True
        ),
    )
    return FORNECEDOR


async def receber_fornecedor(update: Update, context: ContextTypes.DEFAULT_TYPE):
    texto = update.message.text.strip()
    if texto == OUTRO_FORN:
        await update.message.reply_text(
            "🏭 *Digite o nome do fornecedor:*",
            parse_mode="Markdown",
            reply_markup=ReplyKeyboardRemove(),
        )
        return FORNECEDOR
    context.user_data["fornecedor"] = texto
    await update.message.reply_text(
        "📦 *Selecione o material / produto:*",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup(
            TECLADO_MATERIAIS, one_time_keyboard=True, resize_keyboard=True
        ),
    )
    return MATERIAL


async def receber_material(update: Update, context: ContextTypes.DEFAULT_TYPE):
    texto = update.message.text.strip()
    if texto == OUTRO_MAT:
        await update.message.reply_text(
            "📦 *Digite o nome do material:*",
            parse_mode="Markdown",
            reply_markup=ReplyKeyboardRemove(),
        )
        return MATERIAL
    context.user_data["material"] = texto
    await update.message.reply_text(
        "🔢 *Qual a quantidade?*\n_(use ponto para decimais, ex: 10.5)_",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardRemove(),
    )
    return QUANTIDADE


async def receber_quantidade(update: Update, context: ContextTypes.DEFAULT_TYPE):
    texto = update.message.text.strip().replace(",", ".")
    try:
        context.user_data["quantidade"] = float(texto)
    except ValueError:
        await update.message.reply_text("❌ Número inválido. Digite novamente:")
        return QUANTIDADE

    teclado = [["kg", "g"], ["L", "mL"], ["m", "m²"], ["un", "cx"], ["sacos", "peças"]]
    await update.message.reply_text(
        "📏 *Qual a unidade?*\n_(escolha ou digite outra)_",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup(teclado, one_time_keyboard=True, resize_keyboard=True),
    )
    return UNIDADE


async def receber_unidade(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["unidade"] = update.message.text.strip()
    await update.message.reply_text(
        "💲 *Qual o preço unitário? (R$)*\n_(ex: 42.50)_",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardRemove(),
    )
    return PRECO


async def receber_preco(update: Update, context: ContextTypes.DEFAULT_TYPE):
    texto = update.message.text.strip().replace(",", ".").replace("R$", "").replace(" ", "")
    try:
        context.user_data["preco"] = float(texto)
    except ValueError:
        await update.message.reply_text("❌ Preço inválido. Digite novamente (ex: 42.50):")
        return PRECO

    await update.message.reply_text(
        "🧾 *Número da Nota Fiscal?*\n_(ou toque em 'Sem NF' se não tiver)_",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup(
            [["📋 Sem NF"]], one_time_keyboard=True, resize_keyboard=True
        ),
    )
    return NOTA_FISCAL


async def receber_nota_fiscal(update: Update, context: ContextTypes.DEFAULT_TYPE):
    texto = update.message.text.strip()
    context.user_data["nota_fiscal"] = "–" if texto == "📋 Sem NF" else texto
    context.user_data["data"] = datetime.now()

    d = context.user_data
    total = d["quantidade"] * d["preco"]
    nf    = d["nota_fiscal"]

    confirmacao = (
        "✅ *Confirme o registro:*\n\n"
        f"  🏭 Fornecedor: {d['fornecedor']}\n"
        f"  📦 Material:   {d['material']}\n"
        f"  🔢 Quantidade: {d['quantidade']} {d['unidade']}\n"
        f"  💲 Preço unit: R$ {d['preco']:,.2f}\n"
        f"  💰 *Total:      R$ {total:,.2f}*\n"
        f"  🧾 Nota Fiscal: {nf}\n\n"
        "Digite *sim* para salvar ou *não* para cancelar."
    )
    await update.message.reply_text(
        confirmacao,
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup(
            [["✅ Sim", "❌ Não"]], one_time_keyboard=True, resize_keyboard=True
        ),
    )
    return ConversationHandler.END


async def confirmar_e_salvar(update: Update, context: ContextTypes.DEFAULT_TYPE):
    resposta = update.message.text.strip().lower()
    if "sim" in resposta or "✅" in resposta:
        d = context.user_data
        try:
            num = salvar_registro(d)
            total = d["quantidade"] * d["preco"]
            await update.message.reply_text(
                f"✅ *Compra #{num} salva no Google Sheets!*\n"
                f"Total: R$ {total:,.2f} registrado com sucesso.",
                parse_mode="Markdown",
                reply_markup=ReplyKeyboardRemove(),
            )
        except Exception as e:
            logger.error(f"Erro ao salvar no Sheets: {e}")
            await update.message.reply_text(
                "❌ Erro ao salvar. Verifique as configurações do Google Sheets.",
                reply_markup=ReplyKeyboardRemove(),
            )
    else:
        await update.message.reply_text(
            "🚫 Registro cancelado.",
            reply_markup=ReplyKeyboardRemove(),
        )
    context.user_data.clear()
    return ConversationHandler.END


async def cancelar(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text("🚫 Operação cancelada.", reply_markup=ReplyKeyboardRemove())
    return ConversationHandler.END


async def resumo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("⏳ Buscando dados...", parse_mode="Markdown")
    try:
        texto = gerar_resumo()
    except Exception as e:
        logger.error(e)
        texto = "❌ Erro ao acessar a planilha."
    await update.message.reply_text(texto, parse_mode="Markdown")


async def ultimas(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("⏳ Buscando dados...", parse_mode="Markdown")
    try:
        texto = ultimos_registros(5)
    except Exception as e:
        logger.error(e)
        texto = "❌ Erro ao acessar a planilha."
    await update.message.reply_text(texto, parse_mode="Markdown")


async def limpar(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "⚠️ *Tem certeza?*\n\nIsso vai apagar *TODOS* os registros da planilha.\n\nDigite *CONFIRMAR* para continuar ou qualquer outra coisa para cancelar.",
        parse_mode="Markdown"
    )
    context.user_data["aguardando_confirmacao_limpar"] = True


async def confirmar_limpar(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not context.user_data.get("aguardando_confirmacao_limpar"):
        return
    context.user_data.pop("aguardando_confirmacao_limpar", None)

    if update.message.text.strip().upper() == "CONFIRMAR":
        await update.message.reply_text("⏳ Limpando planilha...")
        try:
            limpar_planilha()
            await update.message.reply_text(
                "✅ *Planilha limpa!*\n\nTodos os dados foram apagados. O cabeçalho foi mantido.",
                parse_mode="Markdown"
            )
        except Exception as e:
            logger.error(f"Erro ao limpar: {e}", exc_info=True)
            await update.message.reply_text(f"❌ Erro ao limpar: {e}")
    else:
        await update.message.reply_text("🚫 Operação cancelada.")


async def formatar(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "⏳ Formatando planilha e criando resumos...\n_(pode levar 30 segundos)_",
        parse_mode="Markdown"
    )
    try:
        aplicar_formatacao()
        await update.message.reply_text(
            "✅ *Pronto!*\n\n"
            "Abas criadas na planilha:\n"
            "  • *Por Fornecedor*\n"
            "  • *Por Material*\n\n"
            "Abra o Google Sheets para ver!",
            parse_mode="Markdown"
        )
    except Exception as e:
        logger.error(f"Erro ao formatar: {e}", exc_info=True)
        await update.message.reply_text(f"❌ Erro: {e}")


# ── Main ─────────────────────────────────────────────────────────────────────
def main():
    if not BOT_TOKEN:
        raise RuntimeError("BOT_TOKEN não configurado nas variáveis de ambiente!")
    if not GOOGLE_CREDS:
        raise RuntimeError("GOOGLE_CREDENTIALS não configurado!")
    if not SPREADSHEET_ID:
        raise RuntimeError("SPREADSHEET_ID não configurado!")

    app = Application.builder().token(BOT_TOKEN).build()

    conv = ConversationHandler(
        entry_points=[CommandHandler("adicionar", adicionar_inicio)],
        states={
            FORNECEDOR: [MessageHandler(filters.TEXT & ~filters.COMMAND, receber_fornecedor)],
            MATERIAL:   [MessageHandler(filters.TEXT & ~filters.COMMAND, receber_material)],
            QUANTIDADE:   [MessageHandler(filters.TEXT & ~filters.COMMAND, receber_quantidade)],
            UNIDADE:      [MessageHandler(filters.TEXT & ~filters.COMMAND, receber_unidade)],
            PRECO:        [MessageHandler(filters.TEXT & ~filters.COMMAND, receber_preco)],
            NOTA_FISCAL:  [MessageHandler(filters.TEXT & ~filters.COMMAND, receber_nota_fiscal)],
        },
        fallbacks=[CommandHandler("cancelar", cancelar)],
    )

    app.add_handler(MessageHandler(
        filters.Regex(r"^(✅ Sim|✅ sim|sim|Sim|❌ Não|não|nao|Não)$"),
        confirmar_e_salvar,
    ))
    app.add_handler(conv)
    app.add_handler(CommandHandler("start",    start))
    app.add_handler(CommandHandler("ajuda",    ajuda))
    app.add_handler(CommandHandler("resumo",   resumo))
    app.add_handler(CommandHandler("ultimas",  ultimas))
    app.add_handler(CommandHandler("formatar", formatar))
    app.add_handler(CommandHandler("limpar",   limpar))
    # Deve ficar por ÚLTIMO para não interceptar o ConversationHandler
    app.add_handler(MessageHandler(
        filters.TEXT & ~filters.COMMAND,
        confirmar_limpar,
    ))

    logger.info("🤖 Bot rodando...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()

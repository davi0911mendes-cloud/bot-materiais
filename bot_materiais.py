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
from google.oauth2.service_account import Credentials
from telegram import Update, ReplyKeyboardMarkup, ReplyKeyboardRemove
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    ConversationHandler, ContextTypes, filters,
)

# ── Configuração ────────────────────────────────────────────────────────────
BOT_TOKEN      = os.environ.get("BOT_TOKEN", "")
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "")
GOOGLE_CREDS   = os.environ.get("GOOGLE_CREDENTIALS", "")

SCOPES = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

# Estados do fluxo de conversa
FORNECEDOR, MATERIAL, QUANTIDADE, UNIDADE, PRECO = range(5)

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
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(SPREADSHEET_ID)

    # Cria a aba se não existir
    try:
        sheet = spreadsheet.worksheet("Registro de Compras")
    except gspread.WorksheetNotFound:
        sheet = spreadsheet.add_worksheet("Registro de Compras", rows=1000, cols=8)
        sheet.append_row(
            ["#", "Data", "Fornecedor", "Material / Produto",
             "Qtd", "Unidade", "Preço Unit. (R$)", "Total (R$)"],
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

    d = context.user_data
    total = d["quantidade"] * d["preco"]

    confirmacao = (
        "✅ *Confirme o registro:*\n\n"
        f"  🏭 Fornecedor: {d['fornecedor']}\n"
        f"  📦 Material:   {d['material']}\n"
        f"  🔢 Quantidade: {d['quantidade']} {d['unidade']}\n"
        f"  💲 Preço unit: R$ {d['preco']:,.2f}\n"
        f"  💰 *Total:      R$ {total:,.2f}*\n\n"
        "Digite *sim* para salvar ou *não* para cancelar."
    )
    context.user_data["data"] = datetime.now()
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
            QUANTIDADE: [MessageHandler(filters.TEXT & ~filters.COMMAND, receber_quantidade)],
            UNIDADE:    [MessageHandler(filters.TEXT & ~filters.COMMAND, receber_unidade)],
            PRECO:      [MessageHandler(filters.TEXT & ~filters.COMMAND, receber_preco)],
        },
        fallbacks=[CommandHandler("cancelar", cancelar)],
    )

    app.add_handler(MessageHandler(
        filters.Regex(r"^(✅ Sim|✅ sim|sim|Sim|❌ Não|não|nao|Não)$"),
        confirmar_e_salvar,
    ))
    app.add_handler(conv)
    app.add_handler(CommandHandler("start",   start))
    app.add_handler(CommandHandler("ajuda",   ajuda))
    app.add_handler(CommandHandler("resumo",  resumo))
    app.add_handler(CommandHandler("ultimas", ultimas))

    logger.info("🤖 Bot rodando...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()

"""
Bot de Telegram - Consolidado de Transacciones por Placa y Mes.

Flujo:
1. Usuario manda el Excel consolidado al bot.
2. El bot quita 'Contratos'.
3. Segmenta por Placa -> una hoja por placa.
4. Ordena por fecha de la transaccion.
5. Agrupa por mes y agrega AUTOSUMA (=SUM) por cada mes.
6. Devuelve el Excel ordenado.

El usuario puede elegir el orden con botones (mas reciente / mas antiguo).
"""
import os
import logging
import tempfile

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler, CallbackQueryHandler,
    ContextTypes, filters
)

from processor import procesar

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)
log = logging.getLogger(__name__)

TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")

# guarda temporalmente el archivo de cada usuario mientras elige el orden
PENDIENTES = {}


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Hola. Yo procesar consolidado.\n\n"
        "Manda archivo Excel (.xlsx).\n"
        "Yo hacer esto:\n"
        "- Quitar Contratos\n"
        "- Separar por placa (una hoja por placa)\n"
        "- Agrupar por mes\n"
        "- Poner autosuma por cada mes\n"
        "- Devolver Excel ordenado"
    )


async def recibir_archivo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    nombre = (doc.file_name or "").lower()
    if not (nombre.endswith(".xlsx") or nombre.endswith(".xlsm")):
        await update.message.reply_text("Esto malo. Manda archivo .xlsx por favor.")
        return

    await update.message.reply_text("Archivo recibido. Descargando...")

    tg_file = await doc.get_file()
    tmp_dir = tempfile.mkdtemp()
    in_path = os.path.join(tmp_dir, doc.file_name)
    await tg_file.download_to_drive(in_path)

    PENDIENTES[update.effective_user.id] = in_path

    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("Mes mas reciente primero", callback_data="reciente")],
        [InlineKeyboardButton("Mes mas antiguo primero", callback_data="antiguo")],
    ])
    await update.message.reply_text("Como quieres el orden?", reply_markup=keyboard)


async def elegir_orden(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    orden = query.data
    user_id = query.from_user.id

    in_path = PENDIENTES.get(user_id)
    if not in_path or not os.path.exists(in_path):
        await query.edit_message_text("Esto malo. No hay archivo. Manda Excel otra vez.")
        return

    await query.edit_message_text("Procesando... espera un momento.")

    try:
        out_path = os.path.join(os.path.dirname(in_path), "Consolidado_Procesado.xlsx")
        _, placas = procesar(in_path, out_path, orden=orden)

        with open(out_path, "rb") as f:
            await context.bot.send_document(
                chat_id=query.message.chat_id,
                document=f,
                filename="Consolidado_Procesado.xlsx",
                caption=f"Listo. Placas: {', '.join(placas)}\n"
                        f"Orden: {'mas reciente primero' if orden=='reciente' else 'mas antiguo primero'}\n"
                        f"Contratos: excluidos. Autosuma por mes: incluida."
            )
    except Exception as e:
        log.exception("Error procesando")
        await context.bot.send_message(
            chat_id=query.message.chat_id,
            text=f"Esto malo. Error: {e}"
        )
    finally:
        PENDIENTES.pop(user_id, None)


def main():
    if not TOKEN:
        raise RuntimeError("Falta variable TELEGRAM_BOT_TOKEN")

    app = Application.builder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", start))
    app.add_handler(MessageHandler(filters.Document.ALL, recibir_archivo))
    app.add_handler(CallbackQueryHandler(elegir_orden))

    log.info("Bot iniciado.")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()

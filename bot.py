import os
import json
import logging
import base64
import xml.etree.ElementTree as ET
from datetime import datetime
from anthropic import Anthropic
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
import urllib.request
from telegram import Update, ReplyKeyboardMarkup, KeyboardButton
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

ANTHROPIC_CLIENT = Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])
TELEGRAM_TOKEN = os.environ["TELEGRAM_TOKEN"]
ALLOWED_USER_ID = int(os.environ.get("ALLOWED_USER_ID", "0"))

DATA_FILE = "tickets.json"

SYSTEM_PROMPT = """Ты — ассистент по обработке авиабилетов. Пользователь присылает текст или изображение билета.

Извлеки данные и верни ТОЛЬКО валидный JSON без markdown и блоков кода:
{"action":"add","ticket":{"num":"номер","date":"YYYY-MM-DD","status":"Выписан","name":"ИМЯ ПАССАЖИРА","route":"XXX-XXX","company":"","price":0,"currency":"AZN","cu":0,"ca":0},"missing":[]}

Правила:
- status: "Выписан", "Изменён" или "Отменён". По умолчанию "Выписан"
- date: формат YYYY-MM-DD. Если не указана — сегодняшняя дата
- name: латиница ЗАГЛАВНЫМИ (SURNAME/FIRSTNAME)
- route: ЗАГЛАВНЫМИ, формат BAK-TBS или BAK-TBS-BAK
- price: итоговая сумма TOTAL из билета (в оригинальной валюте)
- currency: валюта цены билета — AZN, EUR, USD, RUB или KZT. Определи по символу или тексту в билете. По умолчанию AZN.
- cu: наша комиссия (всегда в AZN, если не указана — 0)
- ca: комиссия агента (всегда в AZN, если не указана — 0)
- company: компания-заказчик (если не указана — пустая строка)
- missing: список отсутствующих полей

Если это не билет или вопрос: {"action":"chat","text":"твой ответ на русском"}"""


def get_cbar_rates():
    try:
        url = "https://www.cbar.az/currencies/today.xml"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=10) as response:
            xml_data = response.read()
        root = ET.fromstring(xml_data)
        rates = {}
        for valute in root.findall(".//Valute"):
            code = valute.get("Code", "")
            nominal_el = valute.find("Nominal")
            value_el = valute.find("Value")
            if code and nominal_el is not None and value_el is not None:
                nominal = float(nominal_el.text)
                value = float(value_el.text.replace(",", "."))
                rates[code] = value / nominal
        logger.info(f"CBAR rates: {rates}")
        return rates
    except Exception as e:
        logger.error(f"CBAR error: {e}")
        return {"USD": 1.70, "EUR": 1.87, "RUB": 0.019, "KZT": 0.0035}


def convert_to_azn(amount, currency, rates):
    if currency == "AZN" or not currency:
        return amount, 1.0
    rate = rates.get(currency.upper(), 1.0)
    return round(amount * rate, 2), rate


def load_tickets():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def save_tickets(tickets):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(tickets, f, ensure_ascii=False, indent=2)


def fmt_date(d):
    if not d:
        return ""
    try:
        parts = d.split("-")
        return f"{parts[2]}.{parts[1]}.{parts[0]}"
    except:
        return d


def money(v):
    try:
        return f"{float(v):,.2f}"
    except:
        return "0.00"


async def parse_ticket_with_claude(text=None, image_data=None, image_mime=None):
    content = []
    if image_data:
        content.append({"type": "image", "source": {"type": "base64", "media_type": image_mime, "data": image_data}})
    if text:
        content.append({"type": "text", "text": text})
    elif image_data:
        content.append({"type": "text", "text": "Распознай данные билета с этого изображения"})
    response = ANTHROPIC_CLIENT.messages.create(
        model="claude-opus-4-5",
        max_tokens=1000,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": content}]
    )
    raw = response.content[0].text.strip().replace("```json", "").replace("```", "").strip()
    return json.loads(raw)


def generate_excel(tickets):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Билеты"
    headers = ["№", "Номер билета", "Дата", "Статус", "Пассажир", "Маршрут",
               "Компания", "Цена (ориг.)", "Валюта", "Курс", "Цена (AZN)",
               "Ком. наша", "Ком. агента", "Компания должна нам", "Мы должны агенту"]
    header_fill = PatternFill(start_color="2C2C2A", end_color="2C2C2A", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
    col_widths = [4, 18, 12, 11, 24, 14, 20, 12, 8, 8, 12, 12, 13, 20, 20]
    for i, width in enumerate(col_widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = width
    status_colors = {"Выписан": "E1F5EE", "Изменён": "FAEEDA", "Отменён": "FCEBEB"}
    total_us = 0
    total_ag = 0
    for row_idx, t in enumerate(tickets, 2):
        price_azn = float(t.get("price_azn", t.get("price", 0)))
        cu = float(t.get("cu", 0))
        ca = float(t.get("ca", 0))
        owes_us = price_azn + cu
        owes_ag = price_azn + ca
        total_us += owes_us
        total_ag += owes_ag
        row_data = [
            row_idx - 1, t.get("num", ""), fmt_date(t.get("date", "")), t.get("status", ""),
            t.get("name", ""), t.get("route", ""), t.get("company", ""),
            float(t.get("price_orig", t.get("price", 0))), t.get("currency", "AZN"),
            float(t.get("rate", 1.0)), price_azn, cu, ca, owes_us, owes_ag
        ]
        status = t.get("status", "")
        row_color = status_colors.get(status, "FFFFFF")
        for col, val in enumerate(row_data, 1):
            cell = ws.cell(row=row_idx, column=col, value=val)
            if col in [8, 10, 11, 12, 13, 14, 15]:
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal="right")
            if col == 4:
                cell.fill = PatternFill(start_color=row_color, end_color=row_color, fill_type="solid")
    total_row = len(tickets) + 2
    ws.cell(row=total_row, column=7, value="ИТОГО").font = Font(bold=True)
    ws.cell(row=total_row, column=14, value=total_us).number_format = '#,##0.00'
    ws.cell(row=total_row, column=14).font = Font(bold=True, color="0F6E56")
    ws.cell(row=total_row, column=15, value=total_ag).number_format = '#,##0.00'
    ws.cell(row=total_row, column=15).font = Font(bold=True, color="854F0B")
    filename = f"tickets_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    wb.save(filename)
    return filename


def is_allowed(update: Update):
    if ALLOWED_USER_ID == 0:
        return True
    return update.effective_user.id == ALLOWED_USER_ID


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        return
    keyboard = [[KeyboardButton("📊 Отчёт Excel")], [KeyboardButton("📋 Список билетов"), KeyboardButton("🗑 Очистить всё")]]
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    await update.message.reply_text(
        "Привет! Я помогаю собирать авиабилеты в отчёт.\n\n"
        "Просто отправьте мне:\n"
        "• Фото или скриншот билета\n"
        "• Или текст с данными билета\n\n"
        "Цену автоматически переведу в AZN по курсу ЦБ Азербайджана.\n\n"
        "Можно также указать комиссии:\n"
        "_наша комиссия 15, агенту 5, компания Evrascon_",
        reply_markup=reply_markup,
        parse_mode="Markdown"
    )


async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        return
    await update.message.reply_text("Обрабатываю билет, подождите...")
    photo = update.message.photo[-1]
    file = await photo.get_file()
    file_bytes = await file.download_as_bytearray()
    image_data = base64.standard_b64encode(file_bytes).decode("utf-8")
    caption = update.message.caption or ""
    try:
        result = await parse_ticket_with_claude(text=caption if caption else None, image_data=image_data, image_mime="image/jpeg")
        await process_result(update, context, result)
    except Exception as e:
        logger.error(f"Error: {e}")
        await update.message.reply_text("Не смог распознать билет. Попробуйте отправить текстом.")


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        return
    text = update.message.text
    if text == "📊 Отчёт Excel":
        await send_report(update, context)
        return
    elif text == "📋 Список билетов":
        await list_tickets(update, context)
        return
    elif text == "🗑 Очистить всё":
        save_tickets([])
        await update.message.reply_text("Список очищен.")
        return
    await update.message.reply_text("Обрабатываю...")
    try:
        result = await parse_ticket_with_claude(text=text)
        await process_result(update, context, result)
    except Exception as e:
        logger.error(f"Error: {e}")
        await update.message.reply_text("Не смог разобрать. Попробуйте другой формат.")


async def process_result(update: Update, context: ContextTypes.DEFAULT_TYPE, result: dict):
    if result.get("action") == "add" and result.get("ticket"):
        t = result["ticket"]
        price_orig = float(t.get("price", 0))
        currency = t.get("currency", "AZN").upper()
        cu = float(t.get("cu", 0))
        ca = float(t.get("ca", 0))
        rates = get_cbar_rates()
        price_azn, rate = convert_to_azn(price_orig, currency, rates)
        owes_us = price_azn + cu
        owes_ag = price_azn + ca
        ticket = {
            "id": int(datetime.now().timestamp() * 1000),
            "num": t.get("num", "—"),
            "date": t.get("date", datetime.now().strftime("%Y-%m-%d")),
            "status": t.get("status", "Выписан"),
            "name": t.get("name", "—"),
            "route": t.get("route", "—"),
            "company": t.get("company", "—"),
            "price_orig": price_orig,
            "currency": currency,
            "rate": rate,
            "price_azn": price_azn,
            "price": price_azn,
            "cu": cu,
            "ca": ca,
            "owesUs": owes_us,
            "owesAgent": owes_ag
        }
        tickets = load_tickets()
        tickets.append(ticket)
        save_tickets(tickets)
        missing = result.get("missing", [])
        missing_note = ""
        if missing:
            labels = {"cu": "наша комиссия", "ca": "комиссия агента", "company": "компания", "price": "цена"}
            missing_names = [labels.get(m, m) for m in missing]
            missing_note = f"\n\n⚠️ Не нашёл: {', '.join(missing_names)} — поставил 0"
        currency_note = ""
        if currency != "AZN":
            currency_note = f"\n💱 {money(price_orig)} {currency} × {rate:.4f} = {money(price_azn)} AZN"
        msg = (
            f"✅ Добавил билет *{ticket['num']}*\n"
            f"👤 {ticket['name']}\n"
            f"✈️ {ticket['route']} · {fmt_date(ticket['date'])}\n"
            f"🏢 Компания: {ticket['company']}"
            f"{currency_note}\n"
            f"💰 Цена: {money(price_azn)} AZN · Ком. наша: {money(cu)} · Ком. агента: {money(ca)}\n\n"
            f"🟢 Компания должна нам: *{money(owes_us)} AZN*\n"
            f"🟡 Мы должны агенту: *{money(owes_ag)} AZN*"
            f"{missing_note}"
        )
        await update.message.reply_text(msg, parse_mode="Markdown")
    elif result.get("action") == "chat":
        await update.message.reply_text(result.get("text", "Не понял запрос."))
    else:
        await update.message.reply_text("Не смог распознать билет. Попробуйте другой формат.")


async def send_report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    tickets = load_tickets()
    if not tickets:
        await update.message.reply_text("Нет билетов для отчёта.")
        return
    await update.message.reply_text("Формирую Excel...")
    filename = generate_excel(tickets)
    total_us = sum(t.get("owesUs", 0) for t in tickets)
    total_ag = sum(t.get("owesAgent", 0) for t in tickets)
    with open(filename, "rb") as f:
        await update.message.reply_document(
            document=f,
            filename=filename,
            caption=f"📊 Отчёт · {len(tickets)} билетов\n🟢 Должны нам: {money(total_us)} AZN\n🟡 Мы агентам: {money(total_ag)} AZN"
        )
    os.remove(filename)


async def list_tickets(update: Update, context: ContextTypes.DEFAULT_TYPE):
    tickets = load_tickets()
    if not tickets:
        await update.message.reply_text("Список пуст.")
        return
    lines = [f"📋 *Билеты ({len(tickets)} шт.)*\n"]
    for i, t in enumerate(tickets[-10:], 1):
        price_azn = t.get("price_azn", t.get("price", 0))
        curr = t.get("currency", "AZN")
        curr_note = f" ({money(t.get('price_orig', price_azn))} {curr})" if curr != "AZN" else ""
        lines.append(f"{i}. *{t['num']}* — {t['name']}\n   {t['route']} · {fmt_date(t['date'])} · {t.get('company','')}\n   {money(price_azn)} AZN{curr_note}")
    if len(tickets) > 10:
        lines.append(f"\n_...и ещё {len(tickets)-10} билетов_")
    total_us = sum(t.get("owesUs", 0) for t in tickets)
    total_ag = sum(t.get("owesAgent", 0) for t in tickets)
    lines.append(f"\n🟢 Итого нам: *{money(total_us)} AZN*\n🟡 Итого агентам: *{money(total_ag)} AZN*")
    await update.message.reply_text("\n".join(lines), parse_mode="Markdown")


def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    logger.info("Bot started...")
    app.run_polling()


if __name__ == "__main__":
    main()

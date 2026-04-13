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

SYSTEM_PROMPT = """Ты — ассистент по обработке авиабилетов. Пользователь присылает текст или изображение билета, либо команду для управления списком.

Верни ТОЛЬКО валидный JSON без markdown и блоков кода.

=== ДОБАВЛЕНИЕ БИЛЕТА ===
Если пользователь присылает данные билета:
{"action":"add","ticket":{"num":"номер","date":"YYYY-MM-DD","status":"Выписан","name":"ИМЯ ПАССАЖИРА","route":"XXX-XXX","company":"","price":0,"currency":"AZN","cu":0,"ca":0},"missing":[]}

Правила добавления:
- status: "Выписан", "Изменён" или "Отменён". По умолчанию "Выписан"
- date: формат YYYY-MM-DD. Если не указана — сегодняшняя дата
- name: латиница в формате Firstname Lastname (первая буква заглавная, остальные строчные). Например: Ramazanov Elchin
- route: формат Bak-Tbs или Bak-Tbs-Bak (первая буква каждого сегмента заглавная). Например: Bak-Tbs-Bak
- price: итоговая сумма TOTAL из билета (в оригинальной валюте)
- currency: валюта — AZN, EUR, USD, RUB или KZT. По умолчанию AZN.
- cu: наша комиссия в AZN (если не указана — 0)
- ca: комиссия агента в AZN (если не указана — 0)
- company: компания-заказчик (если не указана — пустая строка)
- missing: список отсутствующих полей

=== УДАЛЕНИЕ БИЛЕТА ===
Если пользователь хочет удалить конкретный билет (указывает имя/фамилию и/или маршрут):
{"action":"delete","name":"Имя Пассажира (формат Ramazanov Elchin)","route":"Маршрут (формат Bak-Tbs)"}
Поля name и route — поисковые критерии. Если указано только имя — ищи только по имени. Если только маршрут — только по маршруту.

=== УДАЛЕНИЕ ВСЕГО ===
Если пользователь пишет "удали всё", "очисти всё", "удалить все билеты" и т.п.:
{"action":"delete_all"}

=== ИЗМЕНЕНИЕ БИЛЕТА ===
Если пользователь хочет изменить данные билета:
{"action":"update","name":"Имя Пассажира (формат Ramazanov Elchin)","route":"Маршрут (формат Bak-Tbs)","fields":{"поле":"новое значение"}}
Возможные поля для изменения: status, company, price, cu, ca, route, date
Если пользователь пишет "исправь цену X на Y" или "цена должна быть Y" — это action:update с fields:{"price":Y}
Пример: изменить статус Ivanov Ivan Bak-Tbs на Отменён → fields: {"status":"Отменён"}

=== ВОЗВРАТ БИЛЕТА ===
Если пользователь пишет "возврат", "это возврат", "верни X", "возврат X AZN" — рядом с данными билета или отдельно:
{"action":"refund","name":"Имя Пассажира","route":"Bak-Tbs","amount":50}
- amount: сумма возврата в AZN (число)
- name и route — для поиска билета

=== ДРУГОЕ ===
Если это вопрос или непонятный текст: {"action":"chat","text":"твой ответ на русском"}"""


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


def find_tickets(tickets, name=None, route=None):
    results = []
    for i, t in enumerate(tickets):
        match = True
        if name:
            n = name.upper().strip()
            tn = t.get("name", "").upper().strip()
            if n not in tn and tn not in n:
                match = False
        if route:
            r = route.upper().strip()
            tr = t.get("route", "").upper().strip()
            if r not in tr and tr not in r:
                match = False
        if match:
            results.append(i)
    return results


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
               "Ком. наша", "Ком. агента", "Компания должна нам", "Мы должны агенту", "Возврат (AZN)"]
    header_fill = PatternFill(start_color="2C2C2A", end_color="2C2C2A", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
    col_widths = [4, 18, 12, 11, 24, 14, 20, 12, 8, 8, 12, 12, 13, 20, 20, 14]
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
        refund = float(t.get("refund", 0))
        row_data = [
            row_idx - 1, t.get("num", ""), fmt_date(t.get("date", "")), t.get("status", ""),
            t.get("name", ""), t.get("route", ""), t.get("company", ""),
            float(t.get("price_orig", t.get("price", 0))), t.get("currency", "AZN"),
            float(t.get("rate", 1.0)), price_azn, cu, ca, owes_us, owes_ag, refund if refund else ""
        ]
        status = t.get("status", "")
        row_color = status_colors.get(status, "FFFFFF")
        for col, val in enumerate(row_data, 1):
            cell = ws.cell(row=row_idx, column=col, value=val)
            if col in [8, 10, 11, 12, 13, 14, 15] or (col == 16 and val != ""):
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
    total_refund = sum(float(t.get("refund", 0)) for t in tickets)
    if total_refund:
        ws.cell(row=total_row, column=16, value=total_refund).number_format = '#,##0.00'
        ws.cell(row=total_row, column=16).font = Font(bold=True, color="A32D2D")
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
        "Управление билетами:\n"
        "• _удали Ramazanov Elchin Bak-Tbs_ — удалить билет\n"
        "• _измени Ramazanov Elchin Bak-Tbs статус Отменён_ — изменить поле\n"
        "• _удали всё_ — очистить список\n\n"
        "Комиссии указывайте так:\n"
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
    action = result.get("action")

    if action == "refund":
        tickets = load_tickets()
        name = result.get("name", "")
        route = result.get("route", "")
        amount = float(result.get("amount", 0))
        indices = find_tickets(tickets, name=name, route=route)
        if not indices:
            await update.message.reply_text("❌ Билет не найден. Проверьте имя или маршрут.")
            return
        for i in indices:
            tickets[i]["status"] = "Отменён"
            tickets[i]["refund"] = -abs(amount)
        save_tickets(tickets)
        lines = [f"↩️ Возврат оформлен: *-{money(amount)} AZN*"]
        for i in indices:
            t = tickets[i]
            lines.append(f"• {t['name']} · {t['route']} · {fmt_date(t['date'])}\n  Статус → Отменён")
        await update.message.reply_text("\n".join(lines), parse_mode="Markdown")
        return

    if action == "delete_all":
        save_tickets([])
        await update.message.reply_text("🗑 Все билеты удалены.")
        return

    if action == "delete":
        tickets = load_tickets()
        name = result.get("name", "")
        route = result.get("route", "")
        indices = find_tickets(tickets, name=name, route=route)
        if not indices:
            await update.message.reply_text("❌ Билет не найден. Проверьте имя или маршрут.")
            return
        deleted = [tickets[i] for i in indices]
        tickets = [t for i, t in enumerate(tickets) if i not in indices]
        save_tickets(tickets)
        lines = [f"🗑 Удалено {len(deleted)} билет(ов):"]
        for t in deleted:
            lines.append(f"• {t['name']} · {t['route']} · {fmt_date(t['date'])}")
        await update.message.reply_text("\n".join(lines))
        return

    if action == "update":
        tickets = load_tickets()
        name = result.get("name", "")
        route = result.get("route", "")
        fields = result.get("fields", {})
        indices = find_tickets(tickets, name=name, route=route)
        if not indices:
            await update.message.reply_text("❌ Билет не найден. Проверьте имя или маршрут.")
            return
        if not fields:
            await update.message.reply_text("❌ Не указано что именно изменить.")
            return
        rates = get_cbar_rates()
        for i in indices:
            for field, value in fields.items():
                tickets[i][field] = value
            if "price" in fields:
                currency = tickets[i].get("currency", "AZN")
                new_price_azn, rate = convert_to_azn(float(fields["price"]), currency, rates)
                tickets[i]["price_orig"] = float(fields["price"])
                tickets[i]["price_azn"] = new_price_azn
                tickets[i]["price"] = new_price_azn
                tickets[i]["rate"] = rate
            price_azn = float(tickets[i].get("price_azn", tickets[i].get("price", 0)))
            cu = float(tickets[i].get("cu", 0))
            ca = float(tickets[i].get("ca", 0))
            tickets[i]["owesUs"] = price_azn + cu
            tickets[i]["owesAgent"] = price_azn + ca
        save_tickets(tickets)
        field_names = {"status": "статус", "company": "компания", "price": "цена", "cu": "ком. наша", "ca": "ком. агента", "route": "маршрут", "date": "дата"}
        changed = ", ".join([f"{field_names.get(k, k)}: {v}" for k, v in fields.items()])
        lines = [f"✏️ Изменено {len(indices)} билет(ов): {changed}"]
        for i in indices:
            t = tickets[i]
            lines.append(f"• {t['name']} · {t['route']} · {fmt_date(t['date'])}")
        await update.message.reply_text("\n".join(lines))
        return

    if action == "add" and result.get("ticket"):
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
        return

    if action == "chat":
        await update.message.reply_text(result.get("text", "Не понял запрос."))
        return

    await update.message.reply_text("Не смог распознать. Попробуйте другой формат.")


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

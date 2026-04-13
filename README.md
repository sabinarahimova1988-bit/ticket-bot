# Ticket Bot — Инструкция по запуску

## Шаг 1 — Создать бота в Telegram

1. Откройте Telegram, найдите @BotFather
2. Напишите /newbot
3. Придумайте имя: например "Мои Билеты"
4. Придумайте username: например "mytickets_bot"
5. BotFather выдаст вам TOKEN — сохраните его

## Шаг 2 — Узнать свой Telegram ID

1. Найдите бота @userinfobot в Telegram
2. Напишите ему /start
3. Он покажет ваш ID (число) — сохраните его

## Шаг 3 — Получить Anthropic API ключ

1. Зайдите на https://console.anthropic.com
2. Зарегистрируйтесь
3. Перейдите в API Keys → Create Key
4. Сохраните ключ (начинается с sk-ant-...)

## Шаг 4 — Запустить на Railway (бесплатно)

1. Зайдите на https://railway.app
2. Зарегистрируйтесь через GitHub
3. Нажмите "New Project" → "Deploy from GitHub repo"
4. Загрузите файлы bot.py и requirements.txt в новый GitHub репозиторий
5. В Railway перейдите в Variables и добавьте:
   - TELEGRAM_TOKEN = ваш токен от BotFather
   - ANTHROPIC_API_KEY = ваш ключ от Anthropic
   - ALLOWED_USER_ID = ваш Telegram ID
6. Railway автоматически запустит бота

## Как пользоваться ботом

- Отправьте фото билета — бот сам всё распознает
- Или напишите текст: "ALIYEV FARID BAK-IST 164.70 комиссия 15 агенту 5 Socar"
- /start — главное меню
- Кнопка "Отчёт Excel" — скачать файл
- Кнопка "Список билетов" — показать последние 10
- Кнопка "Очистить всё" — удалить все билеты

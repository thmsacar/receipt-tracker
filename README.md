# Receipt Tracker

Scan store receipts with Gemini Flash and track expenses in Excel.

## Setup

```bash
pip install -r requirements.txt
cp .env.example .env
```

Add your `GEMINI_API_KEY` in `.env`.

## CLI Usage

Scan a receipt:
```bash
python cli.py add receipt.jpg
```

Other commands:
- `python cli.py list` - list receipts
- `python cli.py show <id>` - show receipt items
- `python cli.py stats` - show spending stats
- `python cli.py export` - update depenses.xlsx

## Telegram Bot

You can also run the Telegram bot:
```bash
python bot.py
```

Set `TELEGRAM_BOT_TOKEN` in `.env`.
To restrict access to your household, set `TELEGRAM_ALLOWED_USER_IDS=123,456`.

In Telegram:
- Send a photo to scan and save it
- `/stats` - spending stats
- `/list` - list receipts (with paging buttons)
- `/show <id>` - show receipt items
- `/export` - download depenses.xlsx

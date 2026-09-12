#!/usr/bin/env python3
"""Telegram bot for receipt tracker with clean visual formatting."""

import hashlib
import html
import json
import logging
import os
import re
import sys
import tempfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from dotenv import load_dotenv
from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.ext import (
    ApplicationBuilder,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

# Disable bytecode generation (__pycache__)
sys.dont_write_bytecode = True

load_dotenv()

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger("receipt_bot")

from cli import (
    CATEGORIES_ALIMENTAIRES,
    CATEGORIES_NON_ALIMENTAIRES,
    analyze_receipt_with_gemini,
    compute_financial_statistics,
    find_duplicate_receipt,
    load_receipts,
    normalize_market_name,
    save_receipts,
    sort_receipts_chronologically,
    sync_excel_file,
)

RECEIPTS_PATH = Path("receipts.json")
EXCEL_PATH = Path("depenses.xlsx")
PAGE_SIZE = 5


def is_authorized(user_id: int) -> bool:
    """Check if the user ID is in the allowed list."""
    raw_allowed = os.environ.get("TELEGRAM_ALLOWED_USER_IDS")
    if raw_allowed is None:
        raw_allowed = os.environ.get("TELEGRAM_ALLOWED_USER_ID")

    if not raw_allowed or not raw_allowed.strip():
        return True

    allowed_ids = {
        uid.strip()
        for uid in re.split(r"[,\s;]+", raw_allowed.strip())
        if uid.strip()
    }
    return str(user_id) in allowed_ids


def format_receipt_detail(receipt: Dict[str, Any], index: int) -> str:
    """Format full receipt details with clean visual structure."""
    r_id = html.escape(receipt.get("id", "-"))
    market = html.escape(receipt.get("market", "Commerce local"))
    branch = f" <i>({html.escape(receipt.get('branch'))})</i>" if receipt.get("branch") else ""
    date = html.escape(receipt.get("date", "-"))
    receipt_no = html.escape(receipt.get("receipt_no") or "-")
    total = float(receipt.get("total_amount") or 0.0)
    items = receipt.get("items") or []

    lines = [
        f"<b>RECEIPT #{index}</b> • <code>{r_id}</code>",
        "━━━━━━━━━━━━━━━━━━━━━━━━━",
        f"<b>Store:</b>      {market}{branch}",
        f"<b>Date:</b>       <code>{date}</code>",
        f"<b>Receipt No:</b> {receipt_no}",
        f"<b>Total:</b>      <b>{total:.2f} €</b>  ({len(items)} items)",
        "",
        "<b>Purchased Items</b>",
        "─────────────────────────",
    ]

    for i, it in enumerate(items, start=1):
        desc = html.escape(it.get("description") or it.get("raw_text") or "Article")
        qty = html.escape(f"{it.get('quantity', 1):.3g} {it.get('unit', '')}".strip())
        price = float(it.get("total_price") or 0.0)
        cat = html.escape(it.get("category") or "Autre")

        item_block = f"<b>{i}. {desc}</b>\n   {qty} ── <b>{price:.2f} €</b>  <i>({cat})</i>"

        # Guard against Telegram 4096 character limit
        if len("\n".join(lines)) + len(item_block) > 3800:
            remaining = len(items) - i + 1
            lines.append(f"\n<i>... and {remaining} more items. Use /export for full Excel log.</i>")
            break

        lines.append(item_block)

    return "\n".join(lines)


def build_list_page(receipts: List[Dict[str, Any]], page: int) -> Tuple[str, InlineKeyboardMarkup]:
    """Build visually separated card list of receipts with inline buttons."""
    total_receipts = len(receipts)
    total_pages = max(1, (total_receipts + PAGE_SIZE - 1) // PAGE_SIZE)
    page = max(1, min(page, total_pages))

    start = (page - 1) * PAGE_SIZE
    end = min(start + PAGE_SIZE, total_receipts)

    lines = [
        "<b>RECEIPTS REGISTER</b>",
        f"<code>Page {page} of {total_pages} • {total_receipts} receipts recorded</code>",
        "━━━━━━━━━━━━━━━━━━━━━━━━━",
        "",
    ]

    item_buttons = []
    for idx in range(start, end):
        r = receipts[idx]
        num = idx + 1
        m = html.escape(r.get("market", "Commerce local"))
        branch = f" <i>({html.escape(r.get('branch'))})</i>" if r.get("branch") else ""
        tot = float(r.get("total_amount") or 0.0)
        cnt = len(r.get("items") or [])
        date = html.escape(r.get("date", "-"))
        r_id = html.escape(r.get("id", "-"))

        lines.append(f"<b>#{num}</b>  <b>{m}</b>{branch}")
        lines.append(f"<code>{date}</code> • <b>{tot:.2f} €</b> • {cnt} items")
        lines.append(f"<code>ID: {r_id}</code>")

        if idx < end - 1:
            lines.append("─────────────────────────")

        item_buttons.append(
            InlineKeyboardButton(f"#{num} Details", callback_data=f"show:{num}:{page}")
        )

    keyboard: List[List[InlineKeyboardButton]] = []

    # Row of quick inspection buttons (2 buttons per row for easy finger tapping)
    for i in range(0, len(item_buttons), 2):
        keyboard.append(item_buttons[i : i + 2])

    # Navigation row
    nav_row = []
    if page > 1:
        nav_row.append(InlineKeyboardButton("◀ Prev", callback_data=f"page:{page - 1}"))
    else:
        nav_row.append(InlineKeyboardButton("·", callback_data="noop"))

    nav_row.append(InlineKeyboardButton(f"{page} / {total_pages}", callback_data="noop"))

    if page < total_pages:
        nav_row.append(InlineKeyboardButton("Next ▶", callback_data=f"page:{page + 1}"))
    else:
        nav_row.append(InlineKeyboardButton("·", callback_data="noop"))

    keyboard.append(nav_row)

    return "\n".join(lines), InlineKeyboardMarkup(keyboard)


async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handle /start and /help commands."""
    if not update.effective_user or not update.message:
        return

    if not is_authorized(update.effective_user.id):
        await update.message.reply_text(
            f"Unauthorized user ID: <code>{update.effective_user.id}</code>\n"
            "Add it to <code>TELEGRAM_ALLOWED_USER_IDS</code> in <code>.env</code>.",
            parse_mode="HTML",
        )
        return

    help_text = (
        "<b>RECEIPT TRACKER</b>\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━\n"
        "<b>Commands:</b>\n"
        "• /stats — spending breakdown by category and store\n"
        "• /list — browse receipts with interactive paging\n"
        "• /show <code>&lt;id&gt;</code> — view itemized receipt\n"
        "• /export — download <code>depenses.xlsx</code>\n"
        "• /help — show this message\n\n"
        "Send a photo of any receipt to scan and log it."
    )
    await update.message.reply_text(help_text, parse_mode="HTML")


async def stats_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handle /stats command with clean typography and aligned tables."""
    if not update.effective_user or not update.message:
        return

    if not is_authorized(update.effective_user.id):
        await update.message.reply_text("Unauthorized.")
        return

    receipts = load_receipts(RECEIPTS_PATH)
    if not receipts:
        await update.message.reply_text("No receipts found.")
        return

    sorted_receipts = sort_receipts_chronologically(receipts)
    st = compute_financial_statistics(sorted_receipts)

    lines = [
        "<b>EXPENSE STATISTICS</b>",
        "━━━━━━━━━━━━━━━━━━━━━━━━━",
        f"<b>Total Spent:</b>  <code>{st['total_general']:>7.2f} €</code>",
        f"<b>Receipts:</b>     {st['total_receipts']}",
        f"<b>Items:</b>        {st['total_items']}",
        f"<b>Average:</b>      <code>{st['avg_receipt']:>7.2f} €</code> / receipt",
        "",
        "<b>Macro Split</b>",
        "─────────────────────────",
        f"• <b>Food:</b>     <code>{st['food_total']:>7.2f} €</code>  ({st['food_items']} items • {st['food_pct']:.1f}%)",
        f"• <b>Non-Food:</b> <code>{st['nonfood_total']:>7.2f} €</code>  ({st['nonfood_items']} items • {st['nonfood_pct']:.1f}%)",
        "",
        "<b>Food by Category</b>",
        "─────────────────────────",
    ]

    for cat in CATEGORIES_ALIMENTAIRES:
        amt = st["cat_spend"].get(cat, 0.0)
        cnt = st["cat_items"].get(cat, 0)
        p_food = (amt / st["food_total"] * 100) if st["food_total"] else 0.0
        lines.append(f"• <b>{html.escape(cat)}</b>\n  <code>{amt:>7.2f} €</code>  ({cnt} items • {p_food:.1f}% food)")

    lines.append("")
    lines.append("<b>Non-Food by Category</b>")
    lines.append("─────────────────────────")
    for cat in CATEGORIES_NON_ALIMENTAIRES:
        amt = st["cat_spend"].get(cat, 0.0)
        cnt = st["cat_items"].get(cat, 0)
        p_nf = (amt / st["nonfood_total"] * 100) if st["nonfood_total"] else 0.0
        lines.append(f"• <b>{html.escape(cat)}</b>\n  <code>{amt:>7.2f} €</code>  ({cnt} items • {p_nf:.1f}% non-food)")

    lines.append("")
    lines.append("<b>Stores Breakdown</b>")
    lines.append("─────────────────────────")
    for m_name, amt in sorted(st["market_spend"].items(), key=lambda x: x[1], reverse=True):
        cnt = st["market_counts"].get(m_name, 0)
        pct = (amt / st["total_general"] * 100) if st["total_general"] else 0.0
        rcpt_word = "receipt" if cnt == 1 else "receipts"
        lines.append(f"• <b>{html.escape(m_name)}</b>\n  <code>{amt:>7.2f} €</code>  ({cnt} {rcpt_word} • {pct:.1f}%)")

    await update.message.reply_text("\n".join(lines), parse_mode="HTML")


async def list_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handle /list command with paginated cards."""
    if not update.effective_user or not update.message:
        return

    if not is_authorized(update.effective_user.id):
        await update.message.reply_text("Unauthorized.")
        return

    receipts = load_receipts(RECEIPTS_PATH)
    if not receipts:
        await update.message.reply_text("No receipts found.")
        return

    sorted_receipts = sort_receipts_chronologically(receipts)

    page = 1
    if context.args and len(context.args) > 0:
        try:
            page = int(context.args[0])
        except ValueError:
            page = 1

    text, markup = build_list_page(sorted_receipts, page)
    await update.message.reply_text(text, reply_markup=markup, parse_mode="HTML")


async def show_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handle /show command with structured typography."""
    if not update.effective_user or not update.message:
        return

    if not is_authorized(update.effective_user.id):
        await update.message.reply_text("Unauthorized.")
        return

    if not context.args or len(context.args) == 0:
        await update.message.reply_text("Usage: <code>/show &lt;number or id&gt;</code>", parse_mode="HTML")
        return

    query = context.args[0].strip()
    receipts = load_receipts(RECEIPTS_PATH)
    if not receipts:
        await update.message.reply_text("No receipts found.")
        return

    sorted_receipts = sort_receipts_chronologically(receipts)

    target = None
    target_idx = 0

    if query.isdigit():
        idx = int(query)
        if 1 <= idx <= len(sorted_receipts):
            target = sorted_receipts[idx - 1]
            target_idx = idx

    if not target:
        for i, r in enumerate(sorted_receipts, start=1):
            if r.get("id") == query:
                target = r
                target_idx = i
                break

    if not target:
        await update.message.reply_text(f"No receipt found for <code>{html.escape(query)}</code>.", parse_mode="HTML")
        return

    text = format_receipt_detail(target, target_idx)
    await update.message.reply_text(text, parse_mode="HTML")


async def button_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handle pagination and view-detail button clicks."""
    query = update.callback_query
    if not query or not query.data:
        return

    await query.answer()

    if not update.effective_user or not is_authorized(update.effective_user.id):
        await query.answer("Unauthorized.", show_alert=True)
        return

    data = query.data
    if data == "noop":
        return

    receipts = load_receipts(RECEIPTS_PATH)
    if not receipts:
        await query.edit_message_text("No receipts found.")
        return

    sorted_receipts = sort_receipts_chronologically(receipts)

    if data.startswith("page:"):
        try:
            page = int(data.split(":")[1])
        except ValueError:
            page = 1
        text, markup = build_list_page(sorted_receipts, page)
        await query.edit_message_text(text, reply_markup=markup, parse_mode="HTML")

    elif data.startswith("show:"):
        parts = data.split(":")
        try:
            target_idx = int(parts[1])
            return_page = int(parts[2]) if len(parts) > 2 else 1
        except ValueError:
            return

        if 1 <= target_idx <= len(sorted_receipts):
            target = sorted_receipts[target_idx - 1]
            text = format_receipt_detail(target, target_idx)

            back_markup = InlineKeyboardMarkup([
                [InlineKeyboardButton(f"◀ Back to list (Page {return_page})", callback_data=f"page:{return_page}")]
            ])
            await query.edit_message_text(text, reply_markup=back_markup, parse_mode="HTML")


async def export_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handle /export command: regenerate and upload Excel file."""
    if not update.effective_user or not update.message:
        return

    if not is_authorized(update.effective_user.id):
        await update.message.reply_text("Unauthorized.")
        return

    if not RECEIPTS_PATH.exists():
        await update.message.reply_text("No receipts found.")
        return

    sync_excel_file(RECEIPTS_PATH, EXCEL_PATH)

    if not EXCEL_PATH.exists():
        await update.message.reply_text("Error generating Excel file.")
        return

    with open(EXCEL_PATH, "rb") as doc:
        await update.message.reply_document(
            document=doc,
            filename=EXCEL_PATH.name,
            caption="Excel workbook: <code>depenses.xlsx</code>",
            parse_mode="HTML",
        )


async def photo_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handle incoming receipt photos."""
    if not update.effective_user or not update.message:
        return

    if not is_authorized(update.effective_user.id):
        await update.message.reply_text("Unauthorized.")
        return

    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        await update.message.reply_text("<code>GEMINI_API_KEY</code> is not set in <code>.env</code>.", parse_mode="HTML")
        return

    model = os.environ.get("GEMINI_MODEL", "gemini-2.5-flash")
    status_msg = await update.message.reply_text("Analyzing receipt...")

    photo = update.message.photo[-1]
    tg_file = await photo.get_file()

    with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tmp:
        tmp_path = Path(tmp.name)

    try:
        await tg_file.download_to_drive(custom_path=tmp_path)

        with open(tmp_path, "rb") as f:
            img_bytes = f.read()
        image_hash = hashlib.sha256(img_bytes).hexdigest()

        current_receipts = load_receipts(RECEIPTS_PATH)

        dup_by_hash = next((r for r in current_receipts if r.get("image_hash") == image_hash), None)
        if dup_by_hash:
            d_date = html.escape(dup_by_hash.get("date", "-"))
            d_market = html.escape(dup_by_hash.get("market", "-"))
            d_tot = float(dup_by_hash.get("total_amount") or 0.0)
            await status_msg.edit_text(
                f"<b>Duplicate Photo:</b>\n"
                f"Already added on <code>{d_date}</code> ({d_market}, <b>{d_tot:.2f} €</b>).",
                parse_mode="HTML",
            )
            return

        parsed_data = analyze_receipt_with_gemini(tmp_path, api_key=api_key, model_name=model)
        receipt_dict = parsed_data.model_dump()
        receipt_dict["image_hash"] = image_hash

        existing_markets = [r.get("market") for r in current_receipts if r.get("market")]
        receipt_dict["market"] = normalize_market_name(receipt_dict.get("market"), existing_markets=existing_markets)

        dup = find_duplicate_receipt(receipt_dict, current_receipts, image_hash=image_hash)
        if dup:
            d_id = html.escape(dup.get("id", "-"))
            d_market = html.escape(dup.get("market", "-"))
            d_date = html.escape(dup.get("date", "-"))
            d_tot = float(dup.get("total_amount") or 0.0)
            await status_msg.edit_text(
                f"<b>Duplicate Receipt Detected:</b>\n"
                f"<code>{d_id}</code> • {d_market} • <code>{d_date}</code> • <b>{d_tot:.2f} €</b>\n"
                "Addition skipped.",
                parse_mode="HTML",
            )
            return

        current_receipts.append(receipt_dict)
        current_receipts = sort_receipts_chronologically(current_receipts)
        save_receipts(RECEIPTS_PATH, current_receipts)
        sync_excel_file(RECEIPTS_PATH, EXCEL_PATH)

        assigned_index = len(current_receipts)
        for idx, r in enumerate(current_receipts, start=1):
            if r.get("receipt_no") == receipt_dict.get("receipt_no") and r.get("total_amount") == receipt_dict.get("total_amount"):
                assigned_index = idx
                break

        detail_text = format_receipt_detail(receipt_dict, assigned_index)
        await status_msg.edit_text(detail_text, parse_mode="HTML")

    except Exception as e:
        logger.exception("Error processing receipt photo")
        await status_msg.edit_text(f"Error: {html.escape(str(e))}", parse_mode="HTML")
    finally:
        if tmp_path.exists():
            tmp_path.unlink()


def main() -> None:
    """Run the bot."""
    token = os.environ.get("TELEGRAM_BOT_TOKEN")
    if not token or not token.strip():
        print("Error: TELEGRAM_BOT_TOKEN is not set.")
        print("Set TELEGRAM_BOT_TOKEN in .env.")
        sys.exit(1)

    app = ApplicationBuilder().token(token.strip()).build()

    app.add_handler(CommandHandler("start", start_command))
    app.add_handler(CommandHandler("help", start_command))
    app.add_handler(CommandHandler("stats", stats_command))
    app.add_handler(CommandHandler("list", list_command))
    app.add_handler(CommandHandler("show", show_command))
    app.add_handler(CommandHandler("export", export_command))
    app.add_handler(CallbackQueryHandler(button_callback))
    app.add_handler(MessageHandler(filters.PHOTO, photo_handler))

    print("Bot is running...")
    app.run_polling()


if __name__ == "__main__":
    main()

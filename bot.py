import logging
import asyncio
import io
import json
import base64
from pathlib import Path
from telegram import Update
from telegram.constants import ChatAction, ChatType
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    ContextTypes,
    filters,
)
from openai import OpenAI

from config import TELEGRAM_BOT_TOKEN, ADMIN_USER_ID, OPENAI_API_KEY, OPENAI_MODEL
from tools import (
    BOT_TOOLS,
    create_and_send_word_document,
    create_and_send_excel_sheet,
    convert_pdf_to_word,
    download_instagram_or_video,
    search_and_download_music,
    send_file_to_telegram,
    web_search,
    run_command
)

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.INFO
)
logger = logging.getLogger(__name__)

client = OpenAI(api_key=OPENAI_API_KEY)

group_monitoring_rules = {
    "keywords": set(),
    "enabled": False,
    "last_forwarded": set()
}

SYSTEM_PROMPT = """Siz egangiz uchun shaxsiy, eng qobiliyatli professional super-yordamchisiz (Personal AI Executive & Super Operator).

SIZNING ASBOBLARINGIZ (TOOLS):
1. `download_instagram_or_video`: Instagram (Reels, Post, Story) yoki video havolasi kelganda uni yuklab berish.
2. `search_and_download_music`: Foydalanuvchi biror qo'shiq yoki musiqa so'rasa, uni qidirib audio qilib tashlab berish.
3. `create_and_send_word_document`: Word (.docx) fayl yaratish.
4. `create_and_send_excel_sheet`: Excel (.xlsx) jadval yaratish.
5. `convert_pdf_to_word`: PDF faylni Wordga aylantirish.
6. `web_search`: Internetdan qidirish.
7. `run_command`: Terminalda buyruqlarni yurgizish.

MUHIM QOIDALAR:
- Foydalanuvchi nima buyursa (ovozda aytsa ham, matnda yozsa ham), faqat quruq matn qaytarib qo'ymasdan, darhol kerakli asbobni chaqiring!
- Masalan: "falonchi musiqani top", "shu reelsni yuklab ber", "buni Word/Excel qil" desa — darhol asbobni chaqiring!
- Har doim samimiy, o'zbek tilida va tayyor natijani bering.
"""

OPENAI_TOOLS = [
    {
        "type": "function",
        "function": {
            "name": "download_instagram_or_video",
            "description": "Instagram (Reels, Post) yoki video havolasidan videoni yuklash",
            "parameters": {
                "type": "object",
                "properties": {
                    "url": {"type": "string", "description": "Video havolasi"}
                },
                "required": ["url"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "search_and_download_music",
            "description": "Musiqa yoki qo'shiq qidirib, audioni yuklab berish",
            "parameters": {
                "type": "object",
                "properties": {
                    "query": {"type": "string", "description": "Musiqa nomi yoki ijrochi"}
                },
                "required": ["query"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "create_and_send_word_document",
            "description": "Word (.docx) faylini yaratish",
            "parameters": {
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "Fayl nomi (.docx)"},
                    "title": {"type": "string", "description": "Sarlavha"},
                    "content": {"type": "string", "description": "To'liq matn"}
                },
                "required": ["filename", "title", "content"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "create_and_send_excel_sheet",
            "description": "Professional Excel (.xlsx) jadval yaratish",
            "parameters": {
                "type": "object",
                "properties": {
                    "filename": {"type": "string", "description": "Fayl nomi (.xlsx)"},
                    "sheet_title": {"type": "string", "description": "Varaq nomi"},
                    "headers": {"type": "array", "items": {"type": "string"}, "description": "Ustunlar"},
                    "rows": {"type": "array", "items": {"type": "array", "items": {"type": "string"}}, "description": "Qatorlar"}
                },
                "required": ["filename", "sheet_title", "headers", "rows"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "convert_pdf_to_word",
            "description": "PDF faylni Word (.docx) ga aylantirish",
            "parameters": {
                "type": "object",
                "properties": {
                    "pdf_path": {"type": "string", "description": "PDF fayl manzili"},
                    "output_docx_name": {"type": "string", "description": "Word fayl nomi"}
                },
                "required": ["pdf_path"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "send_file_to_telegram",
            "description": "Faylni Telegramga yuborish",
            "parameters": {
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "Fayl manzili"}
                },
                "required": ["path"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "web_search",
            "description": "Internetdan qidirish",
            "parameters": {
                "type": "object",
                "properties": {
                    "query": {"type": "string", "description": "Qidiruv so'zi"}
                },
                "required": ["query"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "run_command",
            "description": "Terminal buyrug'ini bajarish",
            "parameters": {
                "type": "object",
                "properties": {
                    "command": {"type": "string", "description": "Buyruq"}
                },
                "required": ["command"]
            }
        }
    }
]

tool_funcs = {
    "download_instagram_or_video": download_instagram_or_video,
    "search_and_download_music": search_and_download_music,
    "create_and_send_word_document": create_and_send_word_document,
    "create_and_send_excel_sheet": create_and_send_excel_sheet,
    "convert_pdf_to_word": convert_pdf_to_word,
    "send_file_to_telegram": send_file_to_telegram,
    "web_search": web_search,
    "run_command": run_command
}

conversations = {}

def get_history(user_id: int):
    if user_id not in conversations:
        conversations[user_id] = [
            {"role": "system", "content": SYSTEM_PROMPT}
        ]
    return conversations[user_id]

async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_USER_ID: return
    await update.message.reply_text(
        "👋 Assalomu alaykum, Xo'jayin!\n\n"
        "Men sizning eng aqlli Super Yordamchingizman!\n\n"
        "IMKONIYATLARIM:\n"
        "🎤 Ovozli buyruqlar — nima desangiz darhol bajaraman!\n"
        "📥 Instagram (Reels / Post) — link tashlang, videoni yuklab beraman!\n"
        "🎵 Musiqa topish — istalgan qo'shiq nomini aytsangiz, topib tashlayman!\n"
        "📄 Word & 📊 Excel — fayllar va jadvallar yarataman!\n"
        "📑 PDF ➡️ Word — PDF ni Word ga aylantiraman!\n"
        "🔍 Guruhlar monitoringi — guruhlardan kerakli ma'lumotlarni qidirib topaman!\n\n"
        "Nima buyurasiz?"
    )

async def run_gpt_loop(update: Update, context: ContextTypes.DEFAULT_TYPE, user_msg):
    chat_id = update.effective_chat.id
    user_id = update.effective_user.id
    await context.bot.send_chat_action(chat_id=chat_id, action=ChatAction.TYPING)

    # To'g'ridan-to'g'ri Instagram link tashlangan bo'lsa
    if isinstance(user_msg, str) and ("instagram.com" in user_msg or "tiktok.com" in user_msg or "youtu.be" in user_msg or "youtube.com" in user_msg):
        for word in user_msg.split():
            if "http://" in word or "https://" in word:
                await update.message.reply_text("⏳ Video yuklanmoqda, iltimos kuting...")
                res = await asyncio.get_event_loop().run_in_executor(None, lambda: download_instagram_or_video(word))
                if res.startswith("VIDEO_READY:"):
                    v_path = res.replace("VIDEO_READY:", "").strip()
                    if Path(v_path).exists():
                        with open(v_path, "rb") as vf:
                            await context.bot.send_video(chat_id=chat_id, video=vf, caption="🎬 Marhamat, videongiz tayyor!")
                        return

    # Guruh monitoringi sozlamalari
    msg_str = str(user_msg).lower()
    if "guruh" in msg_str and ("kuzat" in msg_str or "qidir" in msg_str or "top" in msg_str):
        words = [w.strip() for w in msg_str.replace("'", " ").replace('"', " ").replace(",", " ").split() if len(w) > 2 and w not in ["guruh", "kuzat", "menga", "tashla", "yubor", "kerak", "agar"]]
        group_monitoring_rules["keywords"].update(words)
        group_monitoring_rules["enabled"] = True

    history = get_history(user_id)
    history.append({"role": "user", "content": user_msg})

    if len(history) > 26:
        history = [history[0]] + history[-24:]
        conversations[user_id] = history

    try:
        for _ in range(8):
            response = await asyncio.get_event_loop().run_in_executor(
                None,
                lambda: client.chat.completions.create(
                    model=OPENAI_MODEL,
                    messages=history,
                    tools=OPENAI_TOOLS,
                    tool_choice="auto",
                    temperature=0.2
                )
            )

            msg = response.choices[0].message

            if not msg.tool_calls:
                answer = msg.content or "Topshiriq bajarildi."
                history.append({"role": "assistant", "content": answer})
                for i in range(0, len(answer), 4000):
                    await update.message.reply_text(answer[i:i+4000])
                return

            history.append(msg)
            for tool_call in msg.tool_calls:
                fn_name = tool_call.function.name
                fn_args = json.loads(tool_call.function.arguments)

                await context.bot.send_chat_action(chat_id=chat_id, action=ChatAction.TYPING)

                try:
                    res = await asyncio.get_event_loop().run_in_executor(
                        None, lambda: tool_funcs[fn_name](**fn_args)
                    )
                except Exception as err:
                    res = f"Xatolik: {err}"

                # Natijaga qarab faylni jo'natish
                res_str = str(res)
                if res_str.startswith("VIDEO_READY:"):
                    f_path = res_str.replace("VIDEO_READY:", "").strip()
                    if Path(f_path).exists():
                        with open(f_path, "rb") as vf:
                            await context.bot.send_video(chat_id=chat_id, video=vf, caption="🎬 Marhamat, videongiz tayyor!")
                elif res_str.startswith("AUDIO_READY:"):
                    f_path = res_str.replace("AUDIO_READY:", "").strip()
                    if Path(f_path).exists():
                        with open(f_path, "rb") as af:
                            await context.bot.send_audio(chat_id=chat_id, audio=af, caption="🎵 Marhamat, musiqangiz tayyor!")
                elif res_str.startswith("FILE_READY:"):
                    f_path = res_str.replace("FILE_READY:", "").strip()
                    if Path(f_path).exists():
                        with open(f_path, "rb") as df:
                            await context.bot.send_document(chat_id=chat_id, document=df, caption=f"📁 {Path(f_path).name}")

                history.append({
                    "role": "tool",
                    "tool_call_id": tool_call.id,
                    "content": str(res)
                })

    except Exception as e:
        logger.error(f"Xatolik: {e}", exc_info=True)
        await update.message.reply_text(f"⚠️ Xatolik: {str(e)[:300]}")

async def handle_group_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = update.message
    if not msg or not msg.text: return
    text = msg.text.lower()
    chat_title = msg.chat.title or "Telegram Guruhi"
    user_name = msg.from_user.full_name if msg.from_user else "Foydalanuvchi"

    match = False
    for kw in group_monitoring_rules["keywords"]:
        if kw in text:
            match = True
            break

    if match and msg.message_id not in group_monitoring_rules["last_forwarded"]:
        group_monitoring_rules["last_forwarded"].add(msg.message_id)
        notification = (
            f"🔔 *Guruhdan qidiruv bo'yicha xabar topildi!*\n\n"
            f"👥 *Guruh:* {chat_title}\n"
            f"👤 *Yuboruvchi:* {user_name}\n\n"
            f"💬 *Xabar:*\n{msg.text}\n"
        )
        try:
            await context.bot.send_message(
                chat_id=ADMIN_USER_ID,
                text=notification,
                parse_mode="Markdown"
            )
        except Exception as e:
            logger.error(f"Xatolik: {e}")

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_chat.type in [ChatType.GROUP, ChatType.SUPERGROUP]:
        await handle_group_message(update, context)
        return
    if update.effective_user.id != ADMIN_USER_ID: return
    if update.message.text:
        await run_gpt_loop(update, context, update.message.text)

async def handle_voice(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_chat.type in [ChatType.GROUP, ChatType.SUPERGROUP]: return
    if update.effective_user.id != ADMIN_USER_ID: return

    await update.message.reply_chat_action(ChatAction.RECORD_VOICE)
    voice = update.message.voice or update.message.audio
    voice_file = await voice.get_file()
    byte_arr = await voice_file.download_as_bytearray()

    audio_io = io.BytesIO(bytes(byte_arr))
    audio_io.name = "voice.ogg"

    try:
        transcription = await asyncio.get_event_loop().run_in_executor(
            None,
            lambda: client.audio.transcriptions.create(
                model="whisper-1",
                file=audio_io,
                language="uz",
                prompt="O'zbek tili, o'zbekcha buyruqlar, topshiriqlar, hisobotlar, qo'shiqlar va so'rovlar."
            )
        )
        transcribed_text = transcription.text
    except Exception as e:
        await update.message.reply_text(f"⚠️ Ovozni eshitishda xatolik: {e}")
        return

    # Ovozli buyruqni to'g'ridan-to'g'ri ijroga yuborish
    logger.info(f"Ovozli buyruq ijroga yuborilmoqda: {transcribed_text}")
    await run_gpt_loop(update, context, transcribed_text)

async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_chat.type in [ChatType.GROUP, ChatType.SUPERGROUP]: return
    if update.effective_user.id != ADMIN_USER_ID: return

    await update.message.reply_chat_action(ChatAction.TYPING)
    photo = update.message.photo[-1]
    photo_file = await photo.get_file()
    byte_arr = await photo_file.download_as_bytearray()
    b64 = base64.b64encode(bytes(byte_arr)).decode('utf-8')

    caption = update.message.caption or "Ushbu rasmni tahlil qiling va vazifani bajaring."
    user_content = [
        {"type": "text", "text": caption},
        {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64}"}}
    ]
    await run_gpt_loop(update, context, user_content)

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_chat.type in [ChatType.GROUP, ChatType.SUPERGROUP]: return
    if update.effective_user.id != ADMIN_USER_ID: return

    await update.message.reply_chat_action(ChatAction.TYPING)
    doc = update.message.document
    doc_file = await doc.get_file()
    local_path = Path(doc.file_name).resolve()
    await doc_file.download_to_drive(custom_path=str(local_path))

    if doc.file_name.lower().endswith(".pdf"):
        prompt = (
            f"Foydalanuvchi '{doc.file_name}' nomli PDF faylni yukladi. Manzili: '{str(local_path)}'.\n"
            f"Agar uni Wordga o'girish so'ralgan bo'lsa, `convert_pdf_to_word(pdf_path='{str(local_path)}')` asbobini chaqiring!"
        )
        caption = update.message.caption or "Ushbu PDF ni tahlil qiling yoki Word ga aylantirib bering."
        await run_gpt_loop(update, context, f"{prompt}\nFoydalanuvchi izohi: {caption}")
        return

    try:
        with open(local_path, "r", encoding="utf-8", errors="replace") as f:
            file_text = f.read(10000)
    except Exception:
        file_text = f"[{doc.file_name} fayli saqlandi]"

    caption = update.message.caption or f"'{doc.file_name}' fayli yuborildi."
    combined = f"{caption}\n\nFayl mazmuni:\n```\n{file_text}\n```"
    await run_gpt_loop(update, context, combined)

def run_health_server():
    import threading
    from http.server import HTTPServer, BaseHTTPRequestHandler
    import os

    class SimpleHandler(BaseHTTPRequestHandler):
        def do_GET(self):
            self.send_response(200)
            self.end_headers()
            self.wfile.write(b"Bot 24/7 is running!")
        def log_message(self, format, *args):
            return # Jim ishlash

    port = int(os.environ.get("PORT", 8080))
    server = HTTPServer(("0.0.0.0", port), SimpleHandler)
    server.serve_forever()

def main():
    print("Telegram Professional Super Bot ishga tushmoqda...")
    
    # Render.com bepul serveri uchun port tinglovchi
    import threading
    threading.Thread(target=run_health_server, daemon=True).start()

    app = ApplicationBuilder().token(TELEGRAM_BOT_TOKEN).build()

    app.add_handler(CommandHandler("start", start_command))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    app.add_handler(MessageHandler(filters.VOICE | filters.AUDIO, handle_voice))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))

    print("Bot to'liq tayyor! Ovozli buyruqlar, Instagram yuklash va Musiqa qidirish ulandi!")
    app.run_polling()

if __name__ == "__main__":
    main()

import os
import asyncio
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, CallbackQueryHandler, MessageHandler, filters, ContextTypes
from telegram.constants import ParseMode
from dotenv import load_dotenv
from datetime import datetime

load_dotenv()

BOT_TOKEN = os.getenv("BOT_TOKEN", "8625557628:AAHeUC2WxfMjJk-RRq3IxTtUJoc0H4XSsAM")
ADMIN_ID = int(os.getenv("ADMIN_ID", "7758296066"))

# Global data storage
user_data_store = {}
documents_store = {}

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Start command"""
    user_id = update.effective_user.id
    user_name = update.effective_user.first_name
    
    if user_id not in user_data_store:
        user_data_store[user_id] = {
            "name": user_name,
            "documents": 0,
            "created_at": datetime.now()
        }
    
    keyboard = [
        [
            InlineKeyboardButton("📝 Referat", callback_data="doc_referat"),
            InlineKeyboardButton("📚 Kurs Ishi", callback_data="doc_kurs")
        ],
        [
            InlineKeyboardButton("📄 Maqola", callback_data="doc_maqola"),
            InlineKeyboardButton("🎯 Slide", callback_data="doc_slide")
        ],
        [
            InlineKeyboardButton("📊 Statistika", callback_data="stats"),
            InlineKeyboardButton("❓ Yordam", callback_data="help")
        ]
    ]
    
    await update.message.reply_text(
        f"🎓 Assalomu alaykum, {user_name}!\n\n"
        "**Akademik Yordamchi Botga Xush Kelibsiz!**\n\n"
        "Siz quyidagi hujjatlarni yarata olasiz:\n"
        "📝 Referat\n"
        "📚 Kurs Ishi\n"
        "📄 Maqola\n"
        "🎯 Slide\n\n"
        "💡 Birinchi hujjat TEKIN!\n"
        "Keyingisini yaratish uchun: 2,000 UZS",
        reply_markup=InlineKeyboardMarkup(keyboard),
        parse_mode=ParseMode.MARKDOWN
    )

async def handle_doc_type(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle document type selection"""
    query = update.callback_query
    await query.answer()
    
    user_id = query.from_user.id
    doc_type = query.data.replace("doc_", "")
    
    if user_id not in user_data_store:
        user_data_store[user_id] = {}
    
    user_data_store[user_id]["doc_type"] = doc_type
    
    doc_types = {
        "referat": "📝 Referat",
        "kurs": "📚 Kurs Ishi",
        "maqola": "📄 Maqola",
        "slide": "🎯 Slide"
    }
    
    keyboard = [
        [
            InlineKeyboardButton("🎯 APA", callback_data="style_apa"),
            InlineKeyboardButton("🎯 Harvard", callback_data="style_harvard")
        ],
        [
            InlineKeyboardButton("🎯 Uzbek", callback_data="style_uzbek"),
            InlineKeyboardButton("🎯 Chicago", callback_data="style_chicago")
        ]
    ]
    
    await query.edit_message_text(
        f"✅ Siz {doc_types.get(doc_type)} turini tanladingiz\n\n"
        "📋 Endi uslubni tanlang:",
        reply_markup=InlineKeyboardMarkup(keyboard),
        parse_mode=ParseMode.MARKDOWN
    )

async def handle_style(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle style selection"""
    query = update.callback_query
    await query.answer()
    
    user_id = query.from_user.id
    style = query.data.replace("style_", "")
    
    if user_id not in user_data_store:
        user_data_store[user_id] = {}
    
    user_data_store[user_id]["style"] = style
    
    styles = {
        "apa": "APA Style",
        "harvard": "Harvard Style",
        "uzbek": "O'zbek Standarti",
        "chicago": "Chicago Style"
    }
    
    await query.edit_message_text(
        f"✅ Siz {styles.get(style)} ni tanladingiz\n\n"
        "✏️ Hujjatning **sarlavhasini** yuboring:"
    )
    
    context.user_data[user_id] = {"state": "title"}

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle text messages"""
    user_id = update.effective_user.id
    text = update.message.text
    
    if user_id not in context.user_data:
        context.user_data[user_id] = {}
    
    state = context.user_data[user_id].get("state")
    
    if state == "title":
        if user_id not in user_data_store:
            user_data_store[user_id] = {}
        
        user_data_store[user_id]["title"] = text
        context.user_data[user_id]["state"] = "content"
        
        await update.message.reply_text(
            f"✅ Sarlavha qabul qilindi: **{text}**\n\n"
            "📝 Endi **matnni** yuboring:"
        )
    
    elif state == "content":
        if user_id not in user_data_store:
            user_data_store[user_id] = {}
        
        user_data_store[user_id]["content"] = text
        
        doc_id = f"doc_{user_id}_{datetime.now().timestamp()}"
        documents_store[doc_id] = user_data_store[user_id]
        
        keyboard = [
            [InlineKeyboardButton("✅ Yangi Hujjat", callback_data="new_doc")],
            [InlineKeyboardButton("📊 Statistika", callback_data="stats")],
            [InlineKeyboardButton("🏠 Bosh Menyu", callback_data="start")]
        ]
        
        await update.message.reply_text(
            f"✅ **Hujjat Tayyorlandi!**\n\n"
            f"📝 Tur: {user_data_store[user_id].get('doc_type')}\n"
            f"📋 Uslub: {user_data_store[user_id].get('style')}\n"
            f"📖 Sarlavha: {user_data_store[user_id].get('title')}\n\n"
            f"🎁 Birinchi hujjat tekin edi!",
            reply_markup=InlineKeyboardMarkup(keyboard),
            parse_mode=ParseMode.MARKDOWN
        )
        
        if user_id in user_data_store:
            user_data_store[user_id]["documents"] = user_data_store[user_id].get("documents", 0) + 1
        
        context.user_data[user_id]["state"] = None

async def handle_stats(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Show statistics"""
    query = update.callback_query
    await query.answer()
    
    user_id = query.from_user.id
    user_info = user_data_store.get(user_id, {})
    
    text = (
        f"📊 **SHAXSIY STATISTIKA**\n\n"
        f"👤 Ism: {user_info.get('name', 'Unknown')}\n"
        f"📄 Hujjatlar: {user_info.get('documents', 0)}\n"
        f"💾 ID: {user_id}\n\n"
        f"✅ Bot ishga tushgan!"
    )
    
    keyboard = [
        [InlineKeyboardButton("📝 Yangi Hujjat", callback_data="new_doc")],
        [InlineKeyboardButton("🏠 Bosh Menyu", callback_data="start")]
    ]
    
    await query.edit_message_text(
        text,
        reply_markup=InlineKeyboardMarkup(keyboard),
        parse_mode=ParseMode.MARKDOWN
    )

async def handle_help(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Show help"""
    query = update.callback_query
    await query.answer()
    
    text = (
        "📖 **Qo'llanma**\n\n"
        "1. Hujjat turini tanlang\n"
        "2. Uslubni tanlang\n"
        "3. Sarlavha yuboring\n"
        "4. Matnni yuboring\n"
        "5. Hujjat tayyor!\n\n"
        "💡 Birinchi hujjat TEKIN!"
    )
    
    keyboard = [
        [InlineKeyboardButton("📝 Boshlash", callback_data="new_doc")],
        [InlineKeyboardButton("🏠 Bosh Menyu", callback_data="start")]
    ]
    
    await query.edit_message_text(
        text,
        reply_markup=InlineKeyboardMarkup(keyboard),
        parse_mode=ParseMode.MARKDOWN
    )

async def handle_new_doc(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Start new document"""
    query = update.callback_query
    await query.answer()
    
    keyboard = [
        [
            InlineKeyboardButton("📝 Referat", callback_data="doc_referat"),
            InlineKeyboardButton("📚 Kurs Ishi", callback_data="doc_kurs")
        ],
        [
            InlineKeyboardButton("📄 Maqola", callback_data="doc_maqola"),
            InlineKeyboardButton("🎯 Slide", callback_data="doc_slide")
        ]
    ]
    
    await query.edit_message_text(
        "🎓 Hujjat turini tanlang:",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )

async def handle_start_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Start from callback"""
    query = update.callback_query
    await query.answer()
    
    keyboard = [
        [
            InlineKeyboardButton("📝 Referat", callback_data="doc_referat"),
            InlineKeyboardButton("📚 Kurs Ishi", callback_data="doc_kurs")
        ],
        [
            InlineKeyboardButton("📄 Maqola", callback_data="doc_maqola"),
            InlineKeyboardButton("🎯 Slide", callback_data="doc_slide")
        ],
        [
            InlineKeyboardButton("📊 Statistika", callback_data="stats"),
            InlineKeyboardButton("❓ Yordam", callback_data="help")
        ]
    ]
    
    await query.edit_message_text(
        "🎓 **Akademik Yordamchi Bot**\n\n"
        "Hujjat turini tanlang:",
        reply_markup=InlineKeyboardMarkup(keyboard),
        parse_mode=ParseMode.MARKDOWN
    )

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Help command"""
    keyboard = [
        [InlineKeyboardButton("📝 Boshlash", callback_data="new_doc")],
        [InlineKeyboardButton("🏠 Bosh Menyu", callback_data="start")]
    ]
    
    await update.message.reply_text(
        "📖 **Qo'llanma**\n\n"
        "/start - Botni boshlash\n"
        "/help - Yordam",
        reply_markup=InlineKeyboardMarkup(keyboard),
        parse_mode=ParseMode.MARKDOWN
    )

def main():
    """Start bot"""
    application = Application.builder().token(BOT_TOKEN).build()
    
    # Commands
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("help", help_command))
    
    # Callbacks
    application.add_handler(CallbackQueryHandler(handle_doc_type, pattern="^doc_"))
    application.add_handler(CallbackQueryHandler(handle_style, pattern="^style_"))
    application.add_handler(CallbackQueryHandler(handle_stats, pattern="^stats$"))
    application.add_handler(CallbackQueryHandler(handle_help, pattern="^help$"))
    application.add_handler(CallbackQueryHandler(handle_new_doc, pattern="^new_doc$"))
    application.add_handler(CallbackQueryHandler(handle_start_callback, pattern="^start$"))
    
    # Messages
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    
    print("✅ BOT ISHGA TUSHDI!")
    print("🚀 READY TO USE!")
    
    application.run_polling()

if __name__ == "__main__":
    main()

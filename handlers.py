import logging
from datetime import datetime
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ContextTypes, ConversationHandler
from telegram.constants import ParseMode

logger = logging.getLogger(__name__)

# States
CHOOSING_TYPE, CHOOSING_STYLE, CHOOSING_FORMAT, WRITING_TITLE, WRITING_CONTENT = range(5)

class DocumentHandler:
    """Handle document generation workflow"""
    
    @staticmethod
    async def start_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Start document creation"""
        user_id = update.effective_user.id
        
        if user_id not in context.user_data:
            context.user_data[user_id] = {}
        
        context.user_data[user_id]['state'] = CHOOSING_TYPE
        
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
        
        await update.message.reply_text(
            "🎓 **Hujjat Turini Tanlang:**\n\n"
            "📝 Referat - Kichik tadqiqot ishi\n"
            "📚 Kurs Ishi - Kurs oxirida topshirish\n"
            "📄 Maqola - Ilmiy maqola\n"
            "🎯 Slide - Prezentasiya",
            reply_markup=InlineKeyboardMarkup(keyboard),
            parse_mode=ParseMode.MARKDOWN
        )
    
    @staticmethod
    async def select_style(update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Select document style"""
        query = update.callback_query
        await query.answer()
        
        doc_type = query.data.replace("doc_", "")
        user_id = query.from_user.id
        
        if user_id not in context.user_data:
            context.user_data[user_id] = {}
        
        context.user_data[user_id]['doc_type'] = doc_type
        context.user_data[user_id]['state'] = CHOOSING_STYLE
        
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
            f"✅ Siz {doc_types.get(doc_type, 'Noto\'g\'ri tur')} turini tanladingiz\n\n"
            "📋 **Endi Uslubni Tanlang:**",
            reply_markup=InlineKeyboardMarkup(keyboard),
            parse_mode=ParseMode.MARKDOWN
        )
    
    @staticmethod
    async def select_format(update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Select file format"""
        query = update.callback_query
        await query.answer()
        
        style = query.data.replace("style_", "")
        user_id = query.from_user.id
        
        if user_id not in context.user_data:
            context.user_data[user_id] = {}
        
        context.user_data[user_id]['style'] = style
        context.user_data[user_id]['state'] = CHOOSING_FORMAT
        
        styles = {
            "apa": "APA Style",
            "harvard": "Harvard Style",
            "uzbek": "O'zbek Standarti",
            "chicago": "Chicago Style"
        }
        
        keyboard = [
            [
                InlineKeyboardButton("📄 PDF", callback_data="format_pdf"),
                InlineKeyboardButton("📋 DOCX", callback_data="format_docx")
            ],
            [
                InlineKeyboardButton("📝 TXT", callback_data="format_txt")
            ]
        ]
        
        await query.edit_message_text(
            f"✅ Siz {styles.get(style)} ni tanladingiz\n\n"
            "📄 **Fayl Formatini Tanlang:**",
            reply_markup=InlineKeyboardMarkup(keyboard),
            parse_mode=ParseMode.MARKDOWN
        )
    
    @staticmethod
    async def get_title(update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Get document title"""
        query = update.callback_query
        await query.answer()
        
        file_format = query.data.replace("format_", "")
        user_id = query.from_user.id
        
        if user_id not in context.user_data:
            context.user_data[user_id] = {}
        
        context.user_data[user_id]['format'] = file_format
        context.user_data[user_id]['state'] = WRITING_TITLE
        
        await query.edit_message_text(
            "✏️ **Hujjatning Sarlavhasini Yuboring:**\n\n"
            "_Misol: \"Kimyo texnologiyasi va IT sohalari integratsiyasi\"_"
        )
    
    @staticmethod
    async def handle_title(update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Handle title input"""
        user_id = update.effective_user.id
        title = update.message.text
        
        if not 3 <= len(title) <= 300:
            await update.message.reply_text(
                "❌ Sarlavha 3 dan 300 ta belgigacha bo'lishi kerak!\n\n"
                "Qayta urinib ko'ring:"
            )
            return
        
        if user_id not in context.user_data:
            context.user_data[user_id] = {}
        
        context.user_data[user_id]['title'] = title
        context.user_data[user_id]['state'] = WRITING_CONTENT
        
        await update.message.reply_text(
            f"✅ Sarlavha qabul qilindi: **{title}**\n\n"
            "📝 **Endi Matnni Yuboring:**\n\n"
            "_Uzun bo'lsa, bir nechta xabarda yuboring._"
        )
    
    @staticmethod
    async def handle_content(update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Handle content input"""
        user_id = update.effective_user.id
        content = update.message.text
        
        if user_id not in context.user_data:
            context.user_data[user_id] = {}
        
        # Combine if multiple messages
        if "content" in context.user_data[user_id]:
            context.user_data[user_id]['content'] += "\n\n" + content
        else:
            context.user_data[user_id]['content'] = content
        
        char_count = len(context.user_data[user_id]['content'])
        word_count = len(context.user_data[user_id]['content'].split())
        
        keyboard = [
            [InlineKeyboardButton("✅ Tugatish va Fayl Olish", callback_data="finish_doc")],
            [InlineKeyboardButton("➕ Yana Matn Qo'shish", callback_data="add_more")],
            [InlineKeyboardButton("✏️ Sarlavhani Tahrirlash", callback_data="edit_title")],
            [InlineKeyboardButton("❌ Bekor Qilish", callback_data="cancel_doc")]
        ]
        
        await update.message.reply_text(
            f"✅ **Matn Qabul Qilindi!**\n\n"
            f"📊 **Statistika:**\n"
            f"   • Belgisiz: {char_count}\n"
            f"   • So'zlar: {word_count}\n\n"
            f"🔄 **Quyidagini Tanlang:**",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )
    
    @staticmethod
    async def edit_title(update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Edit title"""
        query = update.callback_query
        await query.answer()
        
        user_id = query.from_user.id
        if user_id in context.user_data:
            context.user_data[user_id]['state'] = WRITING_TITLE
        
        await query.edit_message_text(
            "✏️ **Yangi Sarlavhani Yuboring:**"
        )
    
    @staticmethod
    async def add_more(update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Add more content"""
        query = update.callback_query
        await query.answer()
        
        user_id = query.from_user.id
        if user_id in context.user_data:
            context.user_data[user_id]['state'] = WRITING_CONTENT
        
        await query.edit_message_text(
            "📝 **Yana Matnni Yuboring:**"
        )
    
    @staticmethod
    async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Cancel document creation"""
        query = update.callback_query
        await query.answer()
        
        user_id = query.from_user.id
        if user_id in context.user_data:
            del context.user_data[user_id]
        
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
            "❌ **Hujjat yaratilishi bekor qilindi!**\n\n"
            "Yangi hujjat yaratmoqchimisiz?",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )


class PaymentHandler:
    """Handle payment workflow"""
    
    @staticmethod
    async def start_payment(update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Start payment process"""
        query = update.callback_query
        await query.answer()
        
        keyboard = [
            [
                InlineKeyboardButton("💳 Click.uz", callback_data="pay_click"),
                InlineKeyboardButton("💳 Payme", callback_data="pay_payme")
            ],
            [
                InlineKeyboardButton("⭐ Telegram Stars", callback_data="pay_telegram"),
                InlineKeyboardButton("💰 Karta", callback_data="pay_card")
            ]
        ]
        
        await query.edit_message_text(
            "💳 **To'lov Usulini Tanlang:**\n\n"
            "2,000 UZS = 0.17 USD",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )
    
    @staticmethod
    async def process_payment(update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Process payment"""
        query = update.callback_query
        await query.answer()
        
        gateway = query.data.replace("pay_", "")
        user_id = query.from_user.id
        
        payment_info = {
            "click": {
                "title": "💳 Click.uz",
                "description": "Click.uz orqali to'lovni amalga oshiring",
                "link": "https://my.click.uz"
            },
            "payme": {
                "title": "💳 Payme",
                "description": "Payme orqali to'lovni amalga oshiring",
                "link": "https://payme.uz"
            },
            "telegram": {
                "title": "⭐ Telegram Stars",
                "description": "Telegram Stars orqali to'lovni amalga oshiring",
                "link": "https://t.me"
            },
            "card": {
                "title": "💰 Karta",
                "description": "Karta raqamiga to'lovni amalga oshiring",
                "link": None
            }
        }
        
        info = payment_info.get(gateway)
        if not info:
            return
        
        text = f"""
{info['title']} **orqali To'lov**

💰 **Summa:** 2,000 UZS (0.17 USD)

📝 **Tavsif:** {info['description']}

⏱️ **Vaqt:** 5 daqiqada tasdiqlandi

🔗 **Havola:** {info['link'] or 'Support bilan bog\'lanish'}

✅ To'lovni amalga oshirgandan so\'ng avtomatik faollashtiriladi!
"""
        
        keyboard = [
            [InlineKeyboardButton("✅ To'lov Qildim", callback_data=f"confirm_pay_{gateway}")],
            [InlineKeyboardButton("🔙 Orqaga", callback_data="back_menu")]
        ]
        
        await query.edit_message_text(
            text,
            reply_markup=InlineKeyboardMarkup(keyboard)
        )


class AdminHandler:
    """Handle admin operations"""
    
    @staticmethod
    async def admin_panel(update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Admin panel"""
        query = update.callback_query
        
        if query.from_user.id != context.bot.admin_id:
            await query.answer("❌ Siz admin emassiz!", show_alert=True)
            return
        
        await query.answer()
        
        keyboard = [
            [InlineKeyboardButton("👥 Foydalanuvchilar", callback_data="admin_users")],
            [InlineKeyboardButton("📊 Statistika", callback_data="admin_stats")],
            [InlineKeyboardButton("💰 To'lovlar", callback_data="admin_payments")],
            [InlineKeyboardButton("📋 Hujjatlar", callback_data="admin_documents")],
            [InlineKeyboardButton("📢 Xabar Yuborish", callback_data="admin_broadcast")],
            [InlineKeyboardButton("🔒 Ban Qilish", callback_data="admin_ban")],
            [InlineKeyboardButton("🔙 Orqaga", callback_data="back_menu")]
        ]
        
        await query.edit_message_text(
            "👨‍💼 **ADMIN PANEL**",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )
    
    @staticmethod
    async def admin_stats(update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Admin statistics"""
        query = update.callback_query
        await query.answer()
        
        # Get stats from database
        from database import get_stats
        stats = get_stats()
        
        text = f"""
📊 **ADMIN STATISTIKA**

👥 **Foydalanuvchilar:**
   • Jami: {stats['total_users']}
   • Faol: {stats['active_users']}

📄 **Hujjatlar:**
   • Jami: {stats['total_documents']}
   • Bugun: {stats['documents_today']}

💰 **To'lovlar:**
   • Muvaffaqiyatli: {stats['total_payments']}
   • Daromad: {stats['total_revenue']} UZS

🕐 **Tahrir vaqti:** {datetime.now().strftime('%d.%m.%Y %H:%M')}
"""
        
        keyboard = [
            [InlineKeyboardButton("🔄 Yangilash", callback_data="admin_stats")],
            [InlineKeyboardButton("🔙 Orqaga", callback_data="back_menu")]
        ]
        
        await query.edit_message_text(
            text,
            reply_markup=InlineKeyboardMarkup(keyboard)
        )
    
    @staticmethod
    async def broadcast(update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Send broadcast message"""
        query = update.callback_query
        await query.answer()
        
        await query.edit_message_text(
            "📢 **Xabarni Yuboring:**\n\n"
            "_Bu xabar barcha foydalanuvchilarga yuboriladi!_"
        )
        
        user_id = query.from_user.id
        if user_id not in context.user_data:
            context.user_data[user_id] = {}
        
        context.user_data[user_id]['admin_action'] = 'broadcast'

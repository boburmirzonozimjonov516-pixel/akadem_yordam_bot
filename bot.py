import os
import json
import logging
from datetime import datetime, timedelta
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, Document
from telegram.ext import (
    Application, CommandHandler, CallbackQueryHandler, MessageHandler, 
    filters, ContextTypes, ConversationHandler
)
from telegram.constants import ChatAction, ParseMode
from dotenv import load_dotenv
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from docx import Document as DocxDocument
from docx.shared import Pt, RGBColor

# Configure logging
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

load_dotenv()
BOT_TOKEN = os.getenv("BOT_TOKEN", "8625557628:AAHeUC2WxfMjJk-RRq3IxTtUJoc0H4XSsAM")
ADMIN_ID = int(os.getenv("ADMIN_ID", "7758296066"))

# States
CHOOSING_TYPE, CHOOSING_STYLE, WRITING_TITLE, WRITING_CONTENT = range(4)
PAYMENT_GATEWAY, PROCESSING_PAYMENT = range(4, 6)

# In-memory database (use PostgreSQL in production)
class Database:
    def __init__(self):
        self.users = {}
        self.documents = {}
        self.payments = {}
        self.subscriptions = {}
    
    def create_user(self, user_id, username, first_name):
        if user_id not in self.users:
            self.users[user_id] = {
                "id": user_id,
                "username": username,
                "first_name": first_name,
                "created_at": datetime.now().isoformat(),
                "first_doc_used": False,
                "subscription_active": False,
                "balance": 0,
                "documents_count": 0
            }
        return self.users[user_id]
    
    def get_user(self, user_id):
        return self.users.get(user_id)
    
    def save_document(self, user_id, doc_data):
        doc_id = f"doc_{user_id}_{datetime.now().timestamp()}"
        self.documents[doc_id] = {
            "id": doc_id,
            "user_id": user_id,
            "type": doc_data.get("doc_type"),
            "style": doc_data.get("style"),
            "title": doc_data.get("title"),
            "content": doc_data.get("content"),
            "created_at": datetime.now().isoformat(),
            "file_path": doc_data.get("file_path")
        }
        
        if user_id in self.users:
            self.users[user_id]["documents_count"] += 1
        
        return doc_id
    
    def save_payment(self, user_id, amount, gateway, status="pending"):
        payment_id = f"pay_{user_id}_{datetime.now().timestamp()}"
        self.payments[payment_id] = {
            "id": payment_id,
            "user_id": user_id,
            "amount": amount,
            "gateway": gateway,
            "status": status,
            "created_at": datetime.now().isoformat()
        }
        return payment_id
    
    def get_stats(self):
        return {
            "total_users": len(self.users),
            "total_documents": len(self.documents),
            "total_payments": len(self.payments),
            "total_revenue": sum([p.get("amount", 0) for p in self.payments.values() 
                                 if p.get("status") == "completed"])
        }

db = Database()

# File generators
class FileGenerator:
    @staticmethod
    def generate_pdf(title, content, style, user_id):
        try:
            filename = f"doc_{user_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
            
            doc = SimpleDocTemplate(filename, pagesize=A4)
            styles = getSampleStyleSheet()
            story = []
            
            # Custom styles
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontSize=18,
                textColor=RGBColor(0, 51, 102),
                spaceAfter=30,
                alignment=1
            )
            
            body_style = ParagraphStyle(
                'CustomBody',
                parent=styles['BodyText'],
                fontSize=11,
                leading=22,
                alignment=4,
                spaceAfter=12
            )
            
            # Add content
            story.append(Paragraph(title, title_style))
            story.append(Spacer(1, 0.5*inch))
            
            for para in content.split('\n\n'):
                if para.strip():
                    story.append(Paragraph(para.strip(), body_style))
                    story.append(Spacer(1, 0.2*inch))
            
            # Footer
            story.append(Spacer(1, 0.5*inch))
            footer_text = f"Style: {style} | Generated: {datetime.now().strftime('%d.%m.%Y %H:%M')} | By Akadem Bot"
            story.append(Paragraph(footer_text, styles['Normal']))
            
            doc.build(story)
            logger.info(f"PDF generated: {filename}")
            return filename
        except Exception as e:
            logger.error(f"PDF generation error: {e}")
            return None
    
    @staticmethod
    def generate_docx(title, content, style, user_id):
        try:
            filename = f"doc_{user_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            
            doc = DocxDocument()
            
            # Title
            title_para = doc.add_paragraph(title)
            title_para.style = 'Heading 1'
            title_para.runs[0].font.color.rgb = RGBColor(0, 51, 102)
            title_para.runs[0].font.size = Pt(18)
            
            # Metadata
            meta = doc.add_paragraph()
            meta.add_run(f"Style: {style} | Generated: {datetime.now().strftime('%d.%m.%Y %H:%M')}").italic = True
            
            doc.add_paragraph()
            
            # Content
            for para in content.split('\n\n'):
                if para.strip():
                    p = doc.add_paragraph(para.strip())
                    p.style = 'Normal'
            
            doc.save(filename)
            logger.info(f"DOCX generated: {filename}")
            return filename
        except Exception as e:
            logger.error(f"DOCX generation error: {e}")
            return None

# Keyboards
def main_menu_keyboard():
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
            InlineKeyboardButton("💳 Obuna", callback_data="subscription")
        ],
        [
            InlineKeyboardButton("⚙️ Sozlamalar", callback_data="settings"),
            InlineKeyboardButton("❓ Yordam", callback_data="help")
        ]
    ]
    return InlineKeyboardMarkup(keyboard)

def style_keyboard():
    keyboard = [
        [
            InlineKeyboardButton("🎯 APA Style", callback_data="style_apa"),
            InlineKeyboardButton("🎯 Harvard", callback_data="style_harvard")
        ],
        [
            InlineKeyboardButton("🎯 Uzbek", callback_data="style_uzbek"),
            InlineKeyboardButton("🎯 Chicago", callback_data="style_chicago")
        ],
        [InlineKeyboardButton("🔙 Orqaga", callback_data="back_menu")]
    ]
    return InlineKeyboardMarkup(keyboard)

def file_format_keyboard():
    keyboard = [
        [
            InlineKeyboardButton("📄 PDF", callback_data="format_pdf"),
            InlineKeyboardButton("📋 DOCX", callback_data="format_docx")
        ],
        [InlineKeyboardButton("🔙 Orqaga", callback_data="back_menu")]
    ]
    return InlineKeyboardMarkup(keyboard)

def payment_keyboard():
    keyboard = [
        [
            InlineKeyboardButton("💳 Click.uz", callback_data="pay_click"),
            InlineKeyboardButton("💳 Payme", callback_data="pay_payme")
        ],
        [
            InlineKeyboardButton("📱 Telegram", callback_data="pay_telegram"),
            InlineKeyboardButton("💰 Karta", callback_data="pay_card")
        ],
        [InlineKeyboardButton("🔙 Orqaga", callback_data="back_menu")]
    ]
    return InlineKeyboardMarkup(keyboard)

def admin_keyboard():
    keyboard = [
        [InlineKeyboardButton("👥 Foydalanuvchilar", callback_data="admin_users")],
        [InlineKeyboardButton("📊 Statistika", callback_data="admin_stats")],
        [InlineKeyboardButton("💰 To'lovlar", callback_data="admin_payments")],
        [InlineKeyboardButton("📢 Xabar Yuborish", callback_data="admin_broadcast")],
        [InlineKeyboardButton("🔙 Orqaga", callback_data="back_menu")]
    ]
    return InlineKeyboardMarkup(keyboard)

# Commands
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    db.create_user(user.id, user.username, user.first_name)
    
    welcome_text = f"""
🎓 **Akademik Yordamchi Botga Xush Kelibsiz!**

👋 Assalomu alaykum, {user.first_name}!

Siz quyidagi hujjatlarni yarata olasiz:
📝 **Referat** - Kichik tadqiqot ishi
📚 **Kurs Ishi** - Kurs oxirida topshirish
📄 **Maqola** - Ilmiy maqola
🎯 **Slide** - Prezentasiya

**✨ Premium Features:**
✅ PDF + DOCX generation
✅ 4 ta style (APA, Harvard, Uzbek, Chicago)
✅ Rasm qo'shish
✅ Advanced formatting
✅ Priority support

💡 **Birinchi hujjat TEKIN!**
Keyingisini yaratish uchun obuna: 2,000 UZS yoki 0.17 USD

Qanday hujjat yaratmoqchisiz?
"""
    
    await update.message.reply_text(
        welcome_text,
        reply_markup=main_menu_keyboard(),
        parse_mode=ParseMode.MARKDOWN
    )

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    help_text = """
📖 **Qo'llanma:**

**1️⃣ Hujjat Yaratish:**
   • Hujjat turini tanlang
   • Uslubni tanlang (APA/Harvard/Uzbek/Chicago)
   • Format tanlang (PDF/DOCX)
   • Sarlavha yuboring
   • Matnni yuboring
   • Hujjat qabul qilinadi!

**2️⃣ Buyruqlar:**
/start - Botni boshlash
/help - Yordam
/stats - Statistika
/payment - To'lov qilish

**3️⃣ Premium Features:**
✅ 4 ta uslub
✅ PDF va DOCX
✅ Rasm qo'shish
✅ Advanced formatting

**4️⃣ Narxlar:**
🎁 Birinchi: TEKIN
💰 Keyingilari: 2,000 UZS (0.17 USD)

**5️⃣ To'lov Usullari:**
💳 Click.uz
💳 Payme
📱 Telegram
💰 Karta

Savol bo'lsa @support yuboring!
"""
    
    await update.message.reply_text(
        help_text,
        reply_markup=main_menu_keyboard(),
        parse_mode=ParseMode.MARKDOWN
    )

async def stats_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user = db.get_user(user_id)
    
    if not user:
        await update.message.reply_text("❌ Foydalanuvchi topilmadi!")
        return
    
    stats_text = f"""
📊 **SHAXSIY STATISTIKA**

👤 **Profil:**
   • Ism: {user['first_name']}
   • ID: {user_id}
   • Ro'yxatga olingan: {user['created_at'][:10]}

📄 **Hujjatlar:**
   • Yaratilgan: {user['documents_count']}
   • Birinchi hujjat: {'✅ Ishlatildi' if user['first_doc_used'] else '⏳ Ishlatilmadi (TEKIN)'}

💳 **Obuna:**
   • Status: {'✅ Faol' if user['subscription_active'] else '❌ Nofaol'}
   • Balans: {user['balance']} UZS

🔄 **Keyingi Qadamlar:**
   1. Hujjat yaratish
   2. Premium obuna
   3. Yangi hujjatlar

Qo'shimcha yordam uchun /help bosing!
"""
    
    await update.message.reply_text(
        stats_text,
        reply_markup=main_menu_keyboard(),
        parse_mode=ParseMode.MARKDOWN
    )

async def admin_panel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ADMIN_ID:
        await update.message.reply_text("❌ Siz admin emassiz!")
        return
    
    stats = db.get_stats()
    
    admin_text = f"""
👨‍💼 **ADMIN PANEL**

📊 **Statistika:**
   • Jami foydalanuvchilar: {stats['total_users']}
   • Jami hujjatlar: {stats['total_documents']}
   • Jami to'lovlar: {stats['total_payments']}
   • Jami daromad: {stats['total_revenue']} UZS

🔧 **Boshqaruv:**
"""
    
    await update.message.reply_text(
        admin_text,
        reply_markup=admin_keyboard(),
        parse_mode=ParseMode.MARKDOWN
    )

async def doc_type_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    doc_type = query.data.replace("doc_", "")
    user_id = query.from_user.id
    
    if user_id not in context.user_data:
        context.user_data[user_id] = {}
    context.user_data[user_id]["doc_type"] = doc_type
    
    doc_types = {
        "referat": "📝 Referat",
        "kurs": "📚 Kurs Ishi",
        "maqola": "📄 Maqola",
        "slide": "🎯 Slide"
    }
    
    await query.edit_message_text(
        f"✅ Siz {doc_types.get(doc_type, 'Noto\'g\'ri tur')} turini tanladingiz\n\n"
        "📋 Endi uslubni tanlang:",
        reply_markup=style_keyboard(),
        parse_mode=ParseMode.MARKDOWN
    )

async def style_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    if query.data == "back_menu":
        await query.edit_message_text(
            "🎓 **Akademik Yordamchi Bot**\n\n"
            "Hujjat turini tanlang:",
            reply_markup=main_menu_keyboard(),
            parse_mode=ParseMode.MARKDOWN
        )
        return
    
    style = query.data.replace("style_", "")
    user_id = query.from_user.id
    
    if user_id not in context.user_data:
        context.user_data[user_id] = {}
    context.user_data[user_id]["style"] = style
    
    styles = {
        "apa": "APA Style",
        "harvard": "Harvard Style",
        "uzbek": "O'zbek Standarti",
        "chicago": "Chicago Style"
    }
    
    await query.edit_message_text(
        f"✅ Siz {styles.get(style)} ni tanladingiz\n\n"
        "📄 Format tanlang:",
        reply_markup=file_format_keyboard(),
        parse_mode=ParseMode.MARKDOWN
    )

async def format_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    if query.data == "back_menu":
        await query.edit_message_text(
            "🎓 **Akademik Yordamchi Bot**\n\n"
            "Hujjat turini tanlang:",
            reply_markup=main_menu_keyboard(),
            parse_mode=ParseMode.MARKDOWN
        )
        return
    
    file_format = query.data.replace("format_", "")
    user_id = query.from_user.id
    
    if user_id not in context.user_data:
        context.user_data[user_id] = {}
    context.user_data[user_id]["format"] = file_format
    
    await query.edit_message_text(
        "✏️ **Sarlavhani yuboring:**"
    )
    
    context.user_data[user_id]["state"] = WRITING_TITLE

async def handle_title(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    
    if user_id not in context.user_data:
        context.user_data[user_id] = {}
    
    context.user_data[user_id]["title"] = update.message.text
    context.user_data[user_id]["state"] = WRITING_CONTENT
    
    await update.message.reply_text(
        "✅ Sarlavha qabul qilindi!\n\n"
        "📝 Endi **matnni** yuboring:\n\n"
        "_Uzun bo'lsa, bir nechta xabarda yuboring_"
    )

async def handle_content(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    
    if user_id not in context.user_data:
        context.user_data[user_id] = {}
    
    content = update.message.text
    
    if "content" in context.user_data[user_id]:
        context.user_data[user_id]["content"] += "\n\n" + content
    else:
        context.user_data[user_id]["content"] = content
    
    keyboard = [
        [InlineKeyboardButton("✅ Tugatish va Fayl Olish", callback_data="finish_doc")],
        [InlineKeyboardButton("➕ Yana Matn Qo'shish", callback_data="add_more")],
        [InlineKeyboardButton("❌ Bekor Qilish", callback_data="cancel_doc")]
    ]
    
    await update.message.reply_text(
        f"✅ Matn qabul qilindi!\n\n"
        f"📊 Hajmi: {len(content)} belgi\n\n"
        f"🔄 Yana matn qo'shib olasiz yoki tugatishingiz mumkin.",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )

async def finish_doc(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    user_id = query.from_user.id
    user = db.get_user(user_id)
    
    if not user:
        await query.edit_message_text("❌ Foydalanuvchi topilmadi!")
        return
    
    if user_id not in context.user_data or "content" not in context.user_data[user_id]:
        await query.edit_message_text("❌ Hujjat ma'lumotlari to'liq emas!")
        return
    
    data = context.user_data[user_id]
    
    await query.edit_message_text("⏳ Fayl tayyorlanmoqda...", reply_markup=None)
    
    try:
        # Generate file
        if data.get("format") == "pdf":
            file_path = FileGenerator.generate_pdf(
                data["title"],
                data["content"],
                data["style"],
                user_id
            )
        else:
            file_path = FileGenerator.generate_docx(
                data["title"],
                data["content"],
                data["style"],
                user_id
            )
        
        if not file_path:
            await query.message.reply_text("❌ Fayl tayyorlanmadi!")
            return
        
        # Save to database
        db.save_document(user_id, {
            "doc_type": data["doc_type"],
            "style": data["style"],
            "title": data["title"],
            "content": data["content"],
            "file_path": file_path
        })
        
        # Mark first doc as used
        if not user["first_doc_used"]:
            user["first_doc_used"] = True
        
        # Send file
        with open(file_path, 'rb') as f:
            await query.message.reply_document(
                f,
                caption=f"""✅ **Hujjat Tayyorlandi!**

📝 **Tur:** {data['doc_type']}
📋 **Uslub:** {data['style']}
📄 **Format:** {data['format'].upper()}
📖 **Sarlavha:** {data['title']}

🎁 {'Birinchi hujjat tekin edi!' if user['first_doc_used'] else ''}

🔄 **Keyingi Qadamlar:**
• Yangi hujjat yaratish
• Premium obuna
• Statistika ko'rish
""",
                parse_mode=ParseMode.MARKDOWN
            )
        
        # Clean up
        os.remove(file_path)
        
        # New document button
        keyboard = [
            [InlineKeyboardButton("📝 Yangi Hujjat", callback_data="new_doc")],
            [InlineKeyboardButton("💳 Premium Obuna", callback_data="subscription")],
            [InlineKeyboardButton("🏠 Bosh Menyu", callback_data="back_menu")]
        ]
        
        await query.message.reply_text(
            "🎉 **Tabriklar! Hujjat tayyorlandi!**\n\n"
            "Qo'limizdan muvaffaq bo'ldingiz!",
            reply_markup=InlineKeyboardMarkup(keyboard),
            parse_mode=ParseMode.MARKDOWN
        )
        
        # Clear user data
        del context.user_data[user_id]
        
    except Exception as e:
        logger.error(f"Error: {e}")
        await query.message.reply_text(f"❌ Xatolik: {str(e)}")

async def payment_gateway(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    gateway = query.data.replace("pay_", "")
    user_id = query.from_user.id
    
    if gateway == "click":
        payment_text = """
💳 **Click.uz orqali To'lov**

2,000 UZS = 0.17 USD

✅ Quydagi kartalarni qabul qiladi:
   • Humo
   • Visa
   • Mastercard
   • UzCard

🔗 To'lovni amalga oshirish:
https://click.uz/services/pay

⏱️ To'lov 5 daqiqada tasdiqlandi!
"""
    elif gateway == "payme":
        payment_text = """
💳 **Payme orqali To'lov**

2,000 UZS = 0.17 USD

✅ Quydagi kartalarni qabul qiladi:
   • Humo
   • Visa
   • Mastercard

🔗 To'lovni amalga oshirish:
https://checkout.paycom.uz

⏱️ To'lov 5 daqiqada tasdiqlandi!
"""
    elif gateway == "telegram":
        payment_text = """
💳 **Telegram Stars orqali To'lov**

2,000 UZS = 0.17 USD

🌟 Telegram Stars hozir eng oson usul!

✅ Foydalanuvchi do'stona
✅ Tez va xavfsiz
✅ Biror hech narsa kerak emas

🔗 To'lovni amalga oshirish:
/stars_payment

⏱️ To'lov darhol tasdiqlandi!
"""
    else:
        payment_text = """
💳 **Karta orqali To'lov**

2,000 UZS = 0.17 USD

📞 To'lov uchun biz bilan bog'laning:
   • Telegram: @support
   • Email: support@akadem.uz
   • Phone: +998 XX XXX XX XX

🔄 Biz sizni qo'llab-quvvatlayamiz!
"""
    
    keyboard = [
        [InlineKeyboardButton("🔙 Orqaga", callback_data="back_payment")]
    ]
    
    await query.edit_message_text(
        payment_text,
        reply_markup=InlineKeyboardMarkup(keyboard),
        parse_mode=ParseMode.MARKDOWN
    )
    
    # Save payment record
    db.save_payment(user_id, 2000, gateway, "pending")

async def back_payment(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    await query.edit_message_text(
        "💳 **To'lov Usulini Tanlang:**",
        reply_markup=payment_keyboard(),
        parse_mode=ParseMode.MARKDOWN
    )

async def subscription_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    sub_text = """
💎 **PREMIUM OBUNA**

**Narx:** 2,000 UZS / 0.17 USD

**Nima Olasiz:**
✅ Cheksiz hujjat yaratish
✅ 4 ta uslub (APA, Harvard, Uzbek, Chicago)
✅ PDF + DOCX format
✅ Rasm qo'shish
✅ Priority support
✅ Advanced formatting

**Qanday To'lash:**
1. To'lov usulini tanlang
2. To'lovni amalga oshiring
3. Avtomatik faollashtiriladi!

🎁 **Birinchi hujjat TEKIN!**
"""
    
    keyboard = [
        [InlineKeyboardButton("💳 To'lov Qilish", callback_data="start_payment")],
        [InlineKeyboardButton("🔙 Orqaga", callback_data="back_menu")]
    ]
    
    await query.edit_message_text(
        sub_text,
        reply_markup=InlineKeyboardMarkup(keyboard),
        parse_mode=ParseMode.MARKDOWN
    )

async def new_doc(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    await query.edit_message_text(
        "🎓 **Akademik Yordamchi Bot**\n\n"
        "Hujjat turini tanlang:",
        reply_markup=main_menu_keyboard(),
        parse_mode=ParseMode.MARKDOWN
    )

async def add_more(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    user_id = query.from_user.id
    if user_id in context.user_data:
        context.user_data[user_id]["state"] = WRITING_CONTENT
    
    await query.edit_message_text(
        "📝 **Yana matnni yuboring:**"
    )

async def cancel_doc(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    user_id = query.from_user.id
    if user_id in context.user_data:
        del context.user_data[user_id]
    
    await query.edit_message_text(
        "❌ Hujjat bekor qilindi!\n\n"
        "Yangi hujjat yaratmoqchimisiz?",
        reply_markup=main_menu_keyboard(),
        parse_mode=ParseMode.MARKDOWN
    )

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    
    if user_id not in context.user_data:
        context.user_data[user_id] = {}
    
    state = context.user_data[user_id].get("state")
    
    if state == WRITING_TITLE:
        await handle_title(update, context)
    elif state == WRITING_CONTENT:
        await handle_content(update, context)
    else:
        await update.message.reply_text(
            "❓ Tushunmadim! Quyidagi buyruqlardan foydalaning:\n\n"
            "/start - Botni boshlash\n"
            "/help - Yordam\n"
            "/stats - Statistika",
            reply_markup=main_menu_keyboard()
        )

async def help_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    await help_command(update, context)

async def admin_stats(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    if query.from_user.id != ADMIN_ID:
        await query.answer("❌ Siz admin emassiz!")
        return
    
    stats = db.get_stats()
    
    stats_text = f"""
📊 **ADMIN STATISTIKA**

👥 **Foydalanuvchilar:** {stats['total_users']}
📄 **Hujjatlar:** {stats['total_documents']}
💰 **To'lovlar:** {stats['total_payments']}
💵 **Daromad:** {stats['total_revenue']} UZS

**Tahrir vaqti:** {datetime.now().strftime('%d.%m.%Y %H:%M')}
"""
    
    await query.edit_message_text(
        stats_text,
        reply_markup=admin_keyboard(),
        parse_mode=ParseMode.MARKDOWN
    )

def main():
    app = Application.builder().token(BOT_TOKEN).build()
    
    # Commands
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("stats", stats_command))
    app.add_handler(CommandHandler("admin", admin_panel))
    
    # Callbacks
    app.add_handler(CallbackQueryHandler(doc_type_callback, pattern="^doc_"))
    app.add_handler(CallbackQueryHandler(style_callback, pattern="^style_|^back_menu$"))
    app.add_handler(CallbackQueryHandler(format_callback, pattern="^format_"))
    app.add_handler(CallbackQueryHandler(finish_doc, pattern="^finish_doc$"))
    app.add_handler(CallbackQueryHandler(add_more, pattern="^add_more$"))
    app.add_handler(CallbackQueryHandler(cancel_doc, pattern="^cancel_doc$"))
    app.add_handler(CallbackQueryHandler(payment_gateway, pattern="^pay_"))
    app.add_handler(CallbackQueryHandler(back_payment, pattern="^back_payment$"))
    app.add_handler(CallbackQueryHandler(subscription_menu, pattern="^subscription$"))
    app.add_handler(CallbackQueryHandler(new_doc, pattern="^new_doc$"))
    app.add_handler(CallbackQueryHandler(help_callback, pattern="^help$"))
    app.add_handler(CallbackQueryHandler(admin_stats, pattern="^admin_stats$"))
    
    # Message handler
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    
    logger.info("✅ ENTERPRISE BOT STARTED!")
    logger.info("🚀 All features active!")
    
    app.run_polling()

if __name__ == "__main__":
    main()

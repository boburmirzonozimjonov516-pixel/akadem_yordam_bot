import os
import asyncio
import aiohttp
import json
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, CallbackQueryHandler, MessageHandler, filters, ContextTypes
from telegram.constants import ParseMode
from dotenv import load_dotenv
from datetime import datetime
from io import BytesIO

# File generation
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.units import cm
from docx import Document as DocxDocument
from docx.shared import Pt, Inches
from pptx import Presentation
from pptx.util import Inches as PptxInches, Pt as PptxPt, Emu
from pptx.dml.color import RGBColor as PptxRGB
from pptx.enum.text import PP_ALIGN

load_dotenv()

BOT_TOKEN = os.getenv("BOT_TOKEN", "8625557628:AAHeUC2WxfMjJk-RRq3IxTtUJoc0H4XSsAM")
ADMIN_ID = int(os.getenv("ADMIN_ID", "7758296066"))
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")

user_data_store = {}

# ==================== AI CONTENT GENERATION ====================

async def generate_ai_content(topic, doc_type, style, slide_count=8):
    """Generate content using Gemini AI"""
    
    if not GEMINI_API_KEY:
        # Fallback: generate simple content without AI
        return generate_simple_content(topic, doc_type, style, slide_count)
    
    if doc_type == "slide":
        prompt = f"""
        "{topic}" mavzusida {slide_count} ta slayd uchun prezentatsiya mazmunini yoz.
        Uslub: {style}
        
        Har bir slayd uchun:
        SLAYD_1: [Sarlavha]
        MAZMUN_1: [2-3 qator matn]
        
        FAQAT O'ZBEKCHA yoz. JSON emas, oddiy matn.
        """
    else:
        doc_names = {
            "referat": "Referat",
            "kurs": "Kurs ishi",
            "maqola": "Ilmiy maqola"
        }
        prompt = f"""
        "{topic}" mavzusida {doc_names.get(doc_type, 'hujjat')} yoz.
        Uslub: {style}
        Hajm: 500-800 so'z
        
        Tuzilma:
        1. Kirish
        2. Asosiy qism (3-4 bo'lim)
        3. Xulosa
        
        FAQAT O'ZBEKCHA yoz.
        """
    
    try:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key={GEMINI_API_KEY}"
        
        payload = {
            "contents": [{"parts": [{"text": prompt}]}]
        }
        
        async with aiohttp.ClientSession() as session:
            async with session.post(url, json=payload, timeout=aiohttp.ClientTimeout(total=30)) as resp:
                if resp.status == 200:
                    data = await resp.json()
                    return data["candidates"][0]["content"]["parts"][0]["text"]
                else:
                    return generate_simple_content(topic, doc_type, style, slide_count)
    except:
        return generate_simple_content(topic, doc_type, style, slide_count)


def generate_simple_content(topic, doc_type, style, slide_count=8):
    """Generate content without AI (fallback)"""
    
    if doc_type == "slide":
        slides = []
        slides.append(f"SLAYD_1: {topic}")
        slides.append(f"MAZMUN_1: Ushbu prezentatsiya {topic} mavzusiga bag'ishlangan. Biz bu sohaning asosiy jihatlarini ko'rib chiqamiz.")
        
        sections = [
            ("Kirish", f"{topic} - bu zamonaviy dunyoda muhim ahamiyat kasb etuvchi soha. Uning asosiy xususiyatlari va ahamiyatini tushunish zarur."),
            ("Asosiy tushunchalar", f"{topic} sohasidagi asosiy tushunchalar va atamalar bilan tanishib chiqamiz. Bu bilimlar keyingi tadqiqotlar uchun asos bo'ladi."),
            ("Tarixiy rivojlanish", f"{topic} ning rivojlanish tarixi uzoq o'tmishga ega. Har bir davr o'zining muhim kashfiyotlari bilan ajralib turadi."),
            ("Hozirgi holat", f"Hozirgi kunda {topic} sohasida katta yutuqlarga erishilgan. Yangi texnologiyalar va yondashuvlar faol qo'llanilmoqda."),
            ("Muammolar va yechimlar", f"{topic} sohasida bir qator muammolar mavjud. Bular yechish yo'llari faol izlanmoqda va yangi yondashuvlar taklif etilmoqda."),
            ("Kelajak istiqbollari", f"{topic} ning kelajagi juda istiqbolli. Yangi tadqiqotlar va innovatsiyalar bu sohani yanada rivojlantiradi."),
            ("Xulosa", f"{topic} mavzusini o'rganish natijasida ko'pgina muhim xulosalarga keldik. Bu bilimlar amaliyotda keng qo'llanilishi mumkin."),
        ]
        
        for i, (title, content) in enumerate(sections[:slide_count-1], 2):
            slides.append(f"SLAYD_{i}: {title}")
            slides.append(f"MAZMUN_{i}: {content}")
        
        return "\n".join(slides)
    
    else:
        return f"""KIRISH

{topic} - bu zamonaviy fanda va amaliyotda katta ahamiyat kasb etuvchi mavzu. Ushbu {doc_type} da biz bu mavzuning asosiy jihatlarini ko'rib chiqamiz va uning mohiyatini tushuntirishga harakat qilamiz.

ASOSIY QISM

{topic} sohasini o'rganish bir necha sabablarga ko'ra muhimdir. Birinchidan, bu soha hozirgi kunda jadal rivojlanmoqda. Ikkinchidan, uning amaliy ahamiyati nihoyatda katta.

{topic} ning asosiy xususiyatlari quyidagilardan iborat:
- Bu soha keng qamrovli bilimlarni o'z ichiga oladi
- Amaliy tatbiq etish imkoniyatlari kengdir  
- Zamonaviy texnologiyalar bilan uzviy bog'liqdir

Tadqiqotlar shuni ko'rsatadiki, {topic} bo'yicha bilimlar hayotning turli sohalarida qo'llanilishi mumkin. Bu esa uni o'rganishni yanada muhim qiladi.

XULOSA

{topic} mavzusini o'rganish natijasida shunday xulosaga kelish mumkin: bu soha kelajakda yanada rivojlanib, insoniyat hayotiga ijobiy ta'sir ko'rsatadi. Shu sababli, bu mavzuni chuqur o'rganish va tadqiq etish zarurdir.

Sana: {datetime.now().strftime('%d.%m.%Y')}
Uslub: {style}
"""


# ==================== FILE GENERATORS ====================

def create_pdf(title, content, style, doc_type):
    """Create PDF"""
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                           topMargin=2*cm, bottomMargin=2*cm,
                           leftMargin=2.5*cm, rightMargin=2.5*cm)
    
    story = []
    styles_obj = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'Title', fontSize=18, textColor=colors.HexColor('#1a5276'),
        spaceAfter=20, alignment=1, fontName='Helvetica-Bold'
    )
    
    heading_style = ParagraphStyle(
        'Heading', fontSize=13, textColor=colors.HexColor('#2874a6'),
        spaceAfter=10, spaceBefore=15, fontName='Helvetica-Bold'
    )
    
    body_style = ParagraphStyle(
        'Body', fontSize=11, spaceAfter=8, leading=16,
        alignment=4, fontName='Helvetica'
    )
    
    meta_style = ParagraphStyle(
        'Meta', fontSize=9, textColor=colors.HexColor('#888888'),
        spaceAfter=20, alignment=1
    )
    
    # Title
    story.append(Paragraph(title, title_style))
    story.append(Paragraph(
        f"Tur: {doc_type} | Uslub: {style} | Sana: {datetime.now().strftime('%d.%m.%Y')}",
        meta_style
    ))
    story.append(Spacer(1, 0.3*cm))
    
    # Content
    for line in content.split('\n'):
        line = line.strip()
        if not line:
            story.append(Spacer(1, 0.2*cm))
            continue
        
        if line.isupper() and len(line) < 50:
            story.append(Paragraph(line, heading_style))
        else:
            story.append(Paragraph(line, body_style))
    
    doc.build(story)
    buffer.seek(0)
    return buffer


def create_docx(title, content, style, doc_type):
    """Create DOCX"""
    doc = DocxDocument()
    
    # Title
    title_para = doc.add_heading(title, 0)
    title_para.alignment = 1
    
    # Metadata
    meta = doc.add_paragraph(f"Tur: {doc_type} | Uslub: {style} | Sana: {datetime.now().strftime('%d.%m.%Y')}")
    meta.alignment = 1
    meta.runs[0].font.size = Pt(9)
    meta.runs[0].font.color.rgb = None
    
    doc.add_paragraph()
    
    # Content
    for line in content.split('\n'):
        line = line.strip()
        if not line:
            doc.add_paragraph()
            continue
        
        if line.isupper() and len(line) < 50:
            doc.add_heading(line, 1)
        else:
            p = doc.add_paragraph(line)
            p.runs[0].font.size = Pt(11) if p.runs else None
    
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


def create_pptx(title, content, style, slide_count=8):
    """Create PPTX from AI content"""
    prs = Presentation()
    prs.slide_width = PptxInches(13.33)
    prs.slide_height = PptxInches(7.5)
    
    # Parse slides from content
    slides_data = []
    lines = content.split('\n')
    
    current_title = ""
    current_content = ""
    
    for line in lines:
        line = line.strip()
        if line.startswith("SLAYD_"):
            if current_title:
                slides_data.append((current_title, current_content))
                current_content = ""
            current_title = line.split(":", 1)[1].strip() if ":" in line else line
        elif line.startswith("MAZMUN_"):
            current_content = line.split(":", 1)[1].strip() if ":" in line else line
    
    if current_title:
        slides_data.append((current_title, current_content))
    
    # If no slides parsed, create from plain text
    if not slides_data:
        slides_data = [(title, content[:500])]
    
    # Color scheme
    bg_colors = [
        (26, 82, 118),   # Dark blue
        (23, 97, 86),    # Dark green
        (91, 44, 111),   # Dark purple
        (120, 40, 31),   # Dark red
        (30, 132, 73),   # Green
        (21, 67, 96),    # Navy
        (100, 30, 22),   # Maroon
        (17, 122, 101),  # Teal
    ]
    
    for i, (slide_title, slide_content) in enumerate(slides_data):
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank
        
        # Background color
        bg_color = bg_colors[i % len(bg_colors)]
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = PptxRGB(*bg_color)
        
        # Slide number (except first)
        if i > 0:
            num_box = slide.shapes.add_textbox(
                PptxInches(12.5), PptxInches(7.1),
                PptxInches(0.6), PptxInches(0.3)
            )
            num_tf = num_box.text_frame
            num_tf.text = str(i)
            num_tf.paragraphs[0].font.size = PptxPt(12)
            num_tf.paragraphs[0].font.color.rgb = PptxRGB(255, 255, 255)
            num_tf.paragraphs[0].font.bold = True
        
        if i == 0:
            # Title slide
            # Main title
            title_box = slide.shapes.add_textbox(
                PptxInches(1), PptxInches(2.5),
                PptxInches(11.33), PptxInches(2)
            )
            tf = title_box.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = slide_title
            p.font.size = PptxPt(44)
            p.font.bold = True
            p.font.color.rgb = PptxRGB(255, 255, 255)
            p.alignment = PP_ALIGN.CENTER
            
            # Subtitle
            sub_box = slide.shapes.add_textbox(
                PptxInches(1), PptxInches(5),
                PptxInches(11.33), PptxInches(1)
            )
            sub_tf = sub_box.text_frame
            sub_p = sub_tf.paragraphs[0]
            sub_p.text = f"{style} | {datetime.now().strftime('%d.%m.%Y')}"
            sub_p.font.size = PptxPt(18)
            sub_p.font.color.rgb = PptxRGB(200, 200, 200)
            sub_p.alignment = PP_ALIGN.CENTER
        else:
            # Content slide
            # Title bar
            title_bar = slide.shapes.add_textbox(
                PptxInches(0.5), PptxInches(0.3),
                PptxInches(12.33), PptxInches(1)
            )
            title_tf = title_bar.text_frame
            title_p = title_tf.paragraphs[0]
            title_p.text = slide_title
            title_p.font.size = PptxPt(28)
            title_p.font.bold = True
            title_p.font.color.rgb = PptxRGB(255, 255, 255)
            
            # Divider line (rectangle)
            line_shape = slide.shapes.add_shape(
                1,  # Rectangle
                PptxInches(0.5), PptxInches(1.4),
                PptxInches(12.33), Emu(50000)
            )
            line_shape.fill.solid()
            line_shape.fill.fore_color.rgb = PptxRGB(255, 255, 255)
            line_shape.line.fill.background()
            
            # Content
            content_box = slide.shapes.add_textbox(
                PptxInches(0.5), PptxInches(1.6),
                PptxInches(12.33), PptxInches(5.5)
            )
            content_tf = content_box.text_frame
            content_tf.word_wrap = True
            
            content_p = content_tf.paragraphs[0]
            content_p.text = slide_content
            content_p.font.size = PptxPt(20)
            content_p.font.color.rgb = PptxRGB(240, 240, 240)
    
    buffer = BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer


# ==================== BOT HANDLERS ====================

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user_name = update.effective_user.first_name
    
    if user_id not in user_data_store:
        user_data_store[user_id] = {"name": user_name, "documents": 0}
    
    keyboard = [
        [InlineKeyboardButton("📝 Referat", callback_data="doc_referat"),
         InlineKeyboardButton("📚 Kurs Ishi", callback_data="doc_kurs")],
        [InlineKeyboardButton("📄 Maqola", callback_data="doc_maqola"),
         InlineKeyboardButton("🎯 Slide", callback_data="doc_slide")],
        [InlineKeyboardButton("📊 Statistika", callback_data="stats"),
         InlineKeyboardButton("❓ Yordam", callback_data="help")]
    ]
    
    await update.message.reply_text(
        f"🎓 Assalomu alaykum, *{user_name}*!\n\n"
        "Siz faqat *mavzuni* yozing — qolganini bot o'zi qiladi:\n\n"
        "📝 Referat\n📚 Kurs Ishi\n📄 Maqola\n🎯 Slide (PPTX)\n\n"
        "💡 Birinchi hujjat *TEKIN!*",
        reply_markup=InlineKeyboardMarkup(keyboard),
        parse_mode=ParseMode.MARKDOWN
    )

async def handle_doc_type(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id
    doc_type = query.data.replace("doc_", "")
    
    if user_id not in user_data_store:
        user_data_store[user_id] = {}
    user_data_store[user_id]["doc_type"] = doc_type
    
    doc_names = {"referat": "📝 Referat", "kurs": "📚 Kurs Ishi",
                 "maqola": "📄 Maqola", "slide": "🎯 Slide"}
    
    keyboard = [
        [InlineKeyboardButton("🎯 APA", callback_data="style_apa"),
         InlineKeyboardButton("🎯 Harvard", callback_data="style_harvard")],
        [InlineKeyboardButton("🎯 O'zbek", callback_data="style_uzbek"),
         InlineKeyboardButton("🎯 Chicago", callback_data="style_chicago")]
    ]
    
    await query.edit_message_text(
        f"✅ *{doc_names.get(doc_type)}* tanlandi\n\n📋 Uslubni tanlang:",
        reply_markup=InlineKeyboardMarkup(keyboard),
        parse_mode=ParseMode.MARKDOWN
    )

async def handle_style(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id
    style = query.data.replace("style_", "")
    
    if user_id not in user_data_store:
        user_data_store[user_id] = {}
    user_data_store[user_id]["style"] = style
    
    doc_type = user_data_store[user_id].get("doc_type", "")
    
    # Slide uchun format tanlash
    if doc_type == "slide":
        keyboard = [
            [InlineKeyboardButton("5 ta slayd", callback_data="slides_5"),
             InlineKeyboardButton("8 ta slayd", callback_data="slides_8")],
            [InlineKeyboardButton("10 ta slayd", callback_data="slides_10"),
             InlineKeyboardButton("15 ta slayd", callback_data="slides_15")]
        ]
        await query.edit_message_text(
            "🎯 *Nechta slayd bo'lsin?*",
            reply_markup=InlineKeyboardMarkup(keyboard),
            parse_mode=ParseMode.MARKDOWN
        )
    else:
        keyboard = [
            [InlineKeyboardButton("📄 PDF", callback_data="format_pdf"),
             InlineKeyboardButton("📋 DOCX", callback_data="format_docx")]
        ]
        await query.edit_message_text(
            "📄 *Format tanlang:*",
            reply_markup=InlineKeyboardMarkup(keyboard),
            parse_mode=ParseMode.MARKDOWN
        )

async def handle_slide_count(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id
    count = int(query.data.replace("slides_", ""))
    
    if user_id not in user_data_store:
        user_data_store[user_id] = {}
    user_data_store[user_id]["slide_count"] = count
    user_data_store[user_id]["format"] = "pptx"
    
    await query.edit_message_text(
        f"✅ *{count} ta slayd* tanlandi\n\n"
        "✏️ *Mavzuni yozing:*\n\n"
        "_Misol: \"Sun'iy intellekt va ta'lim\"_",
        parse_mode=ParseMode.MARKDOWN
    )
    
    context.user_data[user_id] = {"state": "title"}

async def handle_format(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id
    file_format = query.data.replace("format_", "")
    
    if user_id not in user_data_store:
        user_data_store[user_id] = {}
    user_data_store[user_id]["format"] = file_format
    
    await query.edit_message_text(
        "✏️ *Mavzuni yozing:*\n\n"
        "_Misol: \"Kimyo texnologiyasi va IT\"_",
        parse_mode=ParseMode.MARKDOWN
    )
    
    context.user_data[user_id] = {"state": "title"}

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = update.message.text
    
    if user_id not in context.user_data:
        context.user_data[user_id] = {}
    
    state = context.user_data[user_id].get("state")
    
    if state == "title":
        user_data_store[user_id]["title"] = text
        context.user_data[user_id]["state"] = None
        
        # Start generating immediately!
        await update.message.reply_text(
            f"✅ Mavzu qabul qilindi: *{text}*\n\n"
            "⏳ Hujjat tayyorlanmoqda...\n"
            "_(30-60 sekund kuting)_",
            parse_mode=ParseMode.MARKDOWN
        )
        
        # Generate document automatically
        await auto_generate_document(update, context, user_id)
    else:
        # User sent message without selecting type
        keyboard = [
            [InlineKeyboardButton("📝 Referat", callback_data="doc_referat"),
             InlineKeyboardButton("📚 Kurs Ishi", callback_data="doc_kurs")],
            [InlineKeyboardButton("📄 Maqola", callback_data="doc_maqola"),
             InlineKeyboardButton("🎯 Slide", callback_data="doc_slide")]
        ]
        await update.message.reply_text(
            "📋 Avval hujjat turini tanlang:",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )

async def auto_generate_document(update, context, user_id):
    """Auto generate document with AI"""
    try:
        doc_data = user_data_store.get(user_id, {})
        title = doc_data.get("title", "Mavzu")
        style = doc_data.get("style", "uzbek").upper()
        doc_type = doc_data.get("doc_type", "referat")
        file_format = doc_data.get("format", "pdf")
        slide_count = doc_data.get("slide_count", 8)
        
        # Generate AI content
        content = await generate_ai_content(title, doc_type, style, slide_count)
        
        # Create file
        if file_format == "pdf":
            file_buffer = create_pdf(title, content, style, doc_type)
            filename = f"{title[:30].replace(' ', '_')}.pdf"
        elif file_format == "docx":
            file_buffer = create_docx(title, content, style, doc_type)
            filename = f"{title[:30].replace(' ', '_')}.docx"
        elif file_format == "pptx":
            file_buffer = create_pptx(title, content, style, slide_count)
            filename = f"{title[:30].replace(' ', '_')}.pptx"
        else:
            file_buffer = create_pdf(title, content, style, doc_type)
            filename = f"{title[:30].replace(' ', '_')}.pdf"
        
        # Send file
        await context.bot.send_document(
            chat_id=user_id,
            document=file_buffer,
            filename=filename
        )
        
        # Update stats
        user_data_store[user_id]["documents"] = user_data_store[user_id].get("documents", 0) + 1
        
        # Success message
        keyboard = [
            [InlineKeyboardButton("📝 Yangi Hujjat", callback_data="new_doc")],
            [InlineKeyboardButton("🏠 Bosh Menyu", callback_data="start")]
        ]
        
        await context.bot.send_message(
            chat_id=user_id,
            text=f"✅ *{filename}* tayyor!\n\n"
                 f"📝 Tur: {doc_type}\n"
                 f"📋 Uslub: {style}\n"
                 f"📄 Format: {file_format.upper()}",
            reply_markup=InlineKeyboardMarkup(keyboard),
            parse_mode=ParseMode.MARKDOWN
        )
        
    except Exception as e:
        await context.bot.send_message(
            chat_id=user_id,
            text=f"❌ Xato yuz berdi: {str(e)}\n\nQayta urinib ko'ring: /start"
        )

async def handle_stats(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id
    info = user_data_store.get(user_id, {})
    
    keyboard = [[InlineKeyboardButton("🏠 Bosh Menyu", callback_data="start")]]
    
    await query.edit_message_text(
        f"📊 *STATISTIKA*\n\n"
        f"👤 Ism: {info.get('name', 'N/A')}\n"
        f"📄 Hujjatlar: {info.get('documents', 0)}\n"
        f"🆔 ID: {user_id}",
        reply_markup=InlineKeyboardMarkup(keyboard),
        parse_mode=ParseMode.MARKDOWN
    )

async def handle_help(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    keyboard = [
        [InlineKeyboardButton("📝 Boshlash", callback_data="new_doc")],
        [InlineKeyboardButton("🏠 Bosh Menyu", callback_data="start")]
    ]
    
    await query.edit_message_text(
        "📖 *Qo'llanma*\n\n"
        "1️⃣ Hujjat turini tanlang\n"
        "2️⃣ Uslubni tanlang\n"
        "3️⃣ Slaydlar sonini tanlang (slide uchun)\n"
        "4️⃣ *Faqat mavzuni yozing!*\n"
        "5️⃣ Bot o'zi hujjat tayyorlaydi ✅\n\n"
        "⚡ 30-60 sekund ichida tayyor!",
        reply_markup=InlineKeyboardMarkup(keyboard),
        parse_mode=ParseMode.MARKDOWN
    )

async def new_doc(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    keyboard = [
        [InlineKeyboardButton("📝 Referat", callback_data="doc_referat"),
         InlineKeyboardButton("📚 Kurs Ishi", callback_data="doc_kurs")],
        [InlineKeyboardButton("📄 Maqola", callback_data="doc_maqola"),
         InlineKeyboardButton("🎯 Slide", callback_data="doc_slide")]
    ]
    
    await query.edit_message_text(
        "🎓 Hujjat turini tanlang:",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )

async def start_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    keyboard = [
        [InlineKeyboardButton("📝 Referat", callback_data="doc_referat"),
         InlineKeyboardButton("📚 Kurs Ishi", callback_data="doc_kurs")],
        [InlineKeyboardButton("📄 Maqola", callback_data="doc_maqola"),
         InlineKeyboardButton("🎯 Slide", callback_data="doc_slide")],
        [InlineKeyboardButton("📊 Statistika", callback_data="stats"),
         InlineKeyboardButton("❓ Yordam", callback_data="help")]
    ]
    
    await query.edit_message_text(
        "🎓 *Akademik Yordamchi Bot*\n\nHujjat turini tanlang:",
        reply_markup=InlineKeyboardMarkup(keyboard),
        parse_mode=ParseMode.MARKDOWN
    )

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "📖 *Qo'llanma*\n\n"
        "Faqat *mavzuni* yozing — bot o'zi hujjat yozadi!\n\n"
        "/start - Boshlash",
        parse_mode=ParseMode.MARKDOWN
    )

def main():
    app = Application.builder().token(BOT_TOKEN).build()
    
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))
    
    app.add_handler(CallbackQueryHandler(handle_doc_type, pattern="^doc_"))
    app.add_handler(CallbackQueryHandler(handle_style, pattern="^style_"))
    app.add_handler(CallbackQueryHandler(handle_slide_count, pattern="^slides_"))
    app.add_handler(CallbackQueryHandler(handle_format, pattern="^format_"))
    app.add_handler(CallbackQueryHandler(new_doc, pattern="^new_doc$"))
    app.add_handler(CallbackQueryHandler(handle_stats, pattern="^stats$"))
    app.add_handler(CallbackQueryHandler(handle_help, pattern="^help$"))
    app.add_handler(CallbackQueryHandler(start_callback, pattern="^start$"))
    
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    
    print("✅ BOT ISHGA TUSHDI!")
    print("🤖 AI CONTENT GENERATION ACTIVE!")
    print("📄 PDF + DOCX + PPTX READY!")
    
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()

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

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable
from reportlab.lib import colors
from reportlab.lib.units import cm

from docx import Document as DocxDocument
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

from pptx import Presentation
from pptx.util import Inches as Inch, Pt as PPt, Emu
from pptx.dml.color import RGBColor as PRGB
from pptx.enum.text import PP_ALIGN

load_dotenv()

BOT_TOKEN = os.getenv("BOT_TOKEN", "8625557628:AAHeUC2WxfMjJk-RRq3IxTtUJoc0H4XSsAM")
ADMIN_ID = int(os.getenv("ADMIN_ID", "7758296066"))
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")

user_data_store = {}

# ===================== AI CONTENT =====================

async def generate_with_gemini(prompt):
    if not GEMINI_API_KEY:
        return None
    try:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={GEMINI_API_KEY}"
        payload = {"contents": [{"parts": [{"text": prompt}]}]}
        async with aiohttp.ClientSession() as session:
            async with session.post(url, json=payload, timeout=aiohttp.ClientTimeout(total=45)) as resp:
                if resp.status == 200:
                    data = await resp.json()
                    return data["candidates"][0]["content"]["parts"][0]["text"]
    except:
        pass
    return None

async def generate_slide_content(topic, slide_count, style):
    prompt = f""""{topic}" mavzusida {slide_count} ta slayd uchun O'ZBEK TILIDA prezentatsiya yarat.

Qat'iy format:
===SLAYD_1===
SARLAVHA: {topic}
MAZMUN: Kirish so'zi - bu prezentatsiya haqida 2-3 jumlada

===SLAYD_2===
SARLAVHA: Mavzuga kirish
MAZMUN: {topic} nima ekanligini 3-4 jumlada tushuntir

===SLAYD_3===
SARLAVHA: Tarixiy background
MAZMUN: {topic} ning kelib chiqishi va tarixi haqida 3-4 jumla

===SLAYD_4===
SARLAVHA: Asosiy xususiyatlar
MAZMUN: {topic} ning 4-5 ta muhim xususiyatlarini sanab o't

===SLAYD_5===
SARLAVHA: Afzalliklari
MAZMUN: {topic} ning 4-5 ta asosiy afzalliklarini yoz

===SLAYD_6===
SARLAVHA: Kamchiliklari va muammolar
MAZMUN: {topic} sohasidagi 3-4 ta asosiy muammo va kamchiliklar

===SLAYD_7===
SARLAVHA: Dunyo tajribasi
MAZMUN: Dunyoda {topic} qanday qo'llanilayotgani haqida 3-4 jumla

===SLAYD_8===
SARLAVHA: O'zbekistonda holati
MAZMUN: O'zbekistonda {topic} ning rivojlanishi haqida 3-4 jumla

===SLAYD_9===
SARLAVHA: Kelajak istiqbollari
MAZMUN: {topic} ning kelajakdagi rivojlanishi haqida 3-4 jumla

===SLAYD_10===
SARLAVHA: Xulosa
MAZMUN: {topic} mavzusi bo'yicha 3-4 jumlali yakuniy xulosa

Uslub: {style}
FAQAT O'ZBEKCHA yoz. Har bir slayd uchun mazmunli va batafsil ma'lumot ber."""

    result = await generate_with_gemini(prompt)
    if result:
        return result
    return generate_fallback_slides(topic, slide_count)

def generate_fallback_slides(topic, slide_count):
    """Rich fallback content"""
    templates = [
        f"===SLAYD_1===\nSARLAVHA: {topic}\nMAZMUN: Ushbu prezentatsiya {topic} mavzusiga bag'ishlangan bo'lib, uning asosiy jihatlari, afzalliklari va kelajak istiqbollari ko'rib chiqiladi.",
        f"===SLAYD_2===\nSARLAVHA: {topic} haqida\nMAZMUN: {topic} - bu zamonaviy dunyoda muhim o'rin egallaydi. U turli sohalarda keng qo'llaniladi va insonlar hayotiga sezilarli ta'sir ko'rsatadi. Uni o'rganish va tushunish bugungi kunda zaruriy ehtiyojga aylangan.",
        f"===SLAYD_3===\nSARLAVHA: Rivojlanish tarixi\nMAZMUN: {topic} ning rivojlanish tarixi ko'p asrlarga borib taqaladi. Dastlabki bosqichda sodda ko'rinishda bo'lgan ushbu soha vaqt o'tishi bilan murakkab va mukammal tizimga aylandi. Har bir davr o'zining muhim hissasini qo'shdi.",
        f"===SLAYD_4===\nSARLAVHA: Asosiy xususiyatlar\nMAZMUN: {topic} ning asosiy xususiyatlari: 1) Yuqori samaradorlik va unumdorlik; 2) Keng qo'llanilish imkoniyatlari; 3) Doimiy rivojlanish va takomillashuv; 4) Iqtisodiy jihatdan foydaliligi; 5) Ekologik ta'siri va barqarorligi.",
        f"===SLAYD_5===\nSARLAVHA: Afzalliklari\nMAZMUN: {topic} ning afzalliklari: 1) Ish unumdorligini oshiradi; 2) Vaqt va resurslarni tejaydi; 3) Yangi imkoniyatlar yaratadi; 4) Xalqaro raqobatbardoshlikni ta'minlaydi; 5) Innovatsiyalar uchun zamin hozirlaydi.",
        f"===SLAYD_6===\nSARLAVHA: Muammolar va yechimlar\nMAZMUN: {topic} sohasida bir qator muammolar mavjud: malakali kadrlar yetishmasligi, moliyalashtirish muammolari va texnik cheklovlar. Bularni hal qilish uchun davlat dasturlari, xususiy investitsiyalar va xalqaro hamkorlik yo'lga qo'yilmoqda.",
        f"===SLAYD_7===\nSARLAVHA: Jahon tajribasi\nMAZMUN: Dunyoning yetakchi mamlakatlari {topic} sohasida katta yutuqlarga erishgan. Xususan, Germaniya, Yaponiya va AQSh bu sohada ilg'or tajribaga ega. Ularning tajribasini o'rganish va moslashtirib qo'llash muhim ahamiyat kasb etadi.",
        f"===SLAYD_8===\nSARLAVHA: O'zbekistonda rivojlanish\nMAZMUN: O'zbekistonda {topic} sohasida so'nggi yillarda sezilarli o'zgarishlar yuz berdi. 2022-2026 yillar rivojlanish strategiyasi doirasida bu sohaga alohida e'tibor qaratilmoqda. Yangi loyihalar va dasturlar amalga oshirilmoqda.",
        f"===SLAYD_9===\nSARLAVHA: Kelajak istiqbollari\nMAZMUN: {topic} sohasining kelajagi juda istiqbolli. Yangi texnologiyalar va innovatsiyalar bu sohani yanada rivojlantiradi. 2030 yilga kelib bu soha yanada kengayib, milliy iqtisodiyotga katta hissa qo'shishi kutilmoqda.",
        f"===SLAYD_10===\nSARLAVHA: Xulosa\nMAZMUN: {topic} mavzusini o'rganish natijasida ko'pgina muhim xulosalarga kelindi. Bu soha rivojlanishda davom etmoqda va kelajakda yanada muhim o'rin egallaydi. Barcha tomonlar hamkorligida bu sohani yanada rivojlantirish mumkin.",
    ]
    return "\n\n".join(templates[:slide_count])

async def generate_doc_content(topic, doc_type, style):
    doc_names = {"referat": "Referat", "kurs": "Kurs ishi", "maqola": "Ilmiy maqola"}
    
    prompt = f""""{topic}" mavzusida {doc_names.get(doc_type, 'hujjat')} yoz.
Uslub: {style}
Hajm: 600-900 so'z
Til: O'ZBEK TILI

Tuzilma:
KIRISH
[2-3 paragraf kirish]

ASOSIY QISM

1. {topic} ning mohiyati
[3-4 paragraf]

2. Asosiy muammolar va yechimlar
[3-4 paragraf]

3. Amaliy ahamiyati
[3-4 paragraf]

XULOSA
[2-3 paragraf xulosa]

ADABIYOTLAR RO'YXATI
1. ...
2. ...
3. ...

FAQAT O'ZBEKCHA yoz. Akademik uslubda yoz."""

    result = await generate_with_gemini(prompt)
    if result:
        return result
    return f"""KIRISH

{topic} mavzusi zamonaviy fan va amaliyotda muhim o'rin egallaydi. Ushbu {doc_names.get(doc_type, 'hujjat')}da biz bu mavzuning asosiy jihatlarini, uning ahamiyatini va kelajak istiqbollarini ko'rib chiqamiz.

Mavzuning dolzarbligi shundaki, bugungi kunda {topic} sohasida jadal o'zgarishlar yuz bermoqda. Bu o'zgarishlar nafaqat ilm-fan, balki amaliy hayotda ham o'z aksini topmoqda.

ASOSIY QISM

1. {topic} NING MOHIYATI

{topic} - bu ko'p qirrali va murakkab hodisa bo'lib, u bir necha asrlar davomida rivojlanib kelmoqda. Uning mohiyatini to'liq tushunish uchun tarixiy, nazariy va amaliy jihatlarini o'rganish zarur.

Tadqiqotlar shuni ko'rsatadiki, {topic} ning asosiy xususiyatlari uning keng qo'llanilish imkoniyatlarini belgilaydi. Turli sohalardagi mutaxassislar bu mavzuga o'ziga xos yondashadilar va har biri muayyan jihatlarini o'rganishga e'tibor qaratadi.

2. ASOSIY MUAMMOLAR VA YECHIMLAR

{topic} sohasida bir qator muammolar mavjud bo'lib, ularni hal qilish dolzarb masala hisoblanadi. Birinchi navbatda, kadrlar tayyorlash muammosi ko'zga tashlanadi.

Bundan tashqari, moliyaviy va texnik resurslar etishmovchiligi ham sezilarli to'siq bo'lib qolmoqda. Biroq, bu muammolarni hal qilishning aniq yo'llari mavjud va ular muvaffaqiyatli amalga oshirilmoqda.

3. AMALIY AHAMIYATI

{topic} ning amaliy ahamiyati nihoyatda katta. U iqtisodiy rivojlanishga, ijtimoiy farovonlikka va ilmiy taraqqiyotga bevosita hissa qo'shadi.

O'zbekistonda bu soha bo'yicha amalga oshirilayotgan islohotlar ijobiy natijalar bermoqda. Yangi loyihalar va dasturlar samarali ishlamoqda.

XULOSA

{topic} mavzusini o'rganish jarayonida quyidagi xulosalarga kelindi: bu soha rivojlanish sur'atlari jadallashmoqda, amaliy natijalari sezilarli bo'lyapti va kelajak istiqbollari juda yorqin.

Ushbu {doc_names.get(doc_type, 'hujjat')} natijalarini amaliyotga tadbiq etish va keyingi tadqiqotlar uchun poydevor bo'lishi mumkin.

ADABIYOTLAR RO'YXATI
1. O'zbekiston Respublikasi Prezidentining farmonlari to'plami, 2023.
2. Karimov I.A. O'zbekiston XXI asrga intilmoqda. T.: O'zbekiston, 2019.
3. Mirziyoyev Sh.M. Yangi O'zbekiston strategiyasi. T.: O'zbekiston, 2021.
4. {topic} bo'yicha xalqaro tadqiqotlar to'plami. 2022.
5. O'zbekiston milliy ensiklopediyasi. T.: 2020.

Sana: {datetime.now().strftime('%d.%m.%Y')}"""

# ===================== FILE CREATORS =====================

def parse_slides(content, topic, slide_count):
    slides = []
    parts = content.split("===SLAYD_")
    
    for part in parts[1:]:
        lines = part.strip().split("\n")
        title = ""
        text = ""
        for line in lines:
            if line.startswith("SARLAVHA:"):
                title = line.replace("SARLAVHA:", "").strip()
            elif line.startswith("MAZMUN:"):
                text = line.replace("MAZMUN:", "").strip()
        if title:
            slides.append((title, text))
    
    if not slides:
        slides = [(topic, content[:300])]
    
    return slides[:slide_count]

# 10 professional color themes
THEMES = [
    {"bg": (15, 32, 65), "accent": (52, 152, 219), "title": (255,255,255), "text": (220,230,240)},
    {"bg": (10, 54, 34), "accent": (39, 174, 96), "title": (255,255,255), "text": (210,240,220)},
    {"bg": (50, 10, 80), "accent": (155, 89, 182), "title": (255,255,255), "text": (230,210,245)},
    {"bg": (80, 20, 10), "accent": (231, 76, 60), "title": (255,255,255), "text": (245,215,210)},
    {"bg": (10, 55, 75), "accent": (26, 188, 156), "title": (255,255,255), "text": (200,240,235)},
    {"bg": (30, 30, 30), "accent": (241, 196, 15), "title": (255,255,255), "text": (240,230,180)},
    {"bg": (20, 40, 80), "accent": (52, 73, 94), "title": (255,215,0), "text": (220,220,220)},
    {"bg": (60, 10, 40), "accent": (192, 57, 43), "title": (255,255,255), "text": (245,200,215)},
    {"bg": (5, 45, 65), "accent": (41, 128, 185), "title": (255,215,0), "text": (200,235,255)},
    {"bg": (25, 55, 25), "accent": (46, 204, 113), "title": (255,255,255), "text": (205,245,220)},
]

def create_pptx(title, content, style, slide_count, theme_idx=0):
    prs = Presentation()
    prs.slide_width = Inch(13.33)
    prs.slide_height = Inch(7.5)
    
    slides_data = parse_slides(content, title, slide_count)
    theme = THEMES[theme_idx % len(THEMES)]
    
    for i, (slide_title, slide_text) in enumerate(slides_data):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # Background
        bg = slide.background.fill
        bg.solid()
        bg.fore_color.rgb = PRGB(*theme["bg"])
        
        # Accent bar (left side)
        bar = slide.shapes.add_shape(1, Inch(0), Inch(0), Inch(0.15), Inch(7.5))
        bar.fill.solid()
        bar.fill.fore_color.rgb = PRGB(*theme["accent"])
        bar.line.fill.background()
        
        if i == 0:
            # TITLE SLIDE
            # Bottom accent bar
            bot = slide.shapes.add_shape(1, Inch(0), Inch(6.8), Inch(13.33), Inch(0.7))
            bot.fill.solid()
            bot.fill.fore_color.rgb = PRGB(*theme["accent"])
            bot.line.fill.background()
            
            # Main title
            tb = slide.shapes.add_textbox(Inch(0.5), Inch(2), Inch(12.5), Inch(2.5))
            tf = tb.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = slide_title
            p.font.size = PPt(48)
            p.font.bold = True
            p.font.color.rgb = PRGB(*theme["title"])
            p.alignment = PP_ALIGN.CENTER
            
            # Subtitle
            sb = slide.shapes.add_textbox(Inch(0.5), Inch(5), Inch(12.5), Inch(1))
            sf = sb.text_frame
            sp = sf.paragraphs[0]
            sp.text = f"{style} uslubi  •  {datetime.now().strftime('%d.%m.%Y')}"
            sp.font.size = PPt(20)
            sp.font.color.rgb = PRGB(*theme["accent"])
            sp.alignment = PP_ALIGN.CENTER
            
        else:
            # CONTENT SLIDE
            # Top accent line
            top = slide.shapes.add_shape(1, Inch(0.15), Inch(1.3), Inch(13.18), Emu(40000))
            top.fill.solid()
            top.fill.fore_color.rgb = PRGB(*theme["accent"])
            top.line.fill.background()
            
            # Slide number circle bg
            num_bg = slide.shapes.add_shape(9, Inch(12.3), Inch(6.8), Inch(0.6), Inch(0.6))
            num_bg.fill.solid()
            num_bg.fill.fore_color.rgb = PRGB(*theme["accent"])
            num_bg.line.fill.background()
            
            # Slide number
            nb = slide.shapes.add_textbox(Inch(12.3), Inch(6.8), Inch(0.6), Inch(0.6))
            nf = nb.text_frame
            np_ = nf.paragraphs[0]
            np_.text = str(i)
            np_.font.size = PPt(14)
            np_.font.bold = True
            np_.font.color.rgb = PRGB(255, 255, 255)
            np_.alignment = PP_ALIGN.CENTER
            
            # Title
            ttb = slide.shapes.add_textbox(Inch(0.5), Inch(0.3), Inch(12.5), Inch(1))
            ttf = ttb.text_frame
            ttp = ttf.paragraphs[0]
            ttp.text = slide_title
            ttp.font.size = PPt(30)
            ttp.font.bold = True
            ttp.font.color.rgb = PRGB(*theme["title"])
            
            # Content
            ctb = slide.shapes.add_textbox(Inch(0.5), Inch(1.5), Inch(12.5), Inch(5.5))
            ctf = ctb.text_frame
            ctf.word_wrap = True
            
            # Split content into bullet points
            sentences = slide_text.replace(";", ".").split(".")
            first = True
            for sent in sentences:
                sent = sent.strip()
                if not sent or len(sent) < 5:
                    continue
                if first:
                    cp = ctf.paragraphs[0]
                    first = False
                else:
                    cp = ctf.add_paragraph()
                cp.text = "▸  " + sent
                cp.font.size = PPt(18)
                cp.font.color.rgb = PRGB(*theme["text"])
                cp.space_after = PPt(8)
    
    buffer = BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer

def create_pdf(title, content, style, doc_type):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                           topMargin=2*cm, bottomMargin=2*cm,
                           leftMargin=2.5*cm, rightMargin=2.5*cm)
    story = []
    base = getSampleStyleSheet()
    
    t_style = ParagraphStyle('T', fontSize=20, textColor=colors.HexColor('#1a5276'),
                             spaceAfter=6, alignment=1, fontName='Helvetica-Bold')
    m_style = ParagraphStyle('M', fontSize=9, textColor=colors.HexColor('#888888'),
                             spaceAfter=16, alignment=1)
    h_style = ParagraphStyle('H', fontSize=13, textColor=colors.HexColor('#2874a6'),
                             spaceAfter=8, spaceBefore=14, fontName='Helvetica-Bold')
    b_style = ParagraphStyle('B', fontSize=11, spaceAfter=8, leading=18,
                             alignment=4)
    
    story.append(Paragraph(title, t_style))
    story.append(Paragraph(f"{doc_type.upper()} | {style} uslubi | {datetime.now().strftime('%d.%m.%Y')}", m_style))
    story.append(HRFlowable(width="100%", thickness=2, color=colors.HexColor('#2874a6')))
    story.append(Spacer(1, 0.3*cm))
    
    for line in content.split('\n'):
        line = line.strip()
        if not line:
            story.append(Spacer(1, 0.15*cm))
        elif line.isupper() and len(line) < 60:
            story.append(Spacer(1, 0.2*cm))
            story.append(Paragraph(line, h_style))
        elif line[0].isdigit() and '. ' in line[:4]:
            story.append(Paragraph(f"<b>{line}</b>", h_style))
        else:
            story.append(Paragraph(line, b_style))
    
    doc.build(story)
    buffer.seek(0)
    return buffer

def create_docx(title, content, style, doc_type):
    doc = DocxDocument()
    
    # Title
    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run(title)
    r.font.size = Pt(20)
    r.font.bold = True
    r.font.color.rgb = RGBColor(26, 82, 118)
    
    # Meta
    m = doc.add_paragraph()
    m.alignment = WD_ALIGN_PARAGRAPH.CENTER
    mr = m.add_run(f"{doc_type.upper()} | {style} uslubi | {datetime.now().strftime('%d.%m.%Y')}")
    mr.font.size = Pt(9)
    mr.font.color.rgb = RGBColor(130, 130, 130)
    
    doc.add_paragraph()
    
    for line in content.split('\n'):
        line = line.strip()
        if not line:
            doc.add_paragraph()
            continue
        if line.isupper() and len(line) < 60:
            h = doc.add_paragraph()
            hr = h.add_run(line)
            hr.font.size = Pt(13)
            hr.font.bold = True
            hr.font.color.rgb = RGBColor(26, 82, 118)
        elif line[0].isdigit() and '. ' in line[:4]:
            h = doc.add_paragraph()
            hr = h.add_run(line)
            hr.font.size = Pt(12)
            hr.font.bold = True
        else:
            p = doc.add_paragraph(line)
            p.runs[0].font.size = Pt(11) if p.runs else None
    
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ===================== BOT HANDLERS =====================

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    name = update.effective_user.first_name
    if user_id not in user_data_store:
        user_data_store[user_id] = {"name": name, "documents": 0}
    
    kb = [
        [InlineKeyboardButton("📝 Referat", callback_data="doc_referat"),
         InlineKeyboardButton("📚 Kurs Ishi", callback_data="doc_kurs")],
        [InlineKeyboardButton("📄 Maqola", callback_data="doc_maqola"),
         InlineKeyboardButton("🎯 Slide", callback_data="doc_slide")],
        [InlineKeyboardButton("📊 Statistika", callback_data="stats"),
         InlineKeyboardButton("❓ Yordam", callback_data="help")]
    ]
    await update.message.reply_text(
        f"🎓 Assalomu alaykum, *{name}*!\n\n"
        "Faqat *mavzuni* yozing — bot o'zi yozadi:\n\n"
        "📝 Referat  •  📚 Kurs Ishi\n"
        "📄 Maqola  •  🎯 Slide (PPTX)\n\n"
        "💡 Birinchi hujjat *TEKIN!*",
        reply_markup=InlineKeyboardMarkup(kb),
        parse_mode=ParseMode.MARKDOWN
    )

async def cb_doc(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    uid = q.from_user.id
    dtype = q.data.replace("doc_", "")
    if uid not in user_data_store:
        user_data_store[uid] = {}
    user_data_store[uid]["doc_type"] = dtype
    
    names = {"referat":"📝 Referat","kurs":"📚 Kurs Ishi","maqola":"📄 Maqola","slide":"🎯 Slide"}
    kb = [
        [InlineKeyboardButton("APA", callback_data="style_APA"),
         InlineKeyboardButton("Harvard", callback_data="style_Harvard")],
        [InlineKeyboardButton("O'zbek", callback_data="style_Uzbek"),
         InlineKeyboardButton("Chicago", callback_data="style_Chicago")]
    ]
    await q.edit_message_text(f"✅ *{names[dtype]}* tanlandi\n\n📋 Uslubni tanlang:",
                              reply_markup=InlineKeyboardMarkup(kb), parse_mode=ParseMode.MARKDOWN)

async def cb_style(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    uid = q.from_user.id
    style = q.data.replace("style_", "")
    if uid not in user_data_store:
        user_data_store[uid] = {}
    user_data_store[uid]["style"] = style
    dtype = user_data_store[uid].get("doc_type", "")
    
    if dtype == "slide":
        kb = [
            [InlineKeyboardButton("5 ta", callback_data="slides_5"),
             InlineKeyboardButton("8 ta", callback_data="slides_8")],
            [InlineKeyboardButton("10 ta", callback_data="slides_10"),
             InlineKeyboardButton("15 ta", callback_data="slides_15")]
        ]
        await q.edit_message_text("🎯 *Nechta slayd bo'lsin?*",
                                  reply_markup=InlineKeyboardMarkup(kb), parse_mode=ParseMode.MARKDOWN)
    else:
        kb = [
            [InlineKeyboardButton("📄 PDF", callback_data="fmt_pdf"),
             InlineKeyboardButton("📋 DOCX", callback_data="fmt_docx")]
        ]
        await q.edit_message_text("📄 *Format tanlang:*",
                                  reply_markup=InlineKeyboardMarkup(kb), parse_mode=ParseMode.MARKDOWN)

async def cb_slides(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    uid = q.from_user.id
    count = int(q.data.replace("slides_", ""))
    if uid not in user_data_store:
        user_data_store[uid] = {}
    user_data_store[uid]["slide_count"] = count
    user_data_store[uid]["format"] = "pptx"
    
    # Theme selection
    kb = [
        [InlineKeyboardButton("🔵 Ko'k", callback_data="theme_0"),
         InlineKeyboardButton("🟢 Yashil", callback_data="theme_1")],
        [InlineKeyboardButton("🟣 Binafsha", callback_data="theme_2"),
         InlineKeyboardButton("🔴 Qizil", callback_data="theme_3")],
        [InlineKeyboardButton("🩵 Moviy", callback_data="theme_4"),
         InlineKeyboardButton("🟡 Sariq", callback_data="theme_5")]
    ]
    await q.edit_message_text(f"✅ *{count} ta slayd* tanlandi\n\n🎨 *Dizayn rangini tanlang:*",
                              reply_markup=InlineKeyboardMarkup(kb), parse_mode=ParseMode.MARKDOWN)

async def cb_theme(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    uid = q.from_user.id
    theme_idx = int(q.data.replace("theme_", ""))
    if uid not in user_data_store:
        user_data_store[uid] = {}
    user_data_store[uid]["theme"] = theme_idx
    
    await q.edit_message_text(
        "✏️ *Mavzuni yozing:*\n\n"
        "_Misol: \"Sun'iy intellekt va ta'lim\"_",
        parse_mode=ParseMode.MARKDOWN
    )
    context.user_data[uid] = {"state": "title"}

async def cb_fmt(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    uid = q.from_user.id
    fmt = q.data.replace("fmt_", "")
    if uid not in user_data_store:
        user_data_store[uid] = {}
    user_data_store[uid]["format"] = fmt
    
    await q.edit_message_text(
        "✏️ *Mavzuni yozing:*\n\n_Misol: \"Kimyo va IT\"_",
        parse_mode=ParseMode.MARKDOWN
    )
    context.user_data[uid] = {"state": "title"}

async def on_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    text = update.message.text
    
    if uid not in context.user_data:
        context.user_data[uid] = {}
    
    if context.user_data[uid].get("state") == "title":
        user_data_store[uid]["title"] = text
        context.user_data[uid]["state"] = None
        
        await update.message.reply_text(
            f"✅ Mavzu: *{text}*\n\n⏳ Hujjat tayyorlanmoqda...\n_(30-60 sekund)_",
            parse_mode=ParseMode.MARKDOWN
        )
        await do_generate(update, context, uid)
    else:
        kb = [
            [InlineKeyboardButton("📝 Referat", callback_data="doc_referat"),
             InlineKeyboardButton("🎯 Slide", callback_data="doc_slide")]
        ]
        await update.message.reply_text("📋 Avval hujjat turini tanlang:",
                                        reply_markup=InlineKeyboardMarkup(kb))

async def do_generate(update, context, uid):
    try:
        d = user_data_store.get(uid, {})
        title = d.get("title", "Mavzu")
        style = d.get("style", "Uzbek")
        dtype = d.get("doc_type", "referat")
        fmt = d.get("format", "pdf")
        scount = d.get("slide_count", 10)
        theme = d.get("theme", 0)
        
        if fmt == "pptx":
            content = await generate_slide_content(title, scount, style)
            buf = create_pptx(title, content, style, scount, theme)
            fname = f"{title[:25].replace(' ','_')}.pptx"
        elif fmt == "docx":
            content = await generate_doc_content(title, dtype, style)
            buf = create_docx(title, content, style, dtype)
            fname = f"{title[:25].replace(' ','_')}.docx"
        else:
            content = await generate_doc_content(title, dtype, style)
            buf = create_pdf(title, content, style, dtype)
            fname = f"{title[:25].replace(' ','_')}.pdf"
        
        await context.bot.send_document(chat_id=uid, document=buf, filename=fname)
        
        user_data_store[uid]["documents"] = user_data_store[uid].get("documents", 0) + 1
        
        kb = [
            [InlineKeyboardButton("📝 Yangi Hujjat", callback_data="new_doc")],
            [InlineKeyboardButton("🏠 Bosh Menyu", callback_data="start")]
        ]
        await context.bot.send_message(
            chat_id=uid,
            text=f"✅ *{fname}* yuborildi!\n\n"
                 f"📝 Tur: {dtype}\n📋 Uslub: {style}\n📄 Format: {fmt.upper()}",
            reply_markup=InlineKeyboardMarkup(kb),
            parse_mode=ParseMode.MARKDOWN
        )
    except Exception as e:
        await context.bot.send_message(chat_id=uid, text=f"❌ Xato: {str(e)}\n\n/start dan qayta bosing")

async def cb_stats(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    uid = q.from_user.id
    info = user_data_store.get(uid, {})
    kb = [[InlineKeyboardButton("🏠 Bosh Menyu", callback_data="start")]]
    await q.edit_message_text(
        f"📊 *Statistika*\n\n👤 {info.get('name','N/A')}\n"
        f"📄 Hujjatlar: {info.get('documents',0)}\n🆔 {uid}",
        reply_markup=InlineKeyboardMarkup(kb), parse_mode=ParseMode.MARKDOWN
    )

async def cb_help(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    kb = [[InlineKeyboardButton("📝 Boshlash", callback_data="new_doc")]]
    await q.edit_message_text(
        "📖 *Qo'llanma*\n\n"
        "1️⃣ Hujjat turini tanlang\n"
        "2️⃣ Uslubni tanlang\n"
        "3️⃣ Slayd soni va rangni tanlang\n"
        "4️⃣ Faqat *mavzuni* yozing\n"
        "5️⃣ Bot o'zi yozib beradi! ✅",
        reply_markup=InlineKeyboardMarkup(kb), parse_mode=ParseMode.MARKDOWN
    )

async def cb_new(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    kb = [
        [InlineKeyboardButton("📝 Referat", callback_data="doc_referat"),
         InlineKeyboardButton("📚 Kurs Ishi", callback_data="doc_kurs")],
        [InlineKeyboardButton("📄 Maqola", callback_data="doc_maqola"),
         InlineKeyboardButton("🎯 Slide", callback_data="doc_slide")]
    ]
    await q.edit_message_text("🎓 Hujjat turini tanlang:", reply_markup=InlineKeyboardMarkup(kb))

async def cb_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    kb = [
        [InlineKeyboardButton("📝 Referat", callback_data="doc_referat"),
         InlineKeyboardButton("📚 Kurs Ishi", callback_data="doc_kurs")],
        [InlineKeyboardButton("📄 Maqola", callback_data="doc_maqola"),
         InlineKeyboardButton("🎯 Slide", callback_data="doc_slide")],
        [InlineKeyboardButton("📊 Statistika", callback_data="stats"),
         InlineKeyboardButton("❓ Yordam", callback_data="help")]
    ]
    await q.edit_message_text("🎓 *Akademik Yordamchi Bot*\n\nHujjat turini tanlang:",
                              reply_markup=InlineKeyboardMarkup(kb), parse_mode=ParseMode.MARKDOWN)

def main():
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", lambda u,c: u.message.reply_text("Boshlash uchun /start bosing")))
    app.add_handler(CallbackQueryHandler(cb_doc, pattern="^doc_"))
    app.add_handler(CallbackQueryHandler(cb_style, pattern="^style_"))
    app.add_handler(CallbackQueryHandler(cb_slides, pattern="^slides_"))
    app.add_handler(CallbackQueryHandler(cb_theme, pattern="^theme_"))
    app.add_handler(CallbackQueryHandler(cb_fmt, pattern="^fmt_"))
    app.add_handler(CallbackQueryHandler(cb_stats, pattern="^stats$"))
    app.add_handler(CallbackQueryHandler(cb_help, pattern="^help$"))
    app.add_handler(CallbackQueryHandler(cb_new, pattern="^new_doc$"))
    app.add_handler(CallbackQueryHandler(cb_start, pattern="^start$"))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, on_message))
    
    print("✅ BOT ISHGA TUSHDI!")
    print("🎨 10 TA DIZAYN MAVJUD!")
    print("🤖 AI CONTENT ACTIVE!")
    
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()

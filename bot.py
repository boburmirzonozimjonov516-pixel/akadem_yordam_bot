import os
import asyncio
import aiohttp
import json
import random
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

# ===================== UNSPLASH IMAGES =====================

async def get_topic_image(topic_en):
    """Get real image from Unsplash (free, no key needed)"""
    try:
        queries = [topic_en, "education", "technology", "science", "business"]
        for q in queries:
            url = f"https://source.unsplash.com/1280x720/?{q.replace(' ', ',')}"
            async with aiohttp.ClientSession() as session:
                async with session.get(url, timeout=aiohttp.ClientTimeout(total=10), allow_redirects=True) as resp:
                    if resp.status == 200 and 'image' in resp.headers.get('content-type', ''):
                        return BytesIO(await resp.read())
    except:
        pass
    return None

async def translate_topic(topic):
    """Simple translation for image search"""
    translations = {
        "sun'iy intellekt": "artificial intelligence",
        "texnologiya": "technology",
        "ta'lim": "education",
        "iqtisodiyot": "economy",
        "tibbiyot": "medicine",
        "sport": "sport",
        "kimyo": "chemistry",
        "fizika": "physics",
        "biologiya": "biology",
        "matematik": "mathematics",
        "tarix": "history",
        "siyosat": "politics",
        "ekologiya": "ecology",
        "arxitektura": "architecture",
        "san'at": "art",
        "musiqa": "music",
        "adabiyot": "literature",
        "energetika": "energy",
        "qishloq": "agriculture",
        "transport": "transport",
        "kosmос": "space",
        "internet": "internet",
        "dasturlash": "programming",
        "biznes": "business",
        "moliya": "finance",
        "bank": "banking",
    }
    topic_lower = topic.lower()
    for uz, en in translations.items():
        if uz in topic_lower:
            return en
    return topic

# ===================== AI CONTENT =====================

async def generate_with_gemini(prompt):
    if not GEMINI_API_KEY:
        return None
    try:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={GEMINI_API_KEY}"
        payload = {
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {"temperature": 0.7, "maxOutputTokens": 2048}
        }
        async with aiohttp.ClientSession() as session:
            async with session.post(url, json=payload, timeout=aiohttp.ClientTimeout(total=50)) as resp:
                if resp.status == 200:
                    data = await resp.json()
                    return data["candidates"][0]["content"]["parts"][0]["text"]
    except:
        pass
    return None

async def generate_slide_content(topic, slide_count, style):
    prompt = f"""Sen O'zbekiston universitetining eng yaxshi o'qituvchisisan.
"{topic}" mavzusida {slide_count} ta PROFESSIONAL SLAYD uchun BATAFSIL mazmun yarat.

QOIDA: Har bir slayd uchun 4-6 ta muhim nuqta yoz. Har bir nuqta 2-3 jumladan iborat bo'lsin.
Uslub: {style}
TIL: Faqat O'zbekcha

FORMAT (aniq shunday yoz):
===SLAYD_1===
SARLAVHA: {topic}
MAZMUN: Bu prezentatsiya {topic} mavzusini chuqur tahlil qiladi. Mavzuning barcha muhim jihatlari ko'rib chiqiladi. Tinglovchilar yangi bilim va ko'nikmalar egallaydi.

===SLAYD_2===
SARLAVHA: Mavzuga umumiy kirish
MAZMUN: {topic} bugungi kunda juda dolzarb mavzu hisoblanadi. Bu soha tez sur'atlar bilan rivojlanmoqda va yangi imkoniyatlar yaratmoqda. Uni o'rganish har bir mutaxassis uchun zaruriy ehtiyojga aylandi.

===SLAYD_3===
SARLAVHA: Tarixiy rivojlanish
MAZMUN: {topic} ning tarixi bir necha asrlar oldin boshlangan. Dastlabki bosqichlarda oddiy shaklda bo'lgan bu soha, fan va texnikaning rivojlanishi bilan murakkab tizimga aylandi. Muhim tarixiy voqealar bu sohaning bugungi qiyofasini belgiladi.

===SLAYD_4===
SARLAVHA: Asosiy tushunchalar va ta'riflar
MAZMUN: {topic} sohasida bir qator asosiy tushunchalar mavjud. Ularni to'g'ri tushunish sohani chuqur o'rganish uchun poydevor yaratadi. Mutaxassislar bu tushunchalarni keng qo'llaydilar.

===SLAYD_5===
SARLAVHA: Afzalliklari va imkoniyatlari
MAZMUN: {topic} ning asosiy afzalliklari: samaradorlikni oshirish, resurslarni tejash, yangi imkoniyatlar yaratish va raqobatbardoshlikni ta'minlash. Bu afzalliklar uni boshqa sohalardan ajratib turadi.

===SLAYD_6===
SARLAVHA: Muammolar va yechimlar
MAZMUN: Har qanday sohada bo'lgani kabi, {topic} da ham muammolar mavjud. Asosiy to'siqlar: moliyalashtirish, kadrlar va texnik cheklovlar. Biroq, zamonaviy yondashuv va innovatsiyalar orqali bu muammolar hal etilmoqda.

===SLAYD_7===
SARLAVHA: Jahon tajribasi
MAZMUN: Dunyoning rivojlangan mamlakatlari {topic} sohasida ulkan yutuqlarga erishgan. AQSH, Germaniya, Yaponiya va Xitoy bu sohada yetakchi o'rinlarda turadi. Ularning tajribasi boshqa mamlakatlar uchun muhim ibrat manbai.

===SLAYD_8===
SARLAVHA: O'zbekistonda rivojlanish
MAZMUN: O'zbekistonda {topic} sohasida jadal rivojlanish kuzatilmoqda. Prezident Mirziyoyev rahbarligida qabul qilingan dasturlar bu sohaga yangi turtki berdi. Yangi loyihalar va sarmoyalar natijasida soha yanada o'smoqda.

===SLAYD_9===
SARLAVHA: Kelajak istiqbollari
MAZMUN: {topic} sohasining kelajagi juda porloq. Yangi texnologiyalar joriy etilishi bilan soha yanada rivojlanadi. 2030 yilga kelib bu soha milliy iqtisodiyotda muhim o'rin egallaydi deb kutilmoqda.

===SLAYD_10===
SARLAVHA: Xulosa va tavsiyalar
MAZMUN: {topic} mavzusini o'rganish juda muhim va zarur ekan. Ushbu prezentatsiya asosiy jihatlarni qamrab oldi. Barcha soha vakillari, tadqiqotchilar va talabalar bu bilimlarni amaliyotda qo'llashlari tavsiya etiladi.

Agar {slide_count} > 10 bo'lsa, qo'shimcha slaydlar ham shu formatda qo'sh."""

    result = await generate_with_gemini(prompt)
    if result and "===SLAYD_" in result:
        return result
    return generate_rich_fallback(topic, slide_count)

def generate_rich_fallback(topic, slide_count):
    slides = [
        f"===SLAYD_1===\nSARLAVHA: {topic}\nMAZMUN: Ushbu prezentatsiya {topic} mavzusining barcha muhim jihatlarini qamrab oladi. Zamonaviy ilm-fan va amaliyot nuqtai nazaridan tahlil qilinadi. Tinglovchilar uchun foydali va qiziqarli ma'lumotlar taqdim etiladi.",
        f"===SLAYD_2===\nSARLAVHA: {topic} nima?\nMAZMUN: {topic} - bu zamonaviy dunyoda alohida o'rin egallagan muhim soha. Uning ta'rifi bir necha nuqtai nazardan berilishi mumkin. Turli olimlar va mutaxassislar bu tushunchaga o'ziga xos yondashuv bilan qaraydilar.",
        f"===SLAYD_3===\nSARLAVHA: Rivojlanish bosqichlari\nMAZMUN: {topic} ning rivojlanish tarixi qiziqarli va murakkab. Birinchi bosqich - dastlabki kashfiyotlar va nazariyalar. Ikkinchi bosqich - amaliyotga joriy etish. Uchinchi bosqich - zamonaviy rivojlanish va innovatsiyalar.",
        f"===SLAYD_4===\nSARLAVHA: Asosiy xususiyatlar\nMAZMUN: {topic} ning 5 ta asosiy xususiyati: 1) Yuqori samaradorlik; 2) Keng qo'llanilish doirasi; 3) Doimiy takomillashuv; 4) Iqtisodiy foydaliligi; 5) Ijtimoiy ahamiyati. Bu xususiyatlar uni boshqa sohalardan farqlaydi.",
        f"===SLAYD_5===\nSARLAVHA: Afzalliklari\nMAZMUN: {topic} ning asosiy afzalliklari: ish unumdorligini 40-60% ga oshiradi. Vaqt va moliyaviy resurslarni tejaydi. Yangi ish o'rinlari yaratadi. Xalqaro raqobatbardoshlikni ta'minlaydi. Innovatsion rivojlanishga zamin hozirlaydi.",
        f"===SLAYD_6===\nSARLAVHA: Muammolar va yechimlar\nMAZMUN: Asosiy muammolar: malakali kadrlar tanqisligi (yechim: ta'lim tizimini isloh qilish), moliyalashtirish yetishmasligi (yechim: xususiy-davlat hamkorligi), texnik cheklovlar (yechim: zamonaviy infratuzilma). Har bir muammoning aniq yechimi mavjud.",
        f"===SLAYD_7===\nSARLAVHA: Statistika va raqamlar\nMAZMUN: Jahon bo'yicha {topic} sohasiga har yili 500+ mlrd dollar sarmoya kiritiladi. So'nggi 5 yilda soha 35% ga o'sdi. 2025 yilga kelib 10 mln yangi ish o'rini yaratiladi. O'zbekistonda so'nggi yil 25% o'sish qayd etildi.",
        f"===SLAYD_8===\nSARLAVHA: O'zbekistonda holat\nMAZMUN: O'zbekistonda {topic} rivojlantirishga 2 trln so'm ajratilgan. 15 ta yangi loyiha amalga oshirilmoqda. 5000 dan ortiq mutaxassis tayyorlandi. Xalqaro hamkorlar: Koreya, Germaniya, Xitoy. 2026 yilga qadar yanada kengayishi rejalashtirilgan.",
        f"===SLAYD_9===\nSARLAVHA: Kelajak\nMAZMUN: {topic} ning kelajagi: 2025-2030 yillarda 3 barobar o'sish kutilmoqda. Sun'iy intellekt va raqamlashtirish yangi imkoniyatlar ochadi. Yosh mutaxassislar uchun katta istiqbollar. Innovatsion startaplar ko'payadi. Global bozorda O'zbekistonning ulushi oshadi.",
        f"===SLAYD_10===\nSARLAVHA: Xulosa\nMAZMUN: {topic} - kelajak sohasi. Asosiy xulosalar: 1) Soha jadal rivojlanmoqda; 2) O'zbekiston yaxshi imkoniyatlarga ega; 3) Yoshlar uchun istiqbolli karera; 4) Davlat tomonidan qo'llab-quvvatlanmoqda. Birgalikda kelajakni quramiz!",
    ]

    extra_slides = [
        f"===SLAYD_11===\nSARLAVHA: Amaliy misol\nMAZMUN: {topic} ning real hayotdagi misollari: Toshkent shahrida muvaffaqiyatli loyiha amalga oshirildi. Natijada iqtisodiy samara 150% ga oshdi. Xalqaro ekspertlar yuqori baho berdi. Bu tajriba boshqa viloyatlarda ham qo'llanilmoqda.",
        f"===SLAYD_12===\nSARLAVHA: Tavsiyalar\nMAZMUN: Mutaxassislar uchun: doimo yangiliklarga e'tibor bering. Yoshlar uchun: bu sohada tahsil oling. Tadbirkorlar uchun: investitsiya qiling. Davlat uchun: qulayroq sharoit yarating. Hamkorlik va innovatsiya - muvaffaqiyat kaliti.",
        f"===SLAYD_13===\nSARLAVHA: Xalqaro hamkorlik\nMAZMUN: {topic} sohasida xalqaro hamkorlik juda muhim. O'zbekiston 25+ davlat bilan hamkorlik shartnomasi imzolagan. Xalqaro tashkilotlar: UN, WB, ADB qo'llab-quvvatlaydi. Birgalikda amalga oshirilayotgan loyihalar soni oshmoqda.",
        f"===SLAYD_14===\nSARLAVHA: Yoshlar va ta'lim\nMAZMUN: {topic} sohasida yoshlar uchun katta imkoniyatlar mavjud. 50+ universitet bu yo'nalishda mutaxassislar tayyorlaydi. Stipendiya va grantlar berilmoqda. Stajировка va amaliyot dasturlari kengaytirilmoqda. Yoshlar - kelajak sohasi yetakchilari.",
        f"===SLAYD_15===\nSARLAVHA: E'tiboringiz uchun rahmat!\nMAZMUN: Savollar va muhokama uchun vaqt. Ushbu prezentatsiya {topic} mavzusini to'liq qamrab oldi deb umid qilamiz. Qo'shimcha ma'lumot uchun murojaat qilishingiz mumkin. Birgalikda O'zbekistonni rivojlantiramiz!",
    ]

    all_slides = slides + extra_slides
    return "\n\n".join(all_slides[:slide_count])

async def generate_doc_content(topic, doc_type, style):
    doc_names = {"referat": "Referat", "kurs": "Kurs ishi", "maqola": "Ilmiy maqola"}
    prompt = f"""O'zbek tili va adabiyoti bo'yicha professor darajasida "{topic}" mavzusida {doc_names.get(doc_type, 'hujjat')} yoz.
Uslub: {style} | Hajm: 800-1200 so'z | TIL: Faqat O'zbekcha

Tuzilma:
KIRISH
ASOSIY QISM
1. {topic} ning nazariy asoslari
2. Amaliy jihatlar
3. Muammolar va yechimlar
4. Zamonaviy tendensiyalar
XULOSA
ADABIYOTLAR (5 ta)"""

    result = await generate_with_gemini(prompt)
    if result:
        return result
    return f"""KIRISH

{topic} - zamonaviy fan va amaliyotda muhim o'rin egallagan dolzarb mavzu. Ushbu {doc_names.get(doc_type, 'hujjat')} da mavzuning asosiy jihatlari ilmiy asosda tahlil qilinadi.

Mavzuning dolzarbligi shundaki, bugungi globallashuv sharoitida {topic} sohasida bilim va ko'nikmaga ega bo'lish har bir mutaxassis uchun zaruriy talabga aylanib bormoqda.

ASOSIY QISM

1. {topic.upper()} NING NAZARIY ASOSLARI

{topic} tushunchasi ilmiy adabiyotlarda turlicha ta'riflanadi. Ko'pchilik olimlar uni murakkab tizim sifatida baholaydi. Bu sohaning nazariy poydevori bir necha asrlar davomida shakllanib kelgan.

Asosiy nazariy yondashuvlar: tizimli tahlil, tarixiy taqqoslash va empirik tadqiqot metodlari. Har biri o'ziga xos afzallik va kamchiliklarga ega.

2. AMALIY JIHATLAR

{topic} ning amaliyotdagi qo'llanilishi juda keng. Sanoat, ta'lim, tibbiyot, transport va boshqa ko'plab sohalarda faol foydalanilmoqda.

Amaliy samaradorlik ko'rsatkichlari: unumdorlik 35-50% ga oshadi, xarajatlar 20-30% ga kamayadi, sifat ko'rsatkichlari sezilarli yaxshilanadi.

3. MUAMMOLAR VA YECHIMLAR

Asosiy muammolar:
• Malakali kadrlar yetishmasligi — yechim: maqsadli ta'lim dasturlari
• Moliyaviy resurslar tanqisligi — yechim: xususiy-davlat sherikchiligi
• Texnik infratuzilma kamchiligi — yechim: xalqaro investitsiyalar

4. ZAMONAVIY TENDENSIYALAR

Jahon miqyosida {topic} sohasida bir necha asosiy tendensiyalar kuzatilmoqda:
• Raqamlashtirish va avtomatlashtirish
• Sun'iy intellektni qo'llash
• Barqarorli rivojlanish tamoyillari
• Xalqaro integratsiya va hamkorlik

O'zbekistonda ushbu tendensiyalar "Yangi O'zbekiston" strategiyasi doirasida muvaffaqiyatli amalga oshirilmoqda.

XULOSA

{topic} mavzusini chuqur o'rganish quyidagi muhim xulosalarga olib keldi:

1. Soha jadal rivojlanmoqda va kelajakda yanada muhim o'rin egallaydi
2. O'zbekistonda bu soha uchun qulay sharoit va imkoniyatlar mavjud
3. Kadrlar tayyorlash va innovatsiyalarga e'tibor berish kerak
4. Xalqaro hamkorlik va tajriba almashinuvi zarur

ADABIYOTLAR RO'YXATI

1. Mirziyoyev Sh.M. Yangi O'zbekiston strategiyasi. — T.: O'zbekiston, 2021. — 280 b.
2. Karimov I.A. O'zbekiston XXI asrga intilmoqda. — T.: O'zbekiston, 2019. — 320 b.
3. {topic} bo'yicha xalqaro tadqiqotlar to'plami. — M.: Nauka, 2022. — 450 b.
4. O'zbekiston Respublikasi Prezidentining PQ-4861 son qarori. — T., 2023.
5. UNESCO. Global Education Report 2023. — Paris: UNESCO, 2023. — 200 p.

Sana: {datetime.now().strftime('%d.%m.%Y')} | Uslub: {style}"""

# ===================== PPTX CREATOR =====================

def parse_slides(content, slide_count):
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
    return slides[:slide_count] if slides else [("Mavzu", content[:300])]

# 10 PROFESSIONAL THEMES
THEMES = [
    {"name":"🔵 Ko'k Premium", "bg":(15,32,65), "accent":(52,152,219), "title":(255,255,255), "text":(210,230,255), "sub":(100,180,255)},
    {"name":"🟢 Yashil Natura", "bg":(10,45,30), "accent":(39,174,96), "title":(255,255,255), "text":(200,240,215), "sub":(100,220,150)},
    {"name":"🟣 Binafsha Royal", "bg":(40,10,70), "accent":(155,89,182), "title":(255,255,255), "text":(230,210,250), "sub":(180,130,230)},
    {"name":"🔴 Qizil Dinamik", "bg":(70,10,10), "accent":(231,76,60), "title":(255,255,255), "text":(255,210,205), "sub":(255,140,130)},
    {"name":"🩵 Moviy Ocean", "bg":(5,40,60), "accent":(26,188,156), "title":(255,255,255), "text":(200,245,240), "sub":(100,230,210)},
    {"name":"🟡 Oltin Biznes", "bg":(30,25,5), "accent":(241,196,15), "title":(255,255,255), "text":(255,240,180), "sub":(255,215,80)},
    {"name":"⚫ Qora Elegant", "bg":(15,15,15), "accent":(200,200,200), "title":(255,215,0), "text":(220,220,220), "sub":(180,180,180)},
    {"name":"🟤 Jigarrang", "bg":(50,25,10), "accent":(211,84,0), "title":(255,255,255), "text":(255,220,190), "sub":(255,160,100)},
    {"name":"🌊 Dengiz", "bg":(5,30,60), "accent":(41,128,185), "title":(255,215,0), "text":(200,235,255), "sub":(130,200,255)},
    {"name":"🌿 Nefrit", "bg":(10,40,30), "accent":(22,160,133), "title":(255,255,255), "text":(200,240,230), "sub":(80,200,170)},
]

async def create_pptx_with_images(title, content, style, slide_count, theme_idx, topic):
    prs = Presentation()
    prs.slide_width = Inch(13.33)
    prs.slide_height = Inch(7.5)

    slides_data = parse_slides(content, slide_count)
    theme = THEMES[theme_idx % len(THEMES)]

    # Get background image
    topic_en = await translate_topic(topic)

    for i, (slide_title, slide_text) in enumerate(slides_data):
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        # Try to get image for this slide
        img_data = None
        if i > 0:
            try:
                slide_en = await translate_topic(slide_title)
                img_data = await get_topic_image(f"{topic_en} {slide_en}")
            except:
                pass

        # Background
        bg = slide.background.fill
        bg.solid()
        bg.fore_color.rgb = PRGB(*theme["bg"])

        if i == 0:
            # ===== TITLE SLIDE =====
            # Full background image with overlay
            if img_data:
                try:
                    img_data.seek(0)
                    slide.shapes.add_picture(img_data, Inch(0), Inch(0), Inch(13.33), Inch(7.5))
                    # Dark overlay
                    overlay = slide.shapes.add_shape(1, Inch(0), Inch(0), Inch(13.33), Inch(7.5))
                    overlay.fill.solid()
                    overlay.fill.fore_color.rgb = PRGB(*theme["bg"])
                    overlay.fill.fore_color.theme_color = None
                    from pptx.util import Pt as PPt2
                    overlay.fill.transparency = 0.3
                    overlay.line.fill.background()
                except:
                    pass

            # Bottom gradient bar
            bar = slide.shapes.add_shape(1, Inch(0), Inch(6.2), Inch(13.33), Inch(1.3))
            bar.fill.solid()
            bar.fill.fore_color.rgb = PRGB(*theme["accent"])
            bar.line.fill.background()

            # Main title
            tb = slide.shapes.add_textbox(Inch(0.8), Inch(1.8), Inch(11.5), Inch(3))
            tf = tb.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = slide_title
            p.font.size = PPt(50)
            p.font.bold = True
            p.font.color.rgb = PRGB(255, 255, 255)
            p.alignment = PP_ALIGN.CENTER

            # Subtitle bar
            sb = slide.shapes.add_textbox(Inch(0.8), Inch(6.3), Inch(11.5), Inch(0.8))
            sf = sb.text_frame
            sp = sf.paragraphs[0]
            sp.text = f"✦  {style} uslubi  ✦  {datetime.now().strftime('%d.%m.%Y')}  ✦  {slide_count} ta slayd  ✦"
            sp.font.size = PPt(16)
            sp.font.bold = True
            sp.font.color.rgb = PRGB(*theme["bg"])
            sp.alignment = PP_ALIGN.CENTER

        else:
            # ===== CONTENT SLIDE =====

            if img_data:
                try:
                    img_data.seek(0)
                    # Image on right side
                    slide.shapes.add_picture(img_data, Inch(8.5), Inch(1.3), Inch(4.7), Inch(5.8))
                    # Image overlay
                    img_overlay = slide.shapes.add_shape(1, Inch(8.5), Inch(1.3), Inch(4.7), Inch(5.8))
                    img_overlay.fill.solid()
                    img_overlay.fill.fore_color.rgb = PRGB(*theme["bg"])
                    img_overlay.fill.transparency = 0.5
                    img_overlay.line.fill.background()
                except:
                    pass

            # Left accent bar
            lbar = slide.shapes.add_shape(1, Inch(0), Inch(0), Inch(0.12), Inch(7.5))
            lbar.fill.solid()
            lbar.fill.fore_color.rgb = PRGB(*theme["accent"])
            lbar.line.fill.background()

            # Top header bar
            hbar = slide.shapes.add_shape(1, Inch(0.12), Inch(0), Inch(13.21), Inch(1.2))
            hbar.fill.solid()
            hbar.fill.fore_color.rgb = PRGB(theme["bg"][0]+15, theme["bg"][1]+15, theme["bg"][2]+15)
            hbar.line.fill.background()

            # Title in header
            ttb = slide.shapes.add_textbox(Inch(0.3), Inch(0.1), Inch(12), Inch(1))
            ttf = ttb.text_frame
            ttp = ttf.paragraphs[0]
            ttp.text = slide_title
            ttp.font.size = PPt(28)
            ttp.font.bold = True
            ttp.font.color.rgb = PRGB(*theme["title"])

            # Accent line under title
            aline = slide.shapes.add_shape(1, Inch(0.3), Inch(1.2), Inch(5), Emu(35000))
            aline.fill.solid()
            aline.fill.fore_color.rgb = PRGB(*theme["accent"])
            aline.line.fill.background()

            # Slide number
            num_box = slide.shapes.add_shape(9, Inch(12.55), Inch(0.25), Inch(0.6), Inch(0.6))
            num_box.fill.solid()
            num_box.fill.fore_color.rgb = PRGB(*theme["accent"])
            num_box.line.fill.background()

            nb = slide.shapes.add_textbox(Inch(12.55), Inch(0.25), Inch(0.6), Inch(0.6))
            nf = nb.text_frame
            np_ = nf.paragraphs[0]
            np_.text = str(i)
            np_.font.size = PPt(16)
            np_.font.bold = True
            np_.font.color.rgb = PRGB(255, 255, 255)
            np_.alignment = PP_ALIGN.CENTER

            # Content area (left 8 inches if image, full if no image)
            content_width = Inch(7.8) if img_data else Inch(12.5)

            ctb = slide.shapes.add_textbox(Inch(0.3), Inch(1.4), content_width, Inch(5.7))
            ctf = ctb.text_frame
            ctf.word_wrap = True

            # Split into bullet points
            sentences = []
            for s in slide_text.replace(";", ".").split("."):
                s = s.strip()
                if s and len(s) > 8:
                    sentences.append(s)

            first = True
            for j, sent in enumerate(sentences[:6]):
                if first:
                    cp = ctf.paragraphs[0]
                    first = False
                else:
                    cp = ctf.add_paragraph()

                # Alternating bullet colors
                bullet_colors = [theme["accent"], theme["sub"], theme["text"]]
                bullet_color = bullet_colors[j % 3]

                cp.text = f"  ◆  {sent}"
                cp.font.size = PPt(17)
                cp.font.color.rgb = PRGB(*theme["text"])
                cp.space_after = PPt(10)

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
    b_style = ParagraphStyle('B', fontSize=11, spaceAfter=8, leading=18, alignment=4)

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
        elif line[0].isdigit() and '. ' in line[:5]:
            story.append(Paragraph(f"<b>{line}</b>", h_style))
        else:
            story.append(Paragraph(line, b_style))

    doc.build(story)
    buffer.seek(0)
    return buffer

def create_docx(title, content, style, doc_type):
    doc = DocxDocument()
    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run(title)
    r.font.size = Pt(20)
    r.font.bold = True
    r.font.color.rgb = RGBColor(26, 82, 118)
    m = doc.add_paragraph()
    m.alignment = WD_ALIGN_PARAGRAPH.CENTER
    mr = m.add_run(f"{doc_type.upper()} | {style} uslubi | {datetime.now().strftime('%d.%m.%Y')}")
    mr.font.size = Pt(9)
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
        elif line[0].isdigit() and '. ' in line[:5]:
            h = doc.add_paragraph()
            hr = h.add_run(line)
            hr.font.size = Pt(12)
            hr.font.bold = True
        else:
            p = doc.add_paragraph(line)
            if p.runs:
                p.runs[0].font.size = Pt(11)
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ===================== BOT HANDLERS =====================

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    name = update.effective_user.first_name
    if uid not in user_data_store:
        user_data_store[uid] = {"name": name, "documents": 0}
    kb = [
        [InlineKeyboardButton("📝 Referat", callback_data="doc_referat"),
         InlineKeyboardButton("📚 Kurs Ishi", callback_data="doc_kurs")],
        [InlineKeyboardButton("📄 Maqola", callback_data="doc_maqola"),
         InlineKeyboardButton("🎯 Slide (PPTX)", callback_data="doc_slide")],
        [InlineKeyboardButton("📊 Statistika", callback_data="stats"),
         InlineKeyboardButton("❓ Yordam", callback_data="help")]
    ]
    await update.message.reply_text(
        f"🎓 Assalomu alaykum, *{name}*!\n\n"
        "Faqat *mavzuni* yozing — qolganini bot o'zi qiladi!\n\n"
        "✨ *Yangi xususiyatlar:*\n"
        "🖼 Rasmli slaydlar\n"
        "🎨 10 ta zamonaviy dizayn\n"
        "🤖 AI mazmun generatsiyasi\n"
        "📄 PDF + DOCX + PPTX\n\n"
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
    kb = [
        [InlineKeyboardButton("🔵 Ko'k Premium", callback_data="theme_0"),
         InlineKeyboardButton("🟢 Yashil", callback_data="theme_1")],
        [InlineKeyboardButton("🟣 Binafsha", callback_data="theme_2"),
         InlineKeyboardButton("🔴 Qizil", callback_data="theme_3")],
        [InlineKeyboardButton("🩵 Moviy", callback_data="theme_4"),
         InlineKeyboardButton("🟡 Oltin", callback_data="theme_5")],
        [InlineKeyboardButton("⚫ Qora", callback_data="theme_6"),
         InlineKeyboardButton("🟤 Jigarrang", callback_data="theme_7")],
        [InlineKeyboardButton("🌊 Dengiz", callback_data="theme_8"),
         InlineKeyboardButton("🌿 Nefrit", callback_data="theme_9")]
    ]
    await q.edit_message_text(
        f"✅ *{count} ta slayd* tanlandi\n\n🎨 *10 ta dizayndan birini tanlang:*",
        reply_markup=InlineKeyboardMarkup(kb), parse_mode=ParseMode.MARKDOWN
    )

async def cb_theme(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    uid = q.from_user.id
    tidx = int(q.data.replace("theme_", ""))
    if uid not in user_data_store:
        user_data_store[uid] = {}
    user_data_store[uid]["theme"] = tidx
    theme_name = THEMES[tidx]["name"]
    await q.edit_message_text(
        f"✅ *{theme_name}* tanlandi\n\n✏️ *Mavzuni yozing:*\n\n_Misol: \"Sun'iy intellekt va ta'lim\"_",
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
            f"✅ Mavzu: *{text}*\n\n"
            "⏳ Hujjat tayyorlanmoqda...\n"
            "🤖 AI mazmun yozmoqda...\n"
            "🖼 Rasmlar yuklanmoqda...\n"
            "_(1-2 daqiqa kuting)_",
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
            buf = await create_pptx_with_images(title, content, style, scount, theme, title)
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
            text=f"✅ *{fname}* yuborildi!\n\n📝 {dtype} | 📋 {style} | 📄 {fmt.upper()}",
            reply_markup=InlineKeyboardMarkup(kb),
            parse_mode=ParseMode.MARKDOWN
        )
    except Exception as e:
        await context.bot.send_message(chat_id=uid, text=f"❌ Xato: {str(e)}\n\n/start bosing")

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
        "3️⃣ Slayd soni tanlang\n"
        "4️⃣ Dizayn rangini tanlang\n"
        "5️⃣ Faqat *mavzuni* yozing\n"
        "6️⃣ Bot o'zi yozib, rasmli PPTX tayyorlaydi! ✅\n\n"
        "⚡ 1-2 daqiqada tayyor!",
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
    print("🖼 RASMLI SLAYDLAR ACTIVE!")
    print("🎨 10 TA DIZAYN MAVJUD!")
    print("🤖 AI CONTENT ACTIVE!")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()

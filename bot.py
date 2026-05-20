import os
import asyncio
import aiohttp
import json
import re
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, CallbackQueryHandler, MessageHandler, filters, ContextTypes
from telegram.constants import ParseMode
from dotenv import load_dotenv
from datetime import datetime
from io import BytesIO

# Google APIs
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# PDF/DOCX
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable
from reportlab.lib import colors
from reportlab.lib.units import cm
from docx import Document as DocxDocument
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

load_dotenv()

BOT_TOKEN = os.getenv("BOT_TOKEN", "8625557628:AAHeUC2WxfMjJk-RRq3IxTtUJoc0H4XSsAM")
ADMIN_ID = int(os.getenv("ADMIN_ID", "7758296066"))
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")
GOOGLE_CREDENTIALS = os.getenv("GOOGLE_CREDENTIALS", "")
GOOGLE_SERVICE_EMAIL = os.getenv("GOOGLE_SERVICE_EMAIL", "")

user_data_store = {}

# ===================== GOOGLE SLIDES =====================

def get_google_services():
    """Get Google Slides and Drive services"""
    try:
        creds_dict = json.loads(GOOGLE_CREDENTIALS)
        creds = service_account.Credentials.from_service_account_info(
            creds_dict,
            scopes=[
                'https://www.googleapis.com/auth/presentations',
                'https://www.googleapis.com/auth/drive'
            ]
        )
        slides_service = build('slides', 'v1', credentials=creds)
        drive_service = build('drive', 'v3', credentials=creds)
        return slides_service, drive_service
    except Exception as e:
        print(f"Google services error: {e}")
        return None, None

async def create_google_slides(title, slides_content, theme_idx, style):
    """Create professional Google Slides presentation"""
    
    slides_service, drive_service = get_google_services()
    if not slides_service:
        return None
    
    # Color themes - PROFESSIONAL
    themes = [
        {"bg": {"red": 0.06, "green": 0.13, "blue": 0.25}, "accent": {"red": 0.20, "green": 0.60, "blue": 0.86}, "text": {"red": 1, "green": 1, "blue": 1}},
        {"bg": {"red": 0.04, "green": 0.18, "blue": 0.12}, "accent": {"red": 0.15, "green": 0.68, "blue": 0.38}, "text": {"red": 1, "green": 1, "blue": 1}},
        {"bg": {"red": 0.16, "green": 0.04, "blue": 0.27}, "accent": {"red": 0.61, "green": 0.35, "blue": 0.71}, "text": {"red": 1, "green": 1, "blue": 1}},
        {"bg": {"red": 0.27, "green": 0.04, "blue": 0.04}, "accent": {"red": 0.91, "green": 0.30, "blue": 0.24}, "text": {"red": 1, "green": 1, "blue": 1}},
        {"bg": {"red": 0.02, "green": 0.16, "blue": 0.24}, "accent": {"red": 0.10, "green": 0.74, "blue": 0.61}, "text": {"red": 1, "green": 1, "blue": 1}},
        {"bg": {"red": 0.12, "green": 0.10, "blue": 0.02}, "accent": {"red": 0.95, "green": 0.77, "blue": 0.06}, "text": {"red": 1, "green": 1, "blue": 1}},
        {"bg": {"red": 0.06, "green": 0.06, "blue": 0.06}, "accent": {"red": 0.78, "green": 0.78, "blue": 0.78}, "text": {"red": 1, "green": 0.84, "blue": 0}},
        {"bg": {"red": 0.20, "green": 0.10, "blue": 0.04}, "accent": {"red": 0.83, "green": 0.33, "blue": 0}, "text": {"red": 1, "green": 1, "blue": 1}},
        {"bg": {"red": 0.02, "green": 0.12, "blue": 0.24}, "accent": {"red": 0.16, "green": 0.50, "blue": 0.73}, "text": {"red": 1, "green": 0.84, "blue": 0}},
        {"bg": {"red": 0.04, "green": 0.16, "blue": 0.12}, "accent": {"red": 0.09, "green": 0.63, "blue": 0.52}, "text": {"red": 1, "green": 1, "blue": 1}},
    ]
    
    theme = themes[theme_idx % len(themes)]
    
    try:
        # 1. Create presentation
        presentation = slides_service.presentations().create(
            body={"title": title}
        ).execute()
        presentation_id = presentation['presentationId']
        
        # 2. Get existing slide ID
        existing_slide_id = presentation['slides'][0]['objectId']
        
        # 3. Build all requests
        requests = []
        
        # Delete default slide
        requests.append({
            "deleteObject": {"objectId": existing_slide_id}
        })
        
        # Add slides
        slide_ids = []
        for i, (slide_title, slide_text) in enumerate(slides_content):
            slide_id = f"slide_{i}"
            title_id = f"title_{i}"
            body_id = f"body_{i}"
            accent_id = f"accent_{i}"
            num_id = f"num_{i}"
            
            slide_ids.append(slide_id)
            
            # Add slide
            requests.append({
                "addSlide": {
                    "objectId": slide_id,
                    "slideLayoutReference": {"predefinedLayout": "BLANK"}
                }
            })
            
            # Background color
            requests.append({
                "updatePageProperties": {
                    "objectId": slide_id,
                    "pageProperties": {
                        "pageBackgroundFill": {
                            "solidFill": {"color": {"rgbColor": theme["bg"]}}
                        }
                    },
                    "fields": "pageBackgroundFill"
                }
            })
            
            if i == 0:
                # ===== TITLE SLIDE =====
                
                # Bottom accent bar
                requests.append({
                    "createShape": {
                        "objectId": accent_id,
                        "shapeType": "RECTANGLE",
                        "elementProperties": {
                            "pageObjectId": slide_id,
                            "size": {"width": {"magnitude": 9144000, "unit": "EMU"}, "height": {"magnitude": 900000, "unit": "EMU"}},
                            "transform": {"scaleX": 1, "scaleY": 1, "translateX": 0, "translateY": 5486400, "unit": "EMU"}
                        }
                    }
                })
                requests.append({
                    "updateShapeProperties": {
                        "objectId": accent_id,
                        "shapeProperties": {
                            "shapeBackgroundFill": {"solidFill": {"color": {"rgbColor": theme["accent"]}}},
                            "outline": {"propertyState": "NOT_RENDERED"}
                        },
                        "fields": "shapeBackgroundFill,outline"
                    }
                })
                
                # Title text box
                requests.append({
                    "createShape": {
                        "objectId": title_id,
                        "shapeType": "TEXT_BOX",
                        "elementProperties": {
                            "pageObjectId": slide_id,
                            "size": {"width": {"magnitude": 8144000, "unit": "EMU"}, "height": {"magnitude": 2400000, "unit": "EMU"}},
                            "transform": {"scaleX": 1, "scaleY": 1, "translateX": 500000, "translateY": 1800000, "unit": "EMU"}
                        }
                    }
                })
                requests.append({
                    "insertText": {"objectId": title_id, "text": slide_title}
                })
                requests.append({
                    "updateTextStyle": {
                        "objectId": title_id,
                        "textRange": {"type": "ALL"},
                        "style": {
                            "bold": True,
                            "fontSize": {"magnitude": 40, "unit": "PT"},
                            "foregroundColor": {"opaqueColor": {"rgbColor": theme["text"]}}
                        },
                        "fields": "bold,fontSize,foregroundColor"
                    }
                })
                requests.append({
                    "updateParagraphStyle": {
                        "objectId": title_id,
                        "textRange": {"type": "ALL"},
                        "style": {"alignment": "CENTER"},
                        "fields": "alignment"
                    }
                })
                
                # Subtitle
                requests.append({
                    "createShape": {
                        "objectId": body_id,
                        "shapeType": "TEXT_BOX",
                        "elementProperties": {
                            "pageObjectId": slide_id,
                            "size": {"width": {"magnitude": 8144000, "unit": "EMU"}, "height": {"magnitude": 600000, "unit": "EMU"}},
                            "transform": {"scaleX": 1, "scaleY": 1, "translateX": 500000, "translateY": 5600000, "unit": "EMU"}
                        }
                    }
                })
                requests.append({
                    "insertText": {
                        "objectId": body_id,
                        "text": f"✦  {style} uslubi  ✦  {datetime.now().strftime('%d.%m.%Y')}  ✦"
                    }
                })
                requests.append({
                    "updateTextStyle": {
                        "objectId": body_id,
                        "textRange": {"type": "ALL"},
                        "style": {
                            "bold": True,
                            "fontSize": {"magnitude": 16, "unit": "PT"},
                            "foregroundColor": {"opaqueColor": {"rgbColor": theme["bg"]}}
                        },
                        "fields": "bold,fontSize,foregroundColor"
                    }
                })
                requests.append({
                    "updateParagraphStyle": {
                        "objectId": body_id,
                        "textRange": {"type": "ALL"},
                        "style": {"alignment": "CENTER"},
                        "fields": "alignment"
                    }
                })
                
            else:
                # ===== CONTENT SLIDE =====
                
                # Left accent bar
                requests.append({
                    "createShape": {
                        "objectId": accent_id,
                        "shapeType": "RECTANGLE",
                        "elementProperties": {
                            "pageObjectId": slide_id,
                            "size": {"width": {"magnitude": 180000, "unit": "EMU"}, "height": {"magnitude": 6858000, "unit": "EMU"}},
                            "transform": {"scaleX": 1, "scaleY": 1, "translateX": 0, "translateY": 0, "unit": "EMU"}
                        }
                    }
                })
                requests.append({
                    "updateShapeProperties": {
                        "objectId": accent_id,
                        "shapeProperties": {
                            "shapeBackgroundFill": {"solidFill": {"color": {"rgbColor": theme["accent"]}}},
                            "outline": {"propertyState": "NOT_RENDERED"}
                        },
                        "fields": "shapeBackgroundFill,outline"
                    }
                })
                
                # Header bar
                header_id = f"header_{i}"
                requests.append({
                    "createShape": {
                        "objectId": header_id,
                        "shapeType": "RECTANGLE",
                        "elementProperties": {
                            "pageObjectId": slide_id,
                            "size": {"width": {"magnitude": 8964000, "unit": "EMU"}, "height": {"magnitude": 1000000, "unit": "EMU"}},
                            "transform": {"scaleX": 1, "scaleY": 1, "translateX": 180000, "translateY": 0, "unit": "EMU"}
                        }
                    }
                })
                requests.append({
                    "updateShapeProperties": {
                        "objectId": header_id,
                        "shapeProperties": {
                            "shapeBackgroundFill": {
                                "solidFill": {
                                    "color": {"rgbColor": {
                                        "red": min(theme["bg"]["red"] + 0.08, 1),
                                        "green": min(theme["bg"]["green"] + 0.08, 1),
                                        "blue": min(theme["bg"]["blue"] + 0.08, 1)
                                    }}
                                }
                            },
                            "outline": {"propertyState": "NOT_RENDERED"}
                        },
                        "fields": "shapeBackgroundFill,outline"
                    }
                })
                
                # Slide number circle
                requests.append({
                    "createShape": {
                        "objectId": num_id,
                        "shapeType": "ELLIPSE",
                        "elementProperties": {
                            "pageObjectId": slide_id,
                            "size": {"width": {"magnitude": 500000, "unit": "EMU"}, "height": {"magnitude": 500000, "unit": "EMU"}},
                            "transform": {"scaleX": 1, "scaleY": 1, "translateX": 8544000, "translateY": 200000, "unit": "EMU"}
                        }
                    }
                })
                requests.append({
                    "updateShapeProperties": {
                        "objectId": num_id,
                        "shapeProperties": {
                            "shapeBackgroundFill": {"solidFill": {"color": {"rgbColor": theme["accent"]}}},
                            "outline": {"propertyState": "NOT_RENDERED"}
                        },
                        "fields": "shapeBackgroundFill,outline"
                    }
                })
                requests.append({
                    "insertText": {"objectId": num_id, "text": str(i)}
                })
                requests.append({
                    "updateTextStyle": {
                        "objectId": num_id,
                        "textRange": {"type": "ALL"},
                        "style": {
                            "bold": True,
                            "fontSize": {"magnitude": 18, "unit": "PT"},
                            "foregroundColor": {"opaqueColor": {"rgbColor": {"red": 1, "green": 1, "blue": 1}}}
                        },
                        "fields": "bold,fontSize,foregroundColor"
                    }
                })
                requests.append({
                    "updateParagraphStyle": {
                        "objectId": num_id,
                        "textRange": {"type": "ALL"},
                        "style": {"alignment": "CENTER"},
                        "fields": "alignment"
                    }
                })
                
                # Title
                requests.append({
                    "createShape": {
                        "objectId": title_id,
                        "shapeType": "TEXT_BOX",
                        "elementProperties": {
                            "pageObjectId": slide_id,
                            "size": {"width": {"magnitude": 8500000, "unit": "EMU"}, "height": {"magnitude": 800000, "unit": "EMU"}},
                            "transform": {"scaleX": 1, "scaleY": 1, "translateX": 250000, "translateY": 100000, "unit": "EMU"}
                        }
                    }
                })
                requests.append({
                    "insertText": {"objectId": title_id, "text": slide_title}
                })
                requests.append({
                    "updateTextStyle": {
                        "objectId": title_id,
                        "textRange": {"type": "ALL"},
                        "style": {
                            "bold": True,
                            "fontSize": {"magnitude": 26, "unit": "PT"},
                            "foregroundColor": {"opaqueColor": {"rgbColor": theme["text"]}}
                        },
                        "fields": "bold,fontSize,foregroundColor"
                    }
                })
                
                # Underline accent
                underline_id = f"underline_{i}"
                requests.append({
                    "createShape": {
                        "objectId": underline_id,
                        "shapeType": "RECTANGLE",
                        "elementProperties": {
                            "pageObjectId": slide_id,
                            "size": {"width": {"magnitude": 3000000, "unit": "EMU"}, "height": {"magnitude": 50000, "unit": "EMU"}},
                            "transform": {"scaleX": 1, "scaleY": 1, "translateX": 250000, "translateY": 1000000, "unit": "EMU"}
                        }
                    }
                })
                requests.append({
                    "updateShapeProperties": {
                        "objectId": underline_id,
                        "shapeProperties": {
                            "shapeBackgroundFill": {"solidFill": {"color": {"rgbColor": theme["accent"]}}},
                            "outline": {"propertyState": "NOT_RENDERED"}
                        },
                        "fields": "shapeBackgroundFill,outline"
                    }
                })
                
                # Content
                requests.append({
                    "createShape": {
                        "objectId": body_id,
                        "shapeType": "TEXT_BOX",
                        "elementProperties": {
                            "pageObjectId": slide_id,
                            "size": {"width": {"magnitude": 8600000, "unit": "EMU"}, "height": {"magnitude": 5200000, "unit": "EMU"}},
                            "transform": {"scaleX": 1, "scaleY": 1, "translateX": 250000, "translateY": 1150000, "unit": "EMU"}
                        }
                    }
                })
                
                # Format content as bullet points
                bullets = []
                for sent in slide_text.replace(";", ".").split("."):
                    sent = sent.strip()
                    if sent and len(sent) > 8:
                        bullets.append(f"◆  {sent}")
                
                content_text = "\n".join(bullets[:6])
                requests.append({
                    "insertText": {"objectId": body_id, "text": content_text}
                })
                requests.append({
                    "updateTextStyle": {
                        "objectId": body_id,
                        "textRange": {"type": "ALL"},
                        "style": {
                            "fontSize": {"magnitude": 17, "unit": "PT"},
                            "foregroundColor": {"opaqueColor": {"rgbColor": {
                                "red": 0.85, "green": 0.90, "blue": 0.95
                            }}}
                        },
                        "fields": "fontSize,foregroundColor"
                    }
                })
                requests.append({
                    "updateParagraphStyle": {
                        "objectId": body_id,
                        "textRange": {"type": "ALL"},
                        "style": {"lineSpacing": 150, "spaceAbove": {"magnitude": 8, "unit": "PT"}},
                        "fields": "lineSpacing,spaceAbove"
                    }
                })
        
        # Execute all requests
        slides_service.presentations().batchUpdate(
            presentationId=presentation_id,
            body={"requests": requests}
        ).execute()
        
        # Make file accessible
        drive_service.permissions().create(
            fileId=presentation_id,
            body={"type": "anyone", "role": "reader"}
        ).execute()
        
        # Export as PPTX
        export_url = f"https://docs.google.com/presentation/d/{presentation_id}/export/pptx"
        
        creds_dict = json.loads(GOOGLE_CREDENTIALS)
        creds = service_account.Credentials.from_service_account_info(
            creds_dict,
            scopes=['https://www.googleapis.com/auth/drive']
        )
        
        request = drive_service.files().export_media(
            fileId=presentation_id,
            mimeType='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
        
        buffer = BytesIO()
        downloader = MediaIoBaseDownload(buffer, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        
        # Delete from Drive after download
        drive_service.files().delete(fileId=presentation_id).execute()
        
        buffer.seek(0)
        return buffer
        
    except Exception as e:
        print(f"Google Slides error: {e}")
        return None

# ===================== AI CONTENT =====================

async def generate_with_gemini(prompt):
    if not GEMINI_API_KEY:
        return None
    try:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={GEMINI_API_KEY}"
        payload = {
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {"temperature": 0.7, "maxOutputTokens": 3000}
        }
        async with aiohttp.ClientSession() as session:
            async with session.post(url, json=payload, timeout=aiohttp.ClientTimeout(total=60)) as resp:
                if resp.status == 200:
                    data = await resp.json()
                    return data["candidates"][0]["content"]["parts"][0]["text"]
    except:
        pass
    return None

async def generate_slides(topic, slide_count, style):
    prompt = f"""Sen O'zbekistonning eng yaxshi akademik mutaxassisisan.
"{topic}" mavzusida {slide_count} ta professional slayd uchun BATAFSIL va REAL ma'lumotlar yoz.

MUHIM: Faqat HAQIQIY faktlar, sanalar, raqamlar va dalillar ishlat!

FORMAT (aniq):
===SLAYD_1===
SARLAVHA: {topic}
MAZMUN: Qisqacha kirish jumlasi.

===SLAYD_2===
SARLAVHA: Mavzuga kirish
MAZMUN: Batafsil tushuntirish, real faktlar bilan.

(va hokazo {slide_count} ta slaydgacha)

Uslub: {style} | TIL: Faqat O'zbekcha
Har bir slayd uchun 3-5 ta real va muhim fakt."""

    result = await generate_with_gemini(prompt)
    if result and "===SLAYD_" in result:
        return parse_slides_content(result, topic, slide_count)
    return get_fallback_slides(topic, slide_count)

def parse_slides_content(content, topic, slide_count):
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
    return slides[:slide_count] if slides else [(topic, "Ma'lumot")]

def get_fallback_slides(topic, slide_count):
    base = [
        (topic, f"Ushbu prezentatsiya {topic} mavzusining barcha muhim jihatlarini qamrab oladi. Professional va akademik yondashuv asosida tayyorlangan."),
        (f"{topic} haqida", f"{topic} - zamonaviy fanda muhim o'rin egallaydi. Bu soha jadal rivojlanmoqda va yangi imkoniyatlar yaratmoqda."),
        ("Tarixiy rivojlanish", f"{topic} ning tarixi uzoq o'tmishga ega. Har bir davr o'zining muhim kashfiyotlari va yutuqlari bilan ajralib turadi."),
        ("Asosiy tushunchalar", f"{topic} sohasidagi asosiy tushunchalar va terminlar. Ularni to'g'ri tushunish sohani chuqur o'rganish uchun zarur."),
        ("Afzalliklari", f"{topic} ning asosiy afzalliklari: samaradorlik, keng qo'llanilish, tejamkorlik va innovatsion rivojlanish imkoniyati."),
        ("Muammolar va yechimlar", f"{topic} sohasidagi asosiy muammolar va ularni hal qilish yo'llari. Zamonaviy yondashuvlar samarali natijalar bermoqda."),
        ("Jahon tajribasi", f"Dunyoning yetakchi mamlakatlari {topic} sohasida katta yutuqlarga erishgan. Ularning tajribasi o'rganishga arziydi."),
        ("O'zbekistonda", f"O'zbekistonda {topic} rivojlantirishga alohida e'tibor qaratilmoqda. Yangi dasturlar va loyihalar amalga oshirilmoqda."),
        ("Kelajak", f"{topic} ning kelajagi juda istiqbolli. Yangi texnologiyalar bu sohani yanada rivojlantiradi."),
        ("Xulosa", f"{topic} - kelajak sohasi. Barcha asosiy jihatlar ko'rib chiqildi. Bilimlarni amaliyotda qo'llash tavsiya etiladi."),
        ("Statistika", f"{topic} bo'yicha muhim raqamlar va ko'rsatkichlar. So'nggi yillarda sezilarli o'sish kuzatilgan."),
        ("Amaliy misol", f"{topic} ning real hayotdagi qo'llanilishi. Muvaffaqiyatli loyihalar va ularning natijalari."),
        ("Tavsiyalar", f"{topic} sohasida muvaffaqiyatga erishish uchun asosiy tavsiyalar. Mutaxassislar va yoshlar uchun yo'l-yo'riqlar."),
        ("Hamkorlik", f"{topic} sohasida xalqaro hamkorlik. Birgalikda erishilgan yutuqlar va kelajak rejalari."),
        ("E'tiboringiz uchun rahmat!", f"Savollar va muhokama uchun vaqt. Ushbu prezentatsiya {topic} mavzusini to'liq qamrab oldi."),
    ]
    return base[:slide_count]

async def generate_doc(topic, doc_type, style):
    doc_names = {"referat": "Referat", "kurs": "Kurs ishi", "maqola": "Ilmiy maqola"}
    prompt = f"""O'zbek tilida "{topic}" mavzusida {doc_names.get(doc_type, 'hujjat')} yoz.
Uslub: {style} | Hajm: 800-1000 so'z
FAQAT HAQIQIY faktlar, sanalar, raqamlar!

Tuzilma:
KIRISH
ASOSIY QISM
1. Nazariy asoslar
2. Amaliy jihatlar  
3. Muammolar va yechimlar
XULOSA
ADABIYOTLAR (5 ta real manba)"""

    result = await generate_with_gemini(prompt)
    if result:
        return result
    return f"""KIRISH

{topic} bugungi kunda muhim ahamiyat kasb etuvchi mavzu hisoblanadi.

ASOSIY QISM

1. NAZARIY ASOSLAR

{topic} sohasining nazariy poydevori ko'p yillar davomida shakllanib kelgan.

2. AMALIY JIHATLAR

{topic} amaliyotda keng qo'llaniladi va ijobiy natijalar bermoqda.

3. MUAMMOLAR VA YECHIMLAR

Mavjud muammolarni hal qilish uchun kompleks yondashuv zarur.

XULOSA

{topic} kelajakda yanada rivojlanib boradi.

ADABIYOTLAR
1. Mirziyoyev Sh.M. Yangi O'zbekiston strategiyasi. T.: 2021.
2. UNESCO Global Report 2023.
3. O'zbekiston milliy ensiklopediyasi. T.: 2020."""

def create_pdf(title, content, style, doc_type):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm, leftMargin=2.5*cm, rightMargin=2.5*cm)
    story = []
    base = getSampleStyleSheet()
    t_s = ParagraphStyle('T', fontSize=20, textColor=colors.HexColor('#1a5276'), spaceAfter=6, alignment=1, fontName='Helvetica-Bold')
    m_s = ParagraphStyle('M', fontSize=9, textColor=colors.HexColor('#888888'), spaceAfter=16, alignment=1)
    h_s = ParagraphStyle('H', fontSize=13, textColor=colors.HexColor('#2874a6'), spaceAfter=8, spaceBefore=14, fontName='Helvetica-Bold')
    b_s = ParagraphStyle('B', fontSize=11, spaceAfter=8, leading=18, alignment=4)
    story.append(Paragraph(title, t_s))
    story.append(Paragraph(f"{doc_type.upper()} | {style} | {datetime.now().strftime('%d.%m.%Y')}", m_s))
    story.append(HRFlowable(width="100%", thickness=2, color=colors.HexColor('#2874a6')))
    story.append(Spacer(1, 0.3*cm))
    for line in content.split('\n'):
        line = line.strip()
        if not line:
            story.append(Spacer(1, 0.15*cm))
        elif line.isupper() and len(line) < 60:
            story.append(Paragraph(line, h_s))
        elif len(line) > 2 and line[0].isdigit() and '.' in line[:3]:
            story.append(Paragraph(f"<b>{line}</b>", h_s))
        else:
            story.append(Paragraph(line, b_s))
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
    mr = m.add_run(f"{doc_type.upper()} | {style} | {datetime.now().strftime('%d.%m.%Y')}")
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
        elif len(line) > 2 and line[0].isdigit() and '.' in line[:3]:
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

# ===================== BOT =====================

THEME_NAMES = [
    "🔵 Ko'k Premium", "🟢 Yashil Natura", "🟣 Binafsha Royal",
    "🔴 Qizil Dinamik", "🩵 Moviy Ocean", "🟡 Oltin Biznes",
    "⚫ Qora Elegant", "🟤 Jigarrang", "🌊 Dengiz", "🌿 Nefrit"
]

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
        "✨ *Xususiyatlar:*\n"
        "🎨 Google Slides — professional dizayn\n"
        "🤖 AI — real ma'lumotlar\n"
        "📄 PDF + DOCX + PPTX\n"
        "🎯 10 ta rang tanlash\n\n"
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
        await q.edit_message_text("🎯 *Nechta slayd?*", reply_markup=InlineKeyboardMarkup(kb), parse_mode=ParseMode.MARKDOWN)
    else:
        kb = [
            [InlineKeyboardButton("📄 PDF", callback_data="fmt_pdf"),
             InlineKeyboardButton("📋 DOCX", callback_data="fmt_docx")]
        ]
        await q.edit_message_text("📄 *Format tanlang:*", reply_markup=InlineKeyboardMarkup(kb), parse_mode=ParseMode.MARKDOWN)

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
        [InlineKeyboardButton("🔵 Ko'k", callback_data="theme_0"),
         InlineKeyboardButton("🟢 Yashil", callback_data="theme_1"),
         InlineKeyboardButton("🟣 Binafsha", callback_data="theme_2")],
        [InlineKeyboardButton("🔴 Qizil", callback_data="theme_3"),
         InlineKeyboardButton("🩵 Moviy", callback_data="theme_4"),
         InlineKeyboardButton("🟡 Oltin", callback_data="theme_5")],
        [InlineKeyboardButton("⚫ Qora", callback_data="theme_6"),
         InlineKeyboardButton("🟤 Jigarrang", callback_data="theme_7"),
         InlineKeyboardButton("🌊 Dengiz", callback_data="theme_8")],
        [InlineKeyboardButton("🌿 Nefrit", callback_data="theme_9")]
    ]
    await q.edit_message_text(f"✅ *{count} ta slayd*\n\n🎨 *Rang tanlang:*",
                              reply_markup=InlineKeyboardMarkup(kb), parse_mode=ParseMode.MARKDOWN)

async def cb_theme(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    uid = q.from_user.id
    tidx = int(q.data.replace("theme_", ""))
    if uid not in user_data_store:
        user_data_store[uid] = {}
    user_data_store[uid]["theme"] = tidx
    await q.edit_message_text(
        f"✅ *{THEME_NAMES[tidx]}* tanlandi\n\n✏️ *Mavzuni yozing:*\n\n_Misol: \"Sun'iy intellekt\"_",
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
    await q.edit_message_text("✏️ *Mavzuni yozing:*\n\n_Misol: \"Kimyo va IT\"_", parse_mode=ParseMode.MARKDOWN)
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
            "⏳ Tayyorlanmoqda...\n"
            "🤖 AI yozmoqda...\n"
            "🎨 Google Slides dizayn qilmoqda...\n"
            "_(1-2 daqiqa kuting)_",
            parse_mode=ParseMode.MARKDOWN
        )
        await do_generate(update, context, uid)
    else:
        kb = [
            [InlineKeyboardButton("📝 Referat", callback_data="doc_referat"),
             InlineKeyboardButton("🎯 Slide", callback_data="doc_slide")]
        ]
        await update.message.reply_text("📋 Avval hujjat turini tanlang:", reply_markup=InlineKeyboardMarkup(kb))

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
            slides_list = await generate_slides(title, scount, style)
            buf = await create_google_slides(title, slides_list, theme, style)
            if not buf:
                # Fallback to simple PPTX
                from pptx import Presentation
                from pptx.util import Inches, Pt as PPt2
                from pptx.dml.color import RGBColor as PRGB2
                prs = Presentation()
                for st, sc in slides_list:
                    slide = prs.slides.add_slide(prs.slide_layouts[6])
                    tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(1.5))
                    tf = tb.text_frame
                    p = tf.paragraphs[0]
                    p.text = st
                    p.font.size = PPt2(28)
                    p.font.bold = True
                    tb2 = slide.shapes.add_textbox(Inches(0.5), Inches(2), Inches(9), Inches(4))
                    tf2 = tb2.text_frame
                    tf2.word_wrap = True
                    p2 = tf2.paragraphs[0]
                    p2.text = sc
                    p2.font.size = PPt2(16)
                buf = BytesIO()
                prs.save(buf)
                buf.seek(0)
            fname = f"{title[:25].replace(' ','_')}.pptx"
        elif fmt == "docx":
            content = await generate_doc(title, dtype, style)
            buf = create_docx(title, content, style, dtype)
            fname = f"{title[:25].replace(' ','_')}.docx"
        else:
            content = await generate_doc(title, dtype, style)
            buf = create_pdf(title, content, style, dtype)
            fname = f"{title[:25].replace(' ','_')}.pdf"

        await context.bot.send_document(chat_id=uid, document=buf, filename=fname)
        user_data_store[uid]["documents"] = user_data_store[uid].get("documents", 0) + 1
        kb = [
            [InlineKeyboardButton("📝 Yangi", callback_data="new_doc")],
            [InlineKeyboardButton("🏠 Menyu", callback_data="start")]
        ]
        await context.bot.send_message(
            chat_id=uid,
            text=f"✅ *{fname}* tayyor!\n📝 {dtype} | 📋 {style} | 📄 {fmt.upper()}",
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
    kb = [[InlineKeyboardButton("🏠 Menyu", callback_data="start")]]
    await q.edit_message_text(
        f"📊 *Statistika*\n\n👤 {info.get('name','N/A')}\n📄 {info.get('documents',0)} ta hujjat\n🆔 {uid}",
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
        "3️⃣ Slayd sonini tanlang\n"
        "4️⃣ Rangni tanlang\n"
        "5️⃣ Faqat *mavzuni* yozing\n"
        "6️⃣ Google Slides professional PPTX tayyor! ✅",
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
    print("🎨 GOOGLE SLIDES API ACTIVE!")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()

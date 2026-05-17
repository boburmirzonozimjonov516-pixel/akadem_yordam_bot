import os
from dotenv import load_dotenv

load_dotenv()

# Bot Settings
BOT_TOKEN = os.getenv("BOT_TOKEN", "8625557628:AAHeUC2WxfMjJk-RRq3IxTtUJoc0H4XSsAM")
ADMIN_ID = int(os.getenv("ADMIN_ID", "7758296066"))

# Database Settings
DATABASE_URL = os.getenv("DATABASE_URL", "sqlite:///bot_database.db")

# Payment Settings
CLICK_MERCHANT_ID = os.getenv("CLICK_MERCHANT_ID", "")
CLICK_SERVICE_ID = os.getenv("CLICK_SERVICE_ID", "")
PAYME_MERCHANT_ID = os.getenv("PAYME_MERCHANT_ID", "")
PAYME_API_KEY = os.getenv("PAYME_API_KEY", "")

# Pricing
FIRST_DOC_PRICE = 0  # Free
SUBSEQUENT_DOC_PRICE = 2000  # UZS
SUBSCRIPTION_PRICE = 2000  # UZS per month

# File Settings
MAX_FILE_SIZE = 10 * 1024 * 1024  # 10MB
ALLOWED_FORMATS = ["pdf", "docx", "txt"]

# Logging
LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO")

# Features
ENABLE_PDF = True
ENABLE_DOCX = True
ENABLE_PAYMENT = True
ENABLE_ADMIN_PANEL = True

# Styles
AVAILABLE_STYLES = ["apa", "harvard", "uzbek", "chicago"]

# Document Types
DOCUMENT_TYPES = {
    "referat": "📝 Referat",
    "kurs": "📚 Kurs Ishi",
    "maqola": "📄 Maqola",
    "slide": "🎯 Slide"
}

# Payment Gateways
PAYMENT_GATEWAYS = {
    "click": "💳 Click.uz",
    "payme": "💳 Payme",
    "telegram": "📱 Telegram Stars",
    "card": "💰 Karta"
}

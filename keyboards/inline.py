from telegram import InlineKeyboardButton, InlineKeyboardMarkup

# ===== MAIN MENU =====
def main_menu():
    """Main menu keyboard"""
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
            InlineKeyboardButton("💳 Premium", callback_data="premium")
        ],
        [
            InlineKeyboardButton("⚙️ Sozlamalar", callback_data="settings"),
            InlineKeyboardButton("❓ Yordam", callback_data="help")
        ]
    ]
    return InlineKeyboardMarkup(keyboard)


# ===== DOCUMENT TYPE MENU =====
def document_type_menu():
    """Select document type"""
    keyboard = [
        [InlineKeyboardButton("📝 Referat", callback_data="doc_referat")],
        [InlineKeyboardButton("📚 Kurs Ishi", callback_data="doc_kurs")],
        [InlineKeyboardButton("📄 Maqola", callback_data="doc_maqola")],
        [InlineKeyboardButton("🎯 Slide", callback_data="doc_slide")],
        [InlineKeyboardButton("🔙 Orqaga", callback_data="back_menu")]
    ]
    return InlineKeyboardMarkup(keyboard)


# ===== STYLE MENU =====
def style_menu():
    """Select academic style"""
    keyboard = [
        [
            InlineKeyboardButton("🎯 APA", callback_data="style_apa"),
            InlineKeyboardButton("🎯 Harvard", callback_data="style_harvard")
        ],
        [
            InlineKeyboardButton("🎯 Uzbek", callback_data="style_uzbek"),
            InlineKeyboardButton("🎯 Chicago", callback_data="style_chicago")
        ],
        [InlineKeyboardButton("🔙 Orqaga", callback_data="back_style")]
    ]
    return InlineKeyboardMarkup(keyboard)


# ===== FILE FORMAT MENU =====
def format_menu():
    """Select file format"""
    keyboard = [
        [
            InlineKeyboardButton("📄 PDF", callback_data="format_pdf"),
            InlineKeyboardButton("📋 DOCX", callback_data="format_docx")
        ],
        [
            InlineKeyboardButton("📝 TXT", callback_data="format_txt")
        ],
        [InlineKeyboardButton("🔙 Orqaga", callback_data="back_menu")]
    ]
    return InlineKeyboardMarkup(keyboard)


# ===== DOCUMENT EDIT MENU =====
def document_edit_menu():
    """Edit document options"""
    keyboard = [
        [InlineKeyboardButton("✅ Tugatish va Fayl Olish", callback_data="finish_doc")],
        [InlineKeyboardButton("➕ Yana Matn Qo'shish", callback_data="add_more")],
        [InlineKeyboardButton("✏️ Sarlavhani Tahrirlash", callback_data="edit_title")],
        [InlineKeyboardButton("❌ Bekor Qilish", callback_data="cancel_doc")]
    ]
    return InlineKeyboardMarkup(keyboard)


# ===== PAYMENT GATEWAY MENU =====
def payment_gateway_menu():
    """Select payment gateway"""
    keyboard = [
        [
            InlineKeyboardButton("💳 Click.uz", callback_data="pay_click"),
            InlineKeyboardButton("💳 Payme", callback_data="pay_payme")
        ],
        [
            InlineKeyboardButton("⭐ Telegram Stars", callback_data="pay_telegram"),
            InlineKeyboardButton("💰 Karta", callback_data="pay_card")
        ],
        [InlineKeyboardButton("🔙 Orqaga", callback_data="back_menu")]
    ]
    return InlineKeyboardMarkup(keyboard)


# ===== SUBSCRIPTION MENU =====
def subscription_menu():
    """Subscription options"""
    keyboard = [
        [InlineKeyboardButton("🎁 Birinchi (Tekin)", callback_data="sub_free")],
        [InlineKeyboardButton("💎 Premium", callback_data="sub_premium")],
        [InlineKeyboardButton("👑 Pro", callback_data="sub_pro")],
        [InlineKeyboardButton("📊 Narxlar", callback_data="pricing")],
        [InlineKeyboardButton("🔙 Orqaga", callback_data="back_menu")]
    ]
    return InlineKeyboardMarkup(keyboard)


# ===== FINISH DOCUMENT MENU =====
def finish_document_menu():
    """After document generation"""
    keyboard = [
        [InlineKeyboardButton("📝 Yangi Hujjat", callback_data="new_doc")],
        [InlineKeyboardButton("💳 Premium Obuna", callback_data="premium")],
        [InlineKeyboardButton("📊 Statistika", callback_data="stats")],
        [InlineKeyboardButton("🏠 Bosh Menyu", callback_data="back_menu")]
    ]
    return InlineKeyboardMarkup(keyboard)


# ===== SETTINGS MENU =====
def settings_menu():
    """User settings"""
    keyboard = [
        [InlineKeyboardButton("👤 Profil", callback_data="profile")],
        [InlineKeyboardButton("🔔 Bildirishnomalar", callback_data="notifications")],
        [InlineKeyboardButton("🔐 Xavfsizlik", callback_data="security")],
        [InlineKeyboardButton("📥 Yuklamalar", callback_data="downloads")],
        [InlineKeyboardButton("🔙 Orqaga", callback_data="back_menu")]
    ]
    return InlineKeyboardMarkup(keyboard)


# ===== ADMIN MENU =====
def admin_menu():
    """Admin panel menu"""
    keyboard = [
        [InlineKeyboardButton("👥 Foydalanuvchilar", callback_data="admin_users")],
        [InlineKeyboardButton("📊 Statistika", callback_data="admin_stats")],
        [InlineKeyboardButton("💰 To'lovlar", callback_data="admin_payments")],
        [InlineKeyboardButton("📋 Hujjatlar", callback_data="admin_documents")],
        [InlineKeyboardButton("📢 Xabar Yuborish", callback_data="admin_broadcast")],
        [InlineKeyboardButton("🔒 Kerakli Foydalanuvchilar", callback_data="admin_ban")],
        [InlineKeyboardButton("📈 Analytics", callback_data="admin_analytics")],
        [InlineKeyboardButton("🔙 Orqaga", callback_data="back_menu")]
    ]
    return InlineKeyboardMarkup(keyboard)


# ===== USER MENU =====
def user_management_menu(user_id: int):
    """Manage specific user (admin)"""
    keyboard = [
        [InlineKeyboardButton("👁️ Ko'rish", callback_data=f"view_user_{user_id}")],
        [InlineKeyboardButton("🚫 Ban Qilish", callback_data=f"ban_user_{user_id}")],
        [InlineKeyboardButton("💾 Hujjatlar", callback_data=f"user_docs_{user_id}")],
        [InlineKeyboardButton("💰 To'lovlar", callback_data=f"user_payments_{user_id}")],
        [InlineKeyboardButton("🔙 Orqaga", callback_data="admin_users")]
    ]
    return InlineKeyboardMarkup(keyboard)


# ===== CONFIRMATION MENU =====
def confirmation_menu(action: str, target_id: str = None):
    """Confirmation dialogs"""
    keyboard = [
        [
            InlineKeyboardButton("✅ Ha", callback_data=f"confirm_{action}_{target_id or ''}"),
            InlineKeyboardButton("❌ Yo'q", callback_data="cancel_action")
        ]
    ]
    return InlineKeyboardMarkup(keyboard)


# ===== HELP MENU =====
def help_menu():
    """Help options"""
    keyboard = [
        [InlineKeyboardButton("📖 Qo'llanma", callback_data="guide")],
        [InlineKeyboardButton("❓ FAQ", callback_data="faq")],
        [InlineKeyboardButton("🎥 Video", callback_data="video_tutorial")],
        [InlineKeyboardButton("💬 Support", callback_data="support_chat")],
        [InlineKeyboardButton("🔙 Orqaga", callback_data="back_menu")]
    ]
    return InlineKeyboardMarkup(keyboard)


# ===== DOCUMENT DOWNLOAD MENU =====
def document_download_menu(doc_id: str):
    """Download options"""
    keyboard = [
        [
            InlineKeyboardButton("📄 PDF", callback_data=f"download_{doc_id}_pdf"),
            InlineKeyboardButton("📋 DOCX", callback_data=f"download_{doc_id}_docx")
        ],
        [InlineKeyboardButton("📤 Ulashish", callback_data=f"share_{doc_id}")],
        [InlineKeyboardButton("🗑️ O'chirish", callback_data=f"delete_{doc_id}")],
        [InlineKeyboardButton("🔙 Orqaga", callback_data="my_documents")]
    ]
    return InlineKeyboardMarkup(keyboard)


# ===== PRICING MENU =====
def pricing_menu():
    """Pricing and plans"""
    keyboard = [
        [InlineKeyboardButton("🎁 Birinchi (Tekin)", callback_data="plan_free")],
        [InlineKeyboardButton("💎 Premium (2,000 UZS)", callback_data="plan_premium")],
        [InlineKeyboardButton("👑 Pro (5,000 UZS)", callback_data="plan_pro")],
        [InlineKeyboardButton("🔙 Orqaga", callback_data="back_menu")]
    ]
    return InlineKeyboardMarkup(keyboard)


# ===== QUICK ACTIONS =====
def quick_actions_menu():
    """Quick action buttons"""
    keyboard = [
        [InlineKeyboardButton("⚡ Tezlik Boshlash", callback_data="quick_start")],
        [InlineKeyboardButton("📋 So'nggi Hujjat", callback_data="last_document")],
        [InlineKeyboardButton("⭐ Sevimlilar", callback_data="favorites")]
    ]
    return InlineKeyboardMarkup(keyboard)


# ===== DOCUMENT HISTORY =====
def document_history_menu():
    """View document history"""
    keyboard = [
        [InlineKeyboardButton("📅 Bugun", callback_data="docs_today")],
        [InlineKeyboardButton("📆 Bu Hafta", callback_data="docs_week")],
        [InlineKeyboardButton("📊 Bu Oy", callback_data="docs_month")],
        [InlineKeyboardButton("📈 Barchasi", callback_data="docs_all")],
        [InlineKeyboardButton("🔙 Orqaga", callback_data="back_menu")]
    ]
    return InlineKeyboardMarkup(keyboard)


# ===== SHARE OPTIONS =====
def share_menu(doc_id: str):
    """Share document options"""
    keyboard = [
        [InlineKeyboardButton("📱 Telegram", callback_data=f"share_telegram_{doc_id}")],
        [InlineKeyboardButton("📧 Email", callback_data=f"share_email_{doc_id}")],
        [InlineKeyboardButton("🔗 Havolani Nusxala", callback_data=f"share_link_{doc_id}")],
        [InlineKeyboardButton("🔙 Orqaga", callback_data=f"document_{doc_id}")]
    ]
    return InlineKeyboardMarkup(keyboard)


# ===== INLINE BUTTONS =====
def inline_button(text: str, callback_data: str) -> InlineKeyboardButton:
    """Create single inline button"""
    return InlineKeyboardButton(text, callback_data=callback_data)


def row(*buttons) -> list:
    """Create button row"""
    return list(buttons)


# ===== UTILITY FUNCTIONS =====
def get_menu_by_name(menu_name: str):
    """Get menu by name"""
    menus = {
        "main": main_menu,
        "document_type": document_type_menu,
        "style": style_menu,
        "format": format_menu,
        "payment": payment_gateway_menu,
        "subscription": subscription_menu,
        "settings": settings_menu,
        "admin": admin_menu,
        "help": help_menu,
        "pricing": pricing_menu,
    }
    
    return menus.get(menu_name, main_menu)()


def paginate_buttons(items: list, page: int = 1, items_per_page: int = 5) -> tuple:
    """Paginate button list"""
    total_pages = (len(items) + items_per_page - 1) // items_per_page
    start_idx = (page - 1) * items_per_page
    end_idx = start_idx + items_per_page
    
    page_items = items[start_idx:end_idx]
    
    keyboard = [
        [InlineKeyboardButton(item[0], callback_data=item[1])]
        for item in page_items
    ]
    
    # Add pagination buttons
    if total_pages > 1:
        pagination_row = []
        if page > 1:
            pagination_row.append(InlineKeyboardButton("◀️", callback_data=f"page_{page-1}"))
        pagination_row.append(InlineKeyboardButton(f"{page}/{total_pages}", callback_data="page_info"))
        if page < total_pages:
            pagination_row.append(InlineKeyboardButton("▶️", callback_data=f"page_{page+1}"))
        keyboard.append(pagination_row)
    
    return InlineKeyboardMarkup(keyboard), total_pages

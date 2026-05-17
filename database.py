from sqlalchemy import create_engine, Column, Integer, String, Float, Boolean, DateTime, Text, ForeignKey
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, relationship
from datetime import datetime
import os
from dotenv import load_dotenv

load_dotenv()

DATABASE_URL = os.getenv("DATABASE_URL", "sqlite:///bot_database.db")

# Create engine
if DATABASE_URL.startswith("sqlite"):
    engine = create_engine(DATABASE_URL, connect_args={"check_same_thread": False})
else:
    engine = create_engine(DATABASE_URL)

SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

# Models
class User(Base):
    __tablename__ = "users"
    
    id = Column(Integer, primary_key=True, index=True)
    telegram_id = Column(Integer, unique=True, index=True)
    username = Column(String, nullable=True)
    first_name = Column(String)
    last_name = Column(String, nullable=True)
    
    # Account details
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    # Usage tracking
    first_doc_used = Column(Boolean, default=False)
    documents_count = Column(Integer, default=0)
    total_characters = Column(Integer, default=0)
    
    # Subscription
    subscription_active = Column(Boolean, default=False)
    subscription_start = Column(DateTime, nullable=True)
    subscription_end = Column(DateTime, nullable=True)
    subscription_tier = Column(String, default="free")  # free, premium, pro
    
    # Balance
    balance = Column(Float, default=0)
    total_spent = Column(Float, default=0)
    
    # Status
    is_banned = Column(Boolean, default=False)
    ban_reason = Column(String, nullable=True)
    
    # Relationships
    documents = relationship("Document", back_populates="user", cascade="all, delete-orphan")
    payments = relationship("Payment", back_populates="user", cascade="all, delete-orphan")
    
    def __repr__(self):
        return f"<User(id={self.id}, telegram_id={self.telegram_id}, username={self.username})>"


class Document(Base):
    __tablename__ = "documents"
    
    id = Column(String, primary_key=True, index=True)
    user_id = Column(Integer, ForeignKey("users.id"), index=True)
    
    # Document details
    doc_type = Column(String)  # referat, kurs, maqola, slide
    template_style = Column(String)  # apa, harvard, uzbek, chicago
    file_format = Column(String)  # pdf, docx, txt
    
    title = Column(String)
    content = Column(Text)
    
    # File info
    file_path = Column(String)
    file_size = Column(Integer, nullable=True)
    file_hash = Column(String, nullable=True)
    
    # Metadata
    word_count = Column(Integer, default=0)
    character_count = Column(Integer, default=0)
    page_count = Column(Integer, default=0)
    
    # Timestamps
    created_at = Column(DateTime, default=datetime.utcnow, index=True)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    downloaded_at = Column(DateTime, nullable=True)
    
    # Download tracking
    download_count = Column(Integer, default=0)
    
    # Relationship
    user = relationship("User", back_populates="documents")
    
    def __repr__(self):
        return f"<Document(id={self.id}, user_id={self.user_id}, type={self.doc_type})>"


class Payment(Base):
    __tablename__ = "payments"
    
    id = Column(String, primary_key=True, index=True)
    user_id = Column(Integer, ForeignKey("users.id"), index=True)
    
    # Payment details
    amount = Column(Float)
    currency = Column(String, default="UZS")
    gateway = Column(String)  # click, payme, telegram, card
    
    # Order info
    order_id = Column(String, nullable=True)
    transaction_id = Column(String, nullable=True, index=True)
    
    # Status
    status = Column(String, default="pending")  # pending, completed, failed, refunded
    
    # Description
    description = Column(String, nullable=True)
    error_message = Column(String, nullable=True)
    
    # Timestamps
    created_at = Column(DateTime, default=datetime.utcnow, index=True)
    completed_at = Column(DateTime, nullable=True)
    
    # Relationship
    user = relationship("User", back_populates="payments")
    
    def __repr__(self):
        return f"<Payment(id={self.id}, user_id={self.user_id}, amount={self.amount}, status={self.status})>"


class Subscription(Base):
    __tablename__ = "subscriptions"
    
    id = Column(String, primary_key=True, index=True)
    user_id = Column(Integer, ForeignKey("users.id"), unique=True, index=True)
    
    # Subscription details
    tier = Column(String)  # free, premium, pro
    status = Column(String, default="active")  # active, cancelled, expired
    
    # Dates
    start_date = Column(DateTime, default=datetime.utcnow)
    end_date = Column(DateTime)
    renewal_date = Column(DateTime, nullable=True)
    
    # Features
    documents_limit = Column(Integer)
    storage_limit = Column(Integer)  # MB
    supports_formats = Column(String)  # json: ["pdf", "docx"]
    
    # Payment
    auto_renew = Column(Boolean, default=True)
    payment_method = Column(String, nullable=True)
    
    # Timestamps
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    def __repr__(self):
        return f"<Subscription(user_id={self.user_id}, tier={self.tier}, status={self.status})>"


class AdminLog(Base):
    __tablename__ = "admin_logs"
    
    id = Column(Integer, primary_key=True, index=True)
    admin_id = Column(Integer)
    
    # Action details
    action = Column(String)  # ban_user, delete_document, broadcast, etc
    target_user_id = Column(Integer, nullable=True)
    target_id = Column(String, nullable=True)
    
    # Details
    details = Column(Text, nullable=True)
    result = Column(String)  # success, failed
    
    # Timestamp
    timestamp = Column(DateTime, default=datetime.utcnow, index=True)
    
    def __repr__(self):
        return f"<AdminLog(admin_id={self.admin_id}, action={self.action}, result={self.result})>"


class Analytics(Base):
    __tablename__ = "analytics"
    
    id = Column(Integer, primary_key=True, index=True)
    
    # Date
    date = Column(DateTime, default=datetime.utcnow, index=True)
    
    # Metrics
    active_users = Column(Integer, default=0)
    new_users = Column(Integer, default=0)
    documents_created = Column(Integer, default=0)
    total_revenue = Column(Float, default=0)
    success_payments = Column(Integer, default=0)
    failed_payments = Column(Integer, default=0)
    
    # Performance
    avg_response_time = Column(Float, nullable=True)
    error_count = Column(Integer, default=0)
    
    # Timestamp
    created_at = Column(DateTime, default=datetime.utcnow)
    
    def __repr__(self):
        return f"<Analytics(date={self.date}, active_users={self.active_users}, documents_created={self.documents_created})>"


# Database functions
def init_db():
    """Initialize database"""
    Base.metadata.create_all(bind=engine)
    print("✅ Database initialized successfully!")

def get_session():
    """Get database session"""
    return SessionLocal()

def get_user(user_id: int):
    """Get user by ID"""
    session = get_session()
    user = session.query(User).filter(User.telegram_id == user_id).first()
    session.close()
    return user

def create_user(telegram_id: int, username: str, first_name: str):
    """Create new user"""
    session = get_session()
    user = User(
        telegram_id=telegram_id,
        username=username,
        first_name=first_name,
        created_at=datetime.utcnow()
    )
    session.add(user)
    session.commit()
    session.close()
    return user

def get_stats():
    """Get global statistics"""
    session = get_session()
    
    stats = {
        "total_users": session.query(User).count(),
        "active_users": session.query(User).filter(User.subscription_active == True).count(),
        "total_documents": session.query(Document).count(),
        "total_payments": session.query(Payment).filter(Payment.status == "completed").count(),
        "total_revenue": session.query(Payment).filter(Payment.status == "completed").with_entities(
            __import__('sqlalchemy').func.sum(Payment.amount)
        ).scalar() or 0,
        "documents_today": session.query(Document).filter(
            Document.created_at >= datetime.utcnow().replace(hour=0, minute=0, second=0, microsecond=0)
        ).count(),
    }
    
    session.close()
    return stats

def add_document(user_id: int, doc_id: str, doc_type: str, style: str, title: str, 
                content: str, file_path: str, file_format: str = "pdf"):
    """Add new document"""
    session = get_session()
    
    word_count = len(content.split())
    char_count = len(content)
    
    document = Document(
        id=doc_id,
        user_id=user_id,
        doc_type=doc_type,
        template_style=style,
        file_format=file_format,
        title=title,
        content=content,
        file_path=file_path,
        word_count=word_count,
        character_count=char_count,
        created_at=datetime.utcnow()
    )
    
    session.add(document)
    
    # Update user stats
    user = session.query(User).filter(User.telegram_id == user_id).first()
    if user:
        user.documents_count += 1
        user.total_characters += char_count
    
    session.commit()
    session.close()
    return document

def add_payment(user_id: int, payment_id: str, amount: float, gateway: str, 
               order_id: str = None, description: str = None):
    """Add new payment"""
    session = get_session()
    
    payment = Payment(
        id=payment_id,
        user_id=user_id,
        amount=amount,
        gateway=gateway,
        order_id=order_id,
        description=description,
        status="pending",
        created_at=datetime.utcnow()
    )
    
    session.add(payment)
    session.commit()
    session.close()
    return payment

def complete_payment(payment_id: str, transaction_id: str = None):
    """Complete payment"""
    session = get_session()
    
    payment = session.query(Payment).filter(Payment.id == payment_id).first()
    if payment:
        payment.status = "completed"
        payment.transaction_id = transaction_id
        payment.completed_at = datetime.utcnow()
        
        # Update user
        user = session.query(User).filter(User.telegram_id == payment.user_id).first()
        if user:
            user.balance += payment.amount
            user.total_spent += payment.amount
            user.subscription_active = True
        
        session.commit()
    
    session.close()
    return payment

def ban_user(user_id: int, reason: str = None):
    """Ban user"""
    session = get_session()
    
    user = session.query(User).filter(User.telegram_id == user_id).first()
    if user:
        user.is_banned = True
        user.ban_reason = reason
        session.commit()
    
    session.close()
    return user

def log_admin_action(admin_id: int, action: str, target_user_id: int = None, 
                    target_id: str = None, details: str = None, result: str = "success"):
    """Log admin action"""
    session = get_session()
    
    log = AdminLog(
        admin_id=admin_id,
        action=action,
        target_user_id=target_user_id,
        target_id=target_id,
        details=details,
        result=result,
        timestamp=datetime.utcnow()
    )
    
    session.add(log)
    session.commit()
    session.close()
    return log

if __name__ == "__main__":
    init_db()

import hashlib
import hmac
import json
import logging
from datetime import datetime
from typing import Optional, Dict
import aiohttp

logger = logging.getLogger(__name__)

class PaymentService:
    """Unified payment service for multiple gateways"""
    
    def __init__(self, config: Dict):
        self.config = config
        self.click_merchant_id = config.get("CLICK_MERCHANT_ID")
        self.click_service_id = config.get("CLICK_SERVICE_ID")
        self.payme_merchant_id = config.get("PAYME_MERCHANT_ID")
        self.payme_api_key = config.get("PAYME_API_KEY")
    
    # ===== CLICK.UZ INTEGRATION =====
    async def click_create_invoice(self, user_id: int, amount: float, description: str = "") -> Dict:
        """Create Click.uz invoice"""
        try:
            invoice_id = f"inv_{user_id}_{int(datetime.now().timestamp())}"
            
            # Create signature
            data = f"{self.click_merchant_id}{self.click_service_id}{invoice_id}{amount}"
            signature = hashlib.md5(data.encode()).hexdigest()
            
            return {
                "status": "success",
                "invoice_id": invoice_id,
                "merchant_id": self.click_merchant_id,
                "service_id": self.click_service_id,
                "amount": amount,
                "signature": signature,
                "return_url": f"https://akadem.uz/payment/callback/click/{invoice_id}",
                "description": description or "Akadem Bot - Document Generation",
                "merchant_prepare_id": invoice_id
            }
        except Exception as e:
            logger.error(f"Click invoice creation error: {e}")
            return {"status": "error", "message": str(e)}
    
    async def click_verify_payment(self, click_trans_id: str, merchant_trans_id: str, 
                                  signature: str, amount: float) -> bool:
        """Verify Click.uz payment"""
        try:
            # Verify signature
            data = f"{click_trans_id}{self.click_merchant_id}{merchant_trans_id}{amount}"
            expected_signature = hashlib.md5(data.encode()).hexdigest()
            
            if signature == expected_signature:
                logger.info(f"Click payment verified: {click_trans_id}")
                return True
            else:
                logger.warning(f"Click signature mismatch: {click_trans_id}")
                return False
        except Exception as e:
            logger.error(f"Click verification error: {e}")
            return False
    
    # ===== PAYME INTEGRATION =====
    async def payme_create_invoice(self, user_id: int, amount: float, 
                                  description: str = "") -> Dict:
        """Create Payme invoice"""
        try:
            order_id = f"ord_{user_id}_{int(datetime.now().timestamp())}"
            
            # Convert amount to tiyn (1 UZS = 100 tiyn)
            amount_tiyn = int(amount * 100)
            
            payload = {
                "jsonrpc": "2.0",
                "id": 1,
                "method": "GeneratePaymentForm",
                "params": {
                    "amount": amount_tiyn,
                    "account": {"user_id": user_id},
                    "description": description or "Akadem Bot Document",
                    "returnUrl": f"https://akadem.uz/payment/callback/payme"
                }
            }
            
            async with aiohttp.ClientSession() as session:
                headers = {
                    "Authorization": f"Basic {self.payme_api_key}",
                    "Content-Type": "application/json"
                }
                
                async with session.post(
                    "https://checkout.paycom.uz/api/",
                    json=payload,
                    headers=headers
                ) as resp:
                    if resp.status == 200:
                        result = await resp.json()
                        if "result" in result:
                            return {
                                "status": "success",
                                "invoice_id": order_id,
                                "url": result["result"]["url"],
                                "amount": amount,
                                "order_id": order_id
                            }
            
            return {"status": "error", "message": "Payme API error"}
        except Exception as e:
            logger.error(f"Payme invoice creation error: {e}")
            return {"status": "error", "message": str(e)}
    
    async def payme_check_payment(self, transaction_id: str) -> Dict:
        """Check Payme payment status"""
        try:
            payload = {
                "jsonrpc": "2.0",
                "id": 1,
                "method": "CheckTransaction",
                "params": {
                    "id": transaction_id
                }
            }
            
            async with aiohttp.ClientSession() as session:
                headers = {
                    "Authorization": f"Basic {self.payme_api_key}",
                    "Content-Type": "application/json"
                }
                
                async with session.post(
                    "https://checkout.paycom.uz/api/",
                    json=payload,
                    headers=headers
                ) as resp:
                    if resp.status == 200:
                        result = await resp.json()
                        if "result" in result:
                            return {
                                "status": "success",
                                "transaction": result["result"]
                            }
            
            return {"status": "error"}
        except Exception as e:
            logger.error(f"Payme check error: {e}")
            return {"status": "error"}
    
    # ===== TELEGRAM STARS INTEGRATION =====
    async def telegram_create_invoice(self, user_id: int, amount: float, 
                                     title: str, description: str) -> Dict:
        """Create Telegram Stars invoice"""
        try:
            # Telegram Stars: 1 Star ≈ 0.013 USD
            # Convert UZS to approx Stars
            # For simplicity: 2000 UZS ≈ 2 Stars
            stars = max(1, int(amount / 1000))
            
            return {
                "status": "success",
                "amount": stars,
                "title": title,
                "description": description,
                "currency": "XTR",  # Telegram currency code
                "payload": f"akadem_bot_{user_id}_{int(datetime.now().timestamp())}"
            }
        except Exception as e:
            logger.error(f"Telegram invoice creation error: {e}")
            return {"status": "error", "message": str(e)}
    
    # ===== CARD PAYMENT (Direct) =====
    async def card_create_invoice(self, user_id: int, amount: float) -> Dict:
        """Create card payment invoice (manual verification)"""
        try:
            invoice_id = f"card_{user_id}_{int(datetime.now().timestamp())}"
            
            return {
                "status": "success",
                "invoice_id": invoice_id,
                "amount": amount,
                "type": "manual_card_transfer",
                "instructions": {
                    "bank_name": "Universal Bank",
                    "card_number": "XXXX XXXX XXXX 4242",
                    "cardholder": "AKADEM BOT",
                    "amount": amount,
                    "reference": invoice_id,
                    "description": "Payment for Akadem Bot Premium Subscription"
                },
                "verification_required": True,
                "estimated_time": "30 minutes"
            }
        except Exception as e:
            logger.error(f"Card invoice creation error: {e}")
            return {"status": "error", "message": str(e)}
    
    # ===== COMMON METHODS =====
    async def create_invoice(self, gateway: str, user_id: int, amount: float, 
                           **kwargs) -> Dict:
        """Unified invoice creation"""
        
        if gateway == "click":
            return await self.click_create_invoice(user_id, amount, 
                                                  kwargs.get("description", ""))
        elif gateway == "payme":
            return await self.payme_create_invoice(user_id, amount,
                                                  kwargs.get("description", ""))
        elif gateway == "telegram":
            return await self.telegram_create_invoice(
                user_id, amount,
                kwargs.get("title", "Akadem Bot Subscription"),
                kwargs.get("description", "Premium subscription for document generation")
            )
        elif gateway == "card":
            return await self.card_create_invoice(user_id, amount)
        else:
            return {"status": "error", "message": f"Unknown gateway: {gateway}"}
    
    async def verify_payment(self, gateway: str, **kwargs) -> bool:
        """Unified payment verification"""
        
        if gateway == "click":
            return await self.click_verify_payment(
                kwargs.get("click_trans_id"),
                kwargs.get("merchant_trans_id"),
                kwargs.get("signature"),
                kwargs.get("amount")
            )
        elif gateway == "payme":
            result = await self.payme_check_payment(kwargs.get("transaction_id"))
            return result.get("status") == "success"
        elif gateway in ["telegram", "card"]:
            # Manual verification required
            return kwargs.get("verified", False)
        else:
            return False
    
    def get_payment_url(self, gateway: str, invoice: Dict) -> Optional[str]:
        """Get payment URL for gateway"""
        
        if gateway == "click":
            return f"https://my.click.uz/services/pay?service_id={invoice.get('service_id')}&merchant_id={invoice.get('merchant_id')}&amount={invoice.get('amount')}&transaction_param={invoice.get('invoice_id')}"
        
        elif gateway == "payme":
            return invoice.get("url")
        
        elif gateway == "telegram":
            # Telegram Stars invoice handled by sendInvoice method
            return None
        
        elif gateway == "card":
            return None
        
        return None
    
    def format_amount(self, amount: float, gateway: str) -> str:
        """Format amount for gateway"""
        
        if gateway == "payme":
            return str(int(amount * 100))  # Convert to tiyn
        elif gateway == "telegram":
            return str(max(1, int(amount / 1000)))  # Convert to Stars
        else:
            return str(amount)
    
    def get_supported_currencies(self, gateway: str) -> list:
        """Get supported currencies for gateway"""
        
        currencies = {
            "click": ["UZS"],
            "payme": ["UZS"],
            "telegram": ["XTR"],  # Telegram Stars
            "card": ["UZS"]
        }
        
        return currencies.get(gateway, [])


class PaymentValidator:
    """Validate payment data"""
    
    @staticmethod
    def validate_amount(amount: float) -> bool:
        """Validate payment amount"""
        return isinstance(amount, (int, float)) and amount > 0
    
    @staticmethod
    def validate_user_id(user_id: int) -> bool:
        """Validate user ID"""
        return isinstance(user_id, int) and user_id > 0
    
    @staticmethod
    def validate_gateway(gateway: str) -> bool:
        """Validate payment gateway"""
        valid_gateways = ["click", "payme", "telegram", "card"]
        return gateway in valid_gateways
    
    @staticmethod
    def validate_invoice(invoice: Dict) -> bool:
        """Validate invoice data"""
        required_fields = ["status", "invoice_id", "amount"]
        return all(field in invoice for field in required_fields)


# Usage example
if __name__ == "__main__":
    import asyncio
    
    config = {
        "CLICK_MERCHANT_ID": "12345",
        "CLICK_SERVICE_ID": "67890",
        "PAYME_MERCHANT_ID": "merchant_id",
        "PAYME_API_KEY": "api_key"
    }
    
    service = PaymentService(config)
    
    async def test():
        # Test Click
        invoice = await service.create_invoice("click", 123, 2000)
        print(f"Click Invoice: {invoice}")
        
        # Test Card
        invoice = await service.create_invoice("card", 123, 2000)
        print(f"Card Invoice: {invoice}")
    
    asyncio.run(test())

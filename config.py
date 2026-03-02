import os
from dotenv import load_dotenv

load_dotenv()

class Config:
    # PRODUÇÃO: obrigatório existir no ambiente (sem fallback)
    SECRET_KEY = os.environ["SECRET_KEY"]

    # Mongo
    MONGO_URI = os.getenv("MONGO_URI") or "mongodb://localhost:27017/reinf_prod"

    # Define se está em produção
    IS_PRODUCTION = os.getenv("FLASK_ENV", "production").lower() == "production"

    # Cookies / sessão (hardening ativado apenas em produção)
    SESSION_COOKIE_SECURE = IS_PRODUCTION
    SESSION_COOKIE_HTTPONLY = True
    SESSION_COOKIE_SAMESITE = "Lax"
    
    # "remember me" do Flask-Login (hardening)
    REMEMBER_COOKIE_SECURE = IS_PRODUCTION
    REMEMBER_COOKIE_HTTPONLY = True
    REMEMBER_COOKIE_SAMESITE = "Lax"

    # CSRF
    WTF_CSRF_SSL_STRICT = IS_PRODUCTION

    # Mail
    MAIL_SERVER = os.getenv("MAIL_SERVER")
    MAIL_PORT = int(os.getenv("MAIL_PORT") or 587)
    MAIL_USE_TLS = (os.getenv("MAIL_USE_TLS") or "True").lower() == "true"
    MAIL_USERNAME = os.getenv("MAIL_USERNAME")
    MAIL_PASSWORD = os.getenv("MAIL_PASSWORD")
    MAIL_DEFAULT_SENDER = os.getenv("MAIL_USERNAME")

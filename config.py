import os
from dotenv import load_dotenv

load_dotenv()

class Config:
    SECRET_KEY = os.getenv('SECRET_KEY') or 'dev_key_padrao'

    # Conexão MongoDB Local - Banco: reinf_prod
    MONGO_URI = os.getenv('MONGO_URI') or 'mongodb://localhost:27017/reinf_prod'

    # Configurações do Outlook
    MAIL_SERVER = os.getenv('MAIL_SERVER')
    MAIL_PORT = int(os.getenv('MAIL_PORT') or 587)
    MAIL_USE_TLS = os.getenv('MAIL_USE_TLS') == 'True'
    MAIL_USERNAME = os.getenv('MAIL_USERNAME')
    MAIL_PASSWORD = os.getenv('MAIL_PASSWORD')

    # O Outlook EXIGE que o remetente seja o mesmo do login
    MAIL_DEFAULT_SENDER = os.getenv('MAIL_USERNAME')
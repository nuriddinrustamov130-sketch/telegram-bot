import os
from pathlib import Path
from dotenv import load_dotenv

env_path = Path(__file__).parent / ".env"
load_dotenv(dotenv_path=env_path)

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "8853034375:AAG0xx8Sy1sXSEPTyE2bxk_dNtJa3o4lBQs")
ADMIN_USER_ID = int(os.getenv("ADMIN_USER_ID", "8525754806"))
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-4o")

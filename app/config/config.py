import os
import json
from dotenv import load_dotenv

load_dotenv()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

class Settings:
    TENANT_ID = os.getenv("TENANT_ID")
    CLIENT_ID = os.getenv("CLIENT_ID")
    CLIENT_SECRET = os.getenv("CLIENT_SECRET")
    DRIVE_ID = os.getenv("DRIVE_ID")
    TEMPLATE_PATH_ON_SP = os.getenv("TEMPLATE_PATH_ON_SP", "/Templates")
    OUTPUT_PATH_ON_SP = os.getenv("OUTPUT_PATH_ON_SP", "/Output")
    SITE_URL = os.getenv("SITE_URL")
    SITE_ID = os.getenv("SITE_ID")
    CONFIG_PATH = os.path.join(BASE_DIR, "config.json")

def get_field_config():
    with open(Settings.CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)

settings = Settings()
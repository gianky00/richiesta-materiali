"""
Gestore dei Segreti e Credenziali
Carica le chiavi dal file .env e le gestisce in modo offuscato.
"""

import os
import base64
from dotenv import load_dotenv

# Carica variabili dal file .env
load_dotenv()

class SecretsManager:
    """
    Gestisce il recupero di chiavi API e altri segreti.
    """
    
    @staticmethod
    def get_github_token():
        """Recupera il token GitHub dal file .env"""
        token = os.getenv("GITHUB_TOKEN", "")
        if not token:
            return None
        return token

    @staticmethod
    def get_obfuscated_token():
        """Restituisce il token offuscato in Base64 (semplice offuscamento)"""
        token = SecretsManager.get_github_token()
        if not token:
            return ""
        return base64.b64encode(token.encode()).decode()

    @staticmethod
    def decode_token(obfuscated_token):
        """Decodifica un token precedentemente offuscato"""
        if not obfuscated_token:
            return ""
        try:
            return base64.b64decode(obfuscated_token.encode()).decode()
        except:
            return ""

# Istanza globale
secrets = SecretsManager()

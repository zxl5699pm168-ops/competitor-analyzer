import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app

# Vercel serverless WSGI handler
def handler(environ, start_response):
    return app(environ, start_response)

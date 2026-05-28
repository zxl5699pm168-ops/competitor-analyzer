import sys
import os
import traceback

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    from app import app as _flask_app
    handler = _flask_app
except Exception as e:
    error_msg = f"Import Error: {str(e)}\\n\\n{traceback.format_exc()}"

    def handler(environ, start_response):
        status = '500 Internal Server Error'
        headers = [('Content-Type', 'text/plain; charset=utf-8')]
        start_response(status, headers)
        return [error_msg.encode('utf-8')]

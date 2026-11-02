import os

# Render automatically provides the PORT environment variable (usually 10000)
port = os.environ.get('PORT', '10000')
bind = f"0.0.0.0:{port}"
workers = 2
timeout = 120

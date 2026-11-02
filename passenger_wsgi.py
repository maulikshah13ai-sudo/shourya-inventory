"""
WSGI entry point for cPanel / CloudLinux Passenger hosting.
This allows running the app on shared Linux hosting with cPanel 'Setup Python App'.
"""

import sys
import os

# Add application directory to the Python path
sys.path.insert(0, os.path.dirname(__file__))

from app import app as application

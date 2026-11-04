"""
Production Multi-Threaded Server for Bags & Footwear Inventory System.
Powered by Waitress WSGI Server for high stability, multi-user concurrency,
and zero database locking errors (SQLite WAL mode).
"""

import sys
import socket
from waitress import serve
from app import app, get_local_ip

def main():
    import os
    local_ip = get_local_ip()
    port = int(os.environ.get('PORT', 5000))

    print("=" * 70)
    print("  BAGS & FOOTWEAR WHOLESALE + RETAIL INVENTORY SYSTEM (PRODUCTION)")
    print("=" * 70)
    print(f"  [STATUS] Production Multi-Threaded Engine is RUNNING!")
    print(f"  -> Open on THIS Computer:         http://localhost:{port}")
    print(f"  -> Open on Staff Mobile/Tablets:  http://{local_ip}:{port}")
    print("=" * 70)
    print("  Staff Roles & Default Accounts:")
    print("    • Store Owner (Admin):     admin    / admin123")
    print("    • Floor Salesman:          sales1   / sales123")
    print("    • Accounts / Cashier:      accounts / accounts123")
    print("    • Warehouse Incharge:      godown   / godown123")
    print("=" * 70)
    print("  Keep this window open while using the app in your store.")
    print("  Press Ctrl+C to stop the server safely.")
    print("=" * 70)

    try:
        # 12 concurrent worker threads to handle simultaneous salesmen, accountants & godown
        serve(app, host='0.0.0.0', port=port, threads=12, channel_timeout=60)
    except KeyboardInterrupt:
        print("\nServer stopped safely. Have a great day!")
        sys.exit(0)

if __name__ == '__main__':
    main()

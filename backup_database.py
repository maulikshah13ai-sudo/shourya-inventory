"""
Automated Hot Backup Script for Bags & Shoes Inventory.
Uses SQLite Online Backup API to safely create backups even while the system is actively in use!
"""

import os
import sqlite3
from datetime import datetime

def perform_backup():
    base_dir = os.path.dirname(__file__)
    db_file = os.path.join(base_dir, 'inventory.db')
    backup_dir = os.path.join(base_dir, 'backups')
    os.makedirs(backup_dir, exist_ok=True)

    timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    backup_file = os.path.join(backup_dir, f"inventory_backup_{timestamp}.db")

    print(f"Connecting to live database: {db_file}")
    source_conn = sqlite3.connect(db_file)
    dest_conn = sqlite3.connect(backup_file)

    try:
        # SQLite online backup API ensures consistent snapshot without locking or stopping the app
        with dest_conn:
            source_conn.backup(dest_conn)
        print(f"[SUCCESS] Database successfully backed up to:")
        print(f"  -> {backup_file}")
    except Exception as e:
        print(f"[ERROR] Backup failed: {str(e)}")
    finally:
        source_conn.close()
        dest_conn.close()

if __name__ == '__main__':
    perform_backup()

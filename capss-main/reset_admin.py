"""
SPARK — Admin Password Reset
-----------------------------
Run this script from the terminal when the admin is locked out.

Usage:
    python reset_admin.py
"""

import sqlite3
import hashlib
import os
import getpass

DB_PATH = os.path.join(os.path.dirname(__file__), 'instance', 'spark.db')


def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()


def main():
    print("=" * 45)
    print("  SPARK — Admin Password Reset")
    print("=" * 45)

    # Check DB exists
    if not os.path.exists(DB_PATH):
        print(f"\n[ERROR] Database not found at: {DB_PATH}")
        print("Make sure you run this script from the SPARK project folder.")
        return

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row

    # List all admin accounts
    admins = conn.execute(
        "SELECT id, full_name, email FROM users WHERE role='admin' ORDER BY id"
    ).fetchall()

    if not admins:
        print("\n[ERROR] No admin accounts found in the database.")
        conn.close()
        return

    print(f"\nFound {len(admins)} admin account(s):\n")
    for i, a in enumerate(admins):
        print(f"  [{i + 1}] {a['full_name']} — {a['email']}")

    # Select admin if more than one
    if len(admins) == 1:
        chosen = admins[0]
    else:
        print()
        while True:
            try:
                choice = int(input("Select account number: "))
                if 1 <= choice <= len(admins):
                    chosen = admins[choice - 1]
                    break
                else:
                    print(f"    Please enter a number between 1 and {len(admins)}.")
            except ValueError:
                print("    Invalid input. Enter a number.")

    print(f"\nResetting password for: {chosen['full_name']} ({chosen['email']})")

    # Get new password
    while True:
        new_pass = getpass.getpass("Enter new password (min 6 characters): ")
        if len(new_pass) < 6:
            print("    Password too short. Must be at least 6 characters.")
            continue
        confirm = getpass.getpass("Confirm new password: ")
        if new_pass != confirm:
            print("    Passwords do not match. Try again.")
            continue
        break

    # Update DB
    conn.execute(
        "UPDATE users SET password=? WHERE id=?",
        (hash_password(new_pass), chosen['id'])
    )
    conn.commit()
    conn.close()

    print(f"\n[OK] Password reset successfully for {chosen['full_name']}.")
    print("     You can now log in with the new password.\n")


if __name__ == '__main__':
    main()

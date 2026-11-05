#!/usr/bin/env python3
"""
Admin Password Hash Generator (SHA-256)
----------------------------------------
Run this script to generate the hash for your admin password.
Paste the output into js/config.js -> admin.passwordHash

Usage:
    python python/hash_gen.py
    python python/hash_gen.py "MyNewPassword"
"""
import hashlib, getpass, sys

def hash_password(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()

def main():
    print("=" * 60)
    print("  Admin Password Hash Generator (SHA-256)")
    print("=" * 60)
    print("Note: only the HASH is shown — your password is never saved.\n")

    if len(sys.argv) > 1:
        pwd = sys.argv[1]
    else:
        pwd = getpass.getpass("Enter new admin password: ")
        confirm = getpass.getpass("Confirm password: ")
        if pwd != confirm:
            print("\n[!] Passwords do not match. Aborting.")
            sys.exit(1)

    if len(pwd) < 6:
        print("\n[!] Password must be at least 6 characters. Aborting.")
        sys.exit(1)

    h = hash_password(pwd)
    print("\n" + "-" * 60)
    print("Your SHA-256 hash:")
    print(h)
    print("-" * 60)
    print("\nPaste this into  js/config.js  at:  admin.passwordHash")
    print("\nDefault demo hash (password = 'admin123'):")
    print("  240be518fabd2724ddb6f04eeb1da5967448d7e831c08c8fa822809f74c720a9")
    print()

if __name__ == "__main__":
    main()

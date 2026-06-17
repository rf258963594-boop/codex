from __future__ import annotations

import argparse
import getpass
import os

from db import connect, hash_password, init_db, now


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Reset or create an admin user.")
    parser.add_argument("--username", default=os.environ.get("DEFAULT_ADMIN_USERNAME", "admin"))
    parser.add_argument("--password", default=os.environ.get("RESET_ADMIN_PASSWORD"))
    parser.add_argument("--role", default="admin", choices=["admin", "staff"])
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    password = args.password or getpass.getpass("New password: ")
    if not password:
        raise SystemExit("Password cannot be empty.")

    init_db()
    with connect() as conn:
        row = conn.execute("SELECT id FROM users WHERE username = ?", (args.username,)).fetchone()
        if row:
            conn.execute(
                "UPDATE users SET password_hash = ?, role = ?, active = 1 WHERE id = ?",
                (hash_password(password), args.role, row["id"]),
            )
            action = "updated"
        else:
            conn.execute(
                """
                INSERT INTO users (username, password_hash, role, active, created_at)
                VALUES (?, ?, ?, 1, ?)
                """,
                (args.username, hash_password(password), args.role, now()),
            )
            action = "created"
    print(f"User '{args.username}' {action} and enabled with role '{args.role}'.")


if __name__ == "__main__":
    main()

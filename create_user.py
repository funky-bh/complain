#!/usr/bin/env python3
"""Add or update a user in config.yaml."""

import getpass
import os
import sys

import bcrypt
import yaml

CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config.yaml")


def load_config() -> dict:
    if not os.path.exists(CONFIG_PATH):
        return {"users": {}}
    with open(CONFIG_PATH, encoding="utf-8") as f:
        return yaml.safe_load(f) or {"users": {}}


def save_config(config: dict):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        yaml.dump(config, f, allow_unicode=True, default_flow_style=False)


def main():
    config = load_config()
    users = config.setdefault("users", {})

    print("=== 사용자 추가/수정 ===")
    username = input("아이디: ").strip()
    if not username:
        sys.exit("아이디를 입력해주세요.")

    display_name = input("이름 (표시용): ").strip() or username

    while True:
        password = getpass.getpass("비밀번호: ")
        confirm = getpass.getpass("비밀번호 확인: ")
        if password == confirm:
            break
        print("비밀번호가 일치하지 않습니다. 다시 입력해주세요.")

    if len(password) < 8:
        sys.exit("비밀번호는 8자 이상이어야 합니다.")

    hashed = bcrypt.hashpw(password.encode(), bcrypt.gensalt()).decode()

    action = "수정" if username in users else "추가"
    users[username] = {"name": display_name, "password": hashed}
    save_config(config)

    print(f"\n사용자 '{username}' ({display_name}) {action} 완료.")


if __name__ == "__main__":
    main()

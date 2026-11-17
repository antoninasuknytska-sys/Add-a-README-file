"""Збирає всі непрочитані повідомлення з різних джерел в одному місці.

Використання:
    python unread_messages.py            # демо-дані
    python unread_messages.py --imap     # пошта через IMAP (див. змінні середовища нижче)

Для IMAP (наприклад, Gmail) задайте змінні середовища:
    IMAP_HOST=imap.gmail.com
    IMAP_USER=you@gmail.com
    IMAP_PASSWORD=<пароль застосунку>
"""

from __future__ import annotations

import argparse
import email
import imaplib
import os
import sys
from dataclasses import dataclass
from datetime import datetime, timezone
from email.header import decode_header, make_header
from email.utils import parsedate_to_datetime
from typing import Callable, Iterable, List

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


@dataclass
class Message:
    source: str      # звідки повідомлення (Gmail, Telegram, ...)
    sender: str
    subject: str     # тема або початок тексту
    received: datetime


# Джерело — будь-яка функція без аргументів, що повертає непрочитані повідомлення.
Source = Callable[[], Iterable[Message]]


def get_all_unread(sources: Iterable[Source]) -> List[Message]:
    """Головна функція: збирає непрочитані з усіх джерел, найновіші — першими.

    Якщо одне джерело недоступне, решта все одно показуються.
    """
    messages: List[Message] = []
    for source in sources:
        try:
            messages.extend(source())
        except Exception as exc:  # одне зламане джерело не має ламати все
            name = getattr(source, "__name__", repr(source))
            print(f"⚠️  Не вдалося отримати повідомлення з {name}: {exc}", file=sys.stderr)
    return sorted(messages, key=lambda m: m.received, reverse=True)


# ---------- Джерела ----------

def imap_source(host: str, user: str, password: str, folder: str = "INBOX",
                label: str = "Пошта") -> Source:
    """Непрочитані листи з будь-якої пошти з IMAP (Gmail, Ukr.net, Outlook...)."""

    def _decode(value: str | None) -> str:
        return str(make_header(decode_header(value))) if value else ""

    def fetch() -> List[Message]:
        result = []
        with imaplib.IMAP4_SSL(host) as conn:
            conn.login(user, password)
            conn.select(folder, readonly=True)  # readonly: листи лишаються непрочитаними
            _, data = conn.search(None, "UNSEEN")
            for num in data[0].split():
                _, msg_data = conn.fetch(num, "(BODY.PEEK[HEADER.FIELDS (FROM SUBJECT DATE)])")
                msg = email.message_from_bytes(msg_data[0][1])
                try:
                    received = parsedate_to_datetime(msg["Date"])
                except (TypeError, ValueError):
                    received = datetime.now(timezone.utc)
                if received.tzinfo is None:
                    received = received.replace(tzinfo=timezone.utc)
                result.append(Message(label, _decode(msg["From"]),
                                      _decode(msg["Subject"]) or "(без теми)", received))
        return result

    fetch.__name__ = f"{label} ({user})"
    return fetch


def demo_source() -> List[Message]:
    """Приклад даних, щоб побачити, як це виглядає без підключення акаунтів."""
    utc = timezone.utc
    return [
        Message("Gmail", "Оля <olya@example.com>", "Фото з фотосесії готові!",
                datetime(2026, 11, 17, 9, 30, tzinfo=utc)),
        Message("Telegram", "Мама", "Ти сьогодні прийдеш на вечерю?",
                datetime(2026, 11, 17, 11, 5, tzinfo=utc)),
        Message("Slack", "Андрій", "Глянь, будь ласка, новий макет",
                datetime(2026, 11, 16, 18, 42, tzinfo=utc)),
    ]


# ---------- Вивід ----------

def print_unread(messages: List[Message]) -> None:
    if not messages:
        print("🎉 Непрочитаних повідомлень немає!")
        return
    print(f"📬 Непрочитаних повідомлень: {len(messages)}\n")
    for m in messages:
        when = m.received.astimezone().strftime("%d.%m %H:%M")
        print(f"[{m.source}] {when}  {m.sender}\n    {m.subject}\n")


def main() -> None:
    parser = argparse.ArgumentParser(description="Усі непрочитані повідомлення в одному місці")
    parser.add_argument("--imap", action="store_true", help="читати пошту через IMAP")
    args = parser.parse_args()

    sources: List[Source] = []
    if args.imap:
        try:
            sources.append(imap_source(os.environ["IMAP_HOST"], os.environ["IMAP_USER"],
                                       os.environ["IMAP_PASSWORD"]))
        except KeyError as missing:
            sys.exit(f"Не задано змінну середовища {missing}")
    else:
        sources.append(demo_source)

    print_unread(get_all_unread(sources))


if __name__ == "__main__":
    main()

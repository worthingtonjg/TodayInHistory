from __future__ import annotations

import argparse
import base64
import json
import html
import mailbox
import hashlib
import re
from dataclasses import dataclass
from datetime import datetime, timezone
from email.message import Message
from email.utils import parsedate_to_datetime, parseaddr
from pathlib import Path
from typing import Any, Iterable, Optional


SENDER_TO_KIND = {
    "news@newsletter.theretrospectdaily.com": "retrospect",
    "news@newsletter.backthenhistory.com": "backThen",
}


MONTHS = {
    "january": 1,
    "february": 2,
    "march": 3,
    "april": 4,
    "may": 5,
    "june": 6,
    "july": 7,
    "august": 8,
    "september": 9,
    "october": 10,
    "november": 11,
    "december": 12,
}


@dataclass
class ParsedMessage:
    message_id: str
    sender: str
    kind: str
    subject: str
    received_date: datetime
    historical_date: Optional[datetime]
    event_title: str
    body_text: str
    body_html: str
    raw_source: str

    @property
    def output_date(self) -> datetime:
        return self.historical_date or self.received_date

    @property
    def filename(self) -> str:
        return f"{self.output_date.strftime('%Y-%m-%d')}-{sanitize_filename(self.event_title)}"

    @property
    def folder_name(self) -> str:
        if self.kind == "retrospect":
            return "The Retrospect"
        if self.kind == "backThen":
            return "Back Then History"
        return "Unknown"


def parse_sender_address(from_value: str) -> str:
    _, email_address = parseaddr(from_value or "")
    return (email_address or from_value or "").strip().lower()


def sanitize_filename(value: str) -> str:
    text = re.sub(r"\s+", " ", str(value or "")).strip()
    text = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", text)
    text = text.rstrip(" .")
    return text or "Untitled"


def decode_html_entities(text: str) -> str:
    return html.unescape(text or "")


def html_to_plain_text(html_text: str) -> str:
    source = str(html_text or "")
    source = re.sub(r"<style[\s\S]*?</style>", "", source, flags=re.I)
    source = re.sub(r"<script[\s\S]*?</script>", "", source, flags=re.I)
    source = re.sub(r"<br\s*/?>", "\n", source, flags=re.I)
    source = re.sub(r"</p>", "\n\n", source, flags=re.I)
    source = re.sub(r"</div>", "\n", source, flags=re.I)
    source = re.sub(r"</tr>", "\n", source, flags=re.I)
    source = re.sub(r"</li>", "\n", source, flags=re.I)
    source = re.sub(r"<li[^>]*>", "- ", source, flags=re.I)
    source = re.sub(r"<[^>]+>", "", source)
    source = decode_html_entities(source)
    source = source.replace("\r", "")
    source = re.sub(r"\n{3,}", "\n\n", source)
    return source.strip()


def normalize_lines(text: str) -> list[str]:
    normalized = (
        str(text or "")
        .replace("\r", "")
        .replace("&nbsp;", " ")
        .replace("&#160;", " ")
        .replace("\u00a0", " ")
    )
    lines = []
    for raw_line in re.split(r"\n+", normalized):
        line = re.sub(r"\s+", " ", raw_line).strip()
        if line:
            lines.append(line)
    return lines


def is_content_line(line: str) -> bool:
    return bool(str(line or "").strip()) and not re.fullmatch(r"(&nbsp;|&#160;)+", str(line or "").strip(), flags=re.I)


def is_date_line(line: str) -> bool:
    value = str(line or "").strip()
    return bool(
        re.fullmatch(r"(Sunday|Monday|Tuesday|Wednesday|Thursday|Friday|Saturday),\s+[A-Z][a-z]+\s+\d{1,2},\s+\d{4}", value, flags=re.I)
        or re.fullmatch(r"On\s+[A-Z][a-z]+\s+\d{1,2},\s+\d{4}", value, flags=re.I)
        or re.fullmatch(r"\d{4}-\d{2}-\d{2}", value)
    )


def find_line_index(lines: list[str], predicate) -> int:
    for index, line in enumerate(lines):
        if predicate(line, index):
            return index
    return -1


def parse_event_title(subject: str) -> str:
    title = re.sub(r"\s+", " ", str(subject or "")).strip()
    for prefix in (r"^Today in History:\s*", r"^The Bizarre History of\s*"):
        title = re.sub(prefix, "", title, flags=re.I)
    title = re.sub(r"[.?!]+$", "", title).strip()
    return title or "Untitled"


def parse_historical_date(text: str) -> Optional[datetime]:
    patterns = (
        re.compile(r"On\s+([A-Z][a-z]+)\s+(\d{1,2}),\s+(\d{4})", re.I),
        re.compile(r"\b([A-Z][a-z]+)\s+(\d{1,2}),\s+(\d{4})\b"),
    )
    for pattern in patterns:
        match = pattern.search(str(text or ""))
        if not match:
            continue
        month = MONTHS.get(match.group(1).lower(), 0)
        day = int(match.group(2))
        year = int(match.group(3))
        if month and day and year:
            try:
                return datetime(year, month, day, tzinfo=timezone.utc)
            except ValueError:
                return None
    return None


def extract_retrospect_content(body_text: str, subject: str) -> tuple[str, Optional[datetime], str]:
    lines = normalize_lines(body_text)
    start = find_line_index(
        lines,
        lambda line, _index: bool(line)
        and not re.match(r"^(Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday),\s", line, flags=re.I)
        and not re.match(r"^On This Day in History$", line, flags=re.I)
        and not re.match(r"^Logo$", line, flags=re.I),
    )
    footer = find_line_index(
        lines,
        lambda line, index: bool(re.match(r"^Learn More$", line, flags=re.I)) or (bool(re.match(r"^The Retrospect$", line, flags=re.I)) and index > start),
    )
    slice_end = footer if footer >= 0 else len(lines)
    content_lines = [
        line
        for line in lines[max(0, start):slice_end]
        if is_content_line(line)
        and not is_date_line(line)
        and not re.match(r"^Logo$", line, flags=re.I)
        and not re.match(r"^On This Day in History$", line, flags=re.I)
        and not re.match(r"^image source:", line, flags=re.I)
    ]
    title = content_lines[0] if content_lines else parse_event_title(subject)
    article_lines = content_lines[1:]
    article_text = "\n\n".join([title, *article_lines]).strip()
    return title, parse_historical_date(body_text), article_text


def extract_back_then_content(body_text: str, subject: str) -> tuple[str, Optional[datetime], str]:
    lines = normalize_lines(body_text)
    title_marker_index = find_line_index(
        lines,
        lambda line, _index: bool(re.match(r"^The history of$", line, flags=re.I) or re.match(r"^Today's Object$", line, flags=re.I)),
    )
    prompt_index = find_line_index(
        lines,
        lambda line, _index: bool(re.match(r"^Do you ever wonder who first thought of ", line, flags=re.I)),
    )
    if title_marker_index >= 0:
        title_index = title_marker_index + 1 if title_marker_index + 1 < len(lines) else -1
        body_start = title_marker_index + 1
        while body_start < len(lines) and (
            re.match(r"^Today's Object$", lines[body_start], flags=re.I)
            or re.match(r"^Do you ever wonder who first thought of ", lines[body_start], flags=re.I)
        ):
            body_start += 1
    elif prompt_index > 0:
        title_index = prompt_index - 1
        body_start = prompt_index + 1
    else:
        title_index = 0
        body_start = 0
    footer_start = find_line_index(
        lines,
        lambda line, index: bool(
            re.match(r"^Continue Reading$", line, flags=re.I)
            or re.match(r"^Missed a Few\? Explore Past Objects!$", line, flags=re.I)
            or re.match(r"^How would you rate today(?:'|â€™|Ã¢â‚¬â„¢|ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢)s newsletter\?$", line, flags=re.I)
            or re.match(r"^Logo$", line, flags=re.I)
            or (re.match(r"^Back Then History$", line, flags=re.I) and index > body_start)
        ),
    )
    slice_end = footer_start if footer_start >= 0 else len(lines)
    content_lines = [
        line
        for line in lines[max(0, body_start):slice_end]
        if is_content_line(line)
        and not is_date_line(line)
        and not re.match(r"^Follow us on our social channels!?$", line, flags=re.I)
        and not re.match(r"^Continue Reading$", line, flags=re.I)
        and not re.match(r"^Missed a Few\? Explore Past Objects!$", line, flags=re.I)
        and not re.match(r"^How would you rate today(?:'|â€™|Ã¢â‚¬â„¢|ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢)s newsletter\?$", line, flags=re.I)
        and not re.match(r"^Back Then History$", line, flags=re.I)
        and not re.match(r"^Logo$", line, flags=re.I)
        and not re.match(r"^The history of$", line, flags=re.I)
        and not re.match(r"^Today's Object$", line, flags=re.I)
        and not re.match(r"^Do you ever wonder who first thought of ", line, flags=re.I)
        and not re.match(r"^Download attachment$", line, flags=re.I)
    ]
    title = lines[title_index] if 0 <= title_index < len(lines) else parse_event_title(subject)
    if title.lower() == "the history of" or title.lower() == "today's object":
        title = parse_event_title(subject)
    article_lines = content_lines[1:]
    article_text = " ".join(article_lines)
    body = "\n\n".join([title, *article_lines]).strip()
    event_title = re.sub(r"^The history of\s+", "", title, flags=re.I).strip() or parse_event_title(subject)
    return event_title, parse_historical_date(article_text), body


def get_message_html(message: Message) -> str:
    if message.is_multipart():
        for part in message.walk():
            content_type = (part.get_content_type() or "").lower()
            disposition = (part.get_content_disposition() or "").lower()
            if content_type == "text/html" and disposition != "attachment":
                payload = part.get_payload(decode=True)
                if payload is not None:
                    charset = part.get_content_charset() or "utf-8"
                    return payload.decode(charset, errors="replace")
    else:
        if (message.get_content_type() or "").lower() == "text/html":
            payload = message.get_payload(decode=True)
            if payload is not None:
                charset = message.get_content_charset() or "utf-8"
                return payload.decode(charset, errors="replace")
    return ""


def get_message_text(message: Message) -> str:
    html_body = get_message_html(message)
    if html_body:
        return html_to_plain_text(html_body)
    if message.is_multipart():
        for part in message.walk():
            content_type = (part.get_content_type() or "").lower()
            disposition = (part.get_content_disposition() or "").lower()
            if content_type == "text/plain" and disposition != "attachment":
                payload = part.get_payload(decode=True)
                if payload is not None:
                    charset = part.get_content_charset() or "utf-8"
                    return (payload.decode(charset, errors="replace") or "").replace("\r", "")
    payload = message.get_payload(decode=True)
    if payload is not None:
        charset = message.get_content_charset() or "utf-8"
        return (payload.decode(charset, errors="replace") or "").replace("\r", "")
    return ""


def get_message_source(message: Message) -> str:
    return (message.as_string() or "").replace("\r", "")


def collect_headers(message: Message) -> list[dict[str, str]]:
    headers: list[dict[str, str]] = []
    for name, value in message.raw_items():
        headers.append({"name": str(name), "value": str(value)})
    return headers


def part_to_dict(part: Message) -> dict[str, Any]:
    content_type = (part.get_content_type() or "").lower()
    disposition = (part.get_content_disposition() or "").lower()
    filename = part.get_filename()
    content_id = part.get("Content-ID")
    transfer_encoding = part.get("Content-Transfer-Encoding")

    node: dict[str, Any] = {
        "content_type": content_type,
        "content_disposition": disposition or None,
        "filename": filename,
        "content_id": content_id,
        "content_transfer_encoding": transfer_encoding,
        "headers": collect_headers(part),
    }

    if part.is_multipart():
        node["parts"] = [part_to_dict(child) for child in part.iter_parts()]
        return node

    payload = part.get_payload(decode=True)
    if payload is None:
        node["payload"] = ""
        node["payload_base64"] = ""
        return node

    if content_type.startswith("text/"):
        charset = part.get_content_charset() or "utf-8"
        node["payload"] = payload.decode(charset, errors="replace")
        node["payload_base64"] = base64.b64encode(payload).decode("ascii")
    else:
        node["payload"] = None
        node["payload_base64"] = base64.b64encode(payload).decode("ascii")
    return node


def extract_message(message: Message) -> ParsedMessage:
    subject = str(message.get("Subject", "") or "").strip()
    sender = parse_sender_address(message.get("From", ""))
    kind = SENDER_TO_KIND.get(sender, "unknown")
    received_raw = message.get("Date", "")
    received_date = parsedate_to_datetime(received_raw) if received_raw else datetime.now(timezone.utc)
    if received_date.tzinfo is None:
        received_date = received_date.replace(tzinfo=timezone.utc)

    body_html = get_message_html(message)
    body_text = get_message_text(message)
    raw_source = get_message_source(message)

    if kind == "retrospect":
        event_title, historical_date, parsed_body = extract_retrospect_content(body_text, subject)
    elif kind == "backThen":
        event_title, historical_date, parsed_body = extract_back_then_content(body_text, subject)
    else:
        event_title = parse_event_title(subject)
        historical_date = parse_historical_date(body_text)
        parsed_body = body_text.strip()

    return ParsedMessage(
        message_id=str(message.get("Message-ID", "") or ""),
        sender=sender,
        kind=kind,
        subject=subject,
        received_date=received_date,
        historical_date=historical_date,
        event_title=event_title,
        body_text=parsed_body,
        body_html=body_html,
        raw_source=raw_source,
    )


def iter_parsed_messages(mbox_path: Path) -> Iterable[ParsedMessage]:
    mbox = mailbox.mbox(str(mbox_path))
    for message in mbox:
        yield extract_message(message)


def message_stable_id(parsed: ParsedMessage) -> str:
    source = "\n".join(
        [
            parsed.message_id,
            parsed.sender,
            parsed.subject,
            parsed.received_date.isoformat(),
            parsed.event_title,
        ]
    )
    return hashlib.sha1(source.encode("utf-8")).hexdigest()[:10]


def parsed_message_to_jsonable(parsed: ParsedMessage, message: Message, source_file: str) -> dict[str, Any]:
    return {
        "source_file": source_file,
        "message_id": parsed.message_id,
        "sender": parsed.sender,
        "kind": parsed.kind,
        "folder_name": parsed.folder_name,
        "subject": parsed.subject,
        "received_date": parsed.received_date.isoformat(),
        "historical_date": parsed.historical_date.isoformat() if parsed.historical_date else None,
        "event_title": parsed.event_title,
        "body_text": parsed.body_text,
        "body_html": parsed.body_html,
        "raw_source": parsed.raw_source,
        "headers": collect_headers(message),
        "mime": part_to_dict(message),
    }


def write_json_file(output_dir: Path, kind: str, records: list[dict[str, Any]]) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    path = output_dir / f"{'The Retrospect' if kind == 'retrospect' else 'Back Then History'}.json"
    payload = {
        "kind": kind,
        "folder_name": "The Retrospect" if kind == "retrospect" else "Back Then History",
        "source_directory": str(output_dir),
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "message_count": len(records),
        "messages": records,
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def main() -> int:
    parser = argparse.ArgumentParser(description="Parse Today In History Gmail exports from an mbox file.")
    parser.add_argument(
        "--mbox",
        default="input/School.mbox",
        help="Path to the exported Gmail mbox file.",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=10,
        help="Number of messages to print for inspection.",
    )
    parser.add_argument(
        "--output-dir",
        default="output",
        help="Folder where the grouped JSON output will be written.",
    )
    args = parser.parse_args()

    mbox_path = Path(args.mbox)
    if not mbox_path.exists():
        raise SystemExit(f"mbox file not found: {mbox_path}")

    output_dir = Path(args.output_dir)
    json_records: dict[str, list[dict[str, Any]]] = {
        "backThen": [],
        "retrospect": [],
    }

    mbox = mailbox.mbox(str(mbox_path))
    for index, message in enumerate(mbox):
        parsed = extract_message(message)
        if parsed.kind in json_records:
            json_records[parsed.kind].append(
                parsed_message_to_jsonable(parsed, message, str(mbox_path))
            )
        if index < args.limit:
            print(
                {
                    "index": index,
                    "sender": parsed.sender,
                    "kind": parsed.kind,
                    "subject": parsed.subject,
                    "filename": parsed.filename,
                    "historical_date": parsed.historical_date.isoformat() if parsed.historical_date else None,
                    "received_date": parsed.received_date.isoformat(),
                    "event_title": parsed.event_title,
                    "body_preview": parsed.body_text[:240],
                }
            )

    for kind in ("backThen", "retrospect"):
        path = write_json_file(output_dir, kind, json_records[kind])
        print(f"Wrote {len(json_records[kind])} messages to {path}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())

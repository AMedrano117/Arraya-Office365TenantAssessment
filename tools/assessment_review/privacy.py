"""Redaction of script-derived text and privacy scanning of generated outputs."""
from __future__ import annotations

import re

from .util import ascii_clean

EMAIL = re.compile(r"\b[A-Za-z0-9._%+'-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b")
GUID = re.compile(r"\b[0-9a-fA-F]{8}-(?:[0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}\b")
IPV4 = re.compile(r"\b(?:\d{1,3}\.){3}\d{1,3}\b")
IPV6 = re.compile(r"\b(?:[0-9a-fA-F]{1,4}:){3,7}[0-9a-fA-F]{0,4}\b")
URL = re.compile(r"https?://[^\s)\]\"'<>]+")
PHONE = re.compile(r"(?<![\w-])(?:\+\d[\d ()-]{8,}\d|\(\d{3}\)\s*\d{3}[- ]\d{4}|\b\d{3}[-.]\d{3}[-.]\d{4}\b)")
JWT = re.compile(r"\beyJ[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}")
SECRET = re.compile(r"(?i)(client[_ ]?secret|password|api[_ ]?key|bearer)\s*[:=]\s*\S+|\b[A-Za-z0-9_~.-]{3}8Q~[A-Za-z0-9_~.-]{20,}")
GRAPH_PAYLOAD = re.compile(r"@odata\.|\"userPrincipalName\"\s*:|graph\.microsoft\.com|api\.security\.microsoft\.com")
SHAREPOINT_HOST = re.compile(r"\b[\w-]+\.sharepoint\.com\b|\b[\w-]+\.onmicrosoft\.com\b", re.I)
QUOTED_MULTIWORD = re.compile(r"'([^'\n]{3,80}\s[^'\n]{1,80})'")
ALLOWED_URL = re.compile(r"^https://learn\.microsoft\.com/")


class Sanitizer:
    """Removes identifying values from script text before it reaches any output."""

    def __init__(self, redact_terms):
        self.terms = [t for t in redact_terms if t]
        self._pattern = None
        if self.terms:
            escaped = [re.escape(t) for t in self.terms]
            self._pattern = re.compile(r"(?<![\w])(" + "|".join(escaped) + r")(?![\w])", re.I)

    def __call__(self, value):
        if value is None:
            return None
        text = str(value)
        text = re.sub(r"\{[^{}]*(@odata|\"[A-Za-z]+\"\s*:)[^{}]*\}", "[payload removed]", text)
        text = re.sub(r"@odata\.\w+", "[payload removed]", text)
        text = URL.sub(lambda m: m.group(0) if ALLOWED_URL.match(m.group(0)) else "[link removed]", text)
        text = EMAIL.sub("[account removed]", text)
        text = GUID.sub("[identifier removed]", text)
        text = SHAREPOINT_HOST.sub("[tenant address removed]", text)
        text = IPV4.sub("[address removed]", text)
        text = IPV6.sub("[address removed]", text)
        text = JWT.sub("[token removed]", text)
        text = SECRET.sub("[secret removed]", text)
        if self._pattern is not None:
            text = self._pattern.sub("[named item]", text)
        text = QUOTED_MULTIWORD.sub("'[named item]'", text)
        text = PHONE.sub("[number removed]", text)
        return ascii_clean(text)


def scan_secrets(text):
    """Checks for list sheets, which name tenant objects by design but must never carry credentials or payloads."""
    return [f"{name}: {match.group(0)[:40]}" for name, pattern in
            (("token", JWT), ("secret", SECRET), ("Graph payload", re.compile(r"@odata\.|\"userPrincipalName\"\s*:")))
            for match in pattern.finditer(text)]


def scan_text(text, scan_terms, allow_terms=()):
    """Return a list of privacy violations found in rendered output text."""
    problems = []
    for name, pattern in (("email or UPN", EMAIL), ("GUID", GUID), ("IPv4 address", IPV4),
                          ("IPv6 address", IPV6),
                          ("phone number", PHONE), ("token", JWT), ("secret", SECRET),
                          ("Graph payload or API host", GRAPH_PAYLOAD), ("tenant host name", SHAREPOINT_HOST)):
        for match in pattern.finditer(text):
            problems.append(f"{name}: {match.group(0)[:40]}")
    for match in URL.finditer(text):
        if not ALLOWED_URL.match(match.group(0)):
            problems.append(f"non-Microsoft Learn URL: {match.group(0)[:60]}")
    lowered = text.lower()
    allow = {a.lower() for a in allow_terms}
    for term in scan_terms:
        t = term.lower()
        if t in allow:
            continue
        idx = lowered.find(t)
        while idx != -1:
            before = lowered[idx - 1] if idx > 0 else " "
            after = lowered[idx + len(t)] if idx + len(t) < len(lowered) else " "
            if not (before.isalnum() or after.isalnum()):
                problems.append(f"private value: {term[:3]}... ({len(term)} chars)")
                break
            idx = lowered.find(t, idx + 1)
    return problems

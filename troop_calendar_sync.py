##    python -m venv .venv

##    .\.venv\Scripts\Activate.ps1

##    pip install playwright beautifulsoup4 python-dateutil python-dotenv

##    python -m playwright install


# record by:   python -m playwright codegen https://www.troopwebhost.org/Troop1Mendon/Index.htm

# =========================
# IMPORTS
# =========================
import html
import os
import re
import hashlib
from datetime import datetime, timedelta, timezone

from dotenv import load_dotenv
from playwright.sync_api import sync_playwright
from bs4 import BeautifulSoup
from dateutil import parser as dtparser


# =========================
# ENV SETUP
# =========================
load_dotenv()

BASE = os.getenv("TWH_BASE", "").rstrip("/")
USERNAME = os.getenv("TWH_USERNAME")
PASSWORD = os.getenv("TWH_PASSWORD")

OUTPUT_ICS = r"C:\Users\MatthewCollins\OneDrive\Scouts\Troop_1\CalendarSync\troop_calendar.ics"
LOCAL_TZ = timezone(timedelta(hours=-5))

# =========================
# FUNCTIONS
# =========================


def extract_event_links_from_form_list(html: str):
    """
    Parse the FormList page and return a list of absolute URLs to FormDetail pages.

    TroopWebHost often embeds links as:
      - <a href="FormDetail.aspx?...">
      - javascript:LinkTo('FormDetail.aspx?...','')
      - onclick="LinkTo('FormDetail.aspx?...','')"
    """
    soup = BeautifulSoup(html, "html.parser")

    found = []

    def add_href(href: str, text: str = ""):
        if not href:
            return
        href = href.strip()

        # Pull FormDetail URL out of javascript:LinkTo('...')
        m = re.search(r"FormDetail\.aspx\?[^'\"()]+", href, re.I)
        if m:
            href = m.group(0)

        if "FormDetail.aspx" not in href or "ID=" not in href:
            return

        # Make absolute
        if href.startswith("http"):
            url = href
        else:
            url = "https://www.troopwebhost.org/" + href.lstrip("/")

        title = " ".join((text or "").split())
        found.append({"url": url, "title": title})

    # 1) Normal anchors
    for a in soup.select("a"):
        href = a.get("href", "") or ""
        txt = a.get_text(" ", strip=True)
        add_href(href, txt)

        # Some pages use onclick with LinkTo(...)
        onclick = a.get("onclick", "") or ""
        add_href(onclick, txt)

    # 2) Any element with onclick LinkTo(...)
    for el in soup.select("[onclick]"):
        onclick = el.get("onclick", "") or ""
        txt = el.get_text(" ", strip=True)
        add_href(onclick, txt)

    # 3) As a last resort: regex scan entire HTML for FormDetail.aspx?...ID=...
    for m in re.finditer(r"FormDetail\.aspx\?[^\"'<> ]*ID=\d+[^\"'<> ]*", html, re.I):
        add_href(m.group(0), "")

    # Deduplicate
    seen = set()
    out = []
    for x in found:
        if x["url"] not in seen:
            seen.add(x["url"])
            out.append(x)

    return out



def build_ics(events):
    """
    events: list of dicts with keys:
      - summary (str)
      - dtstart (datetime)
      - dtend (datetime)
      - description (str)
      - location (str)
      - url (str)
      - all_day (bool)
    """
    def fmt_dt(dt: datetime) -> str:
        # Use UTC for ICS timestamps unless all-day
        return dt.astimezone(timezone.utc).strftime("%Y%m%dT%H%M%SZ")

    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//TroopCalendarSync//EN",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
    ]

    now_utc = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")

    for ev in events:
        uid_src = (ev.get("url","") + "|" + ev.get("summary","")).encode("utf-8", errors="ignore")
        uid = hashlib.sha1(uid_src).hexdigest() + "@troopwebhost"

        lines.append("BEGIN:VEVENT")
        lines.append(f"UID:{uid}")
        lines.append(f"DTSTAMP:{now_utc}")

        if ev.get("all_day"):
            # All-day uses DATE (no time)
            ds = ev["dtstart"].date().strftime("%Y%m%d")
            de = ev["dtend"].date().strftime("%Y%m%d")
            lines.append(f"DTSTART;VALUE=DATE:{ds}")
            lines.append(f"DTEND;VALUE=DATE:{de}")
        else:
            lines.append(f"DTSTART:{fmt_dt(ev['dtstart'])}")
            lines.append(f"DTEND:{fmt_dt(ev['dtend'])}")

        summary = ev.get("summary", "").replace("\n", " ").strip()
        lines.append(f"SUMMARY:{summary}")

        loc = (ev.get("location") or "").replace("\n", " ").strip()
        if loc:
            lines.append(f"LOCATION:{loc}")

        desc_parts = []
        if ev.get("description"):
            desc_parts.append(ev["description"].strip())
        if ev.get("url"):
            desc_parts.append(ev["url"].strip())
        desc = "\\n\\n".join(desc_parts).replace("\n", "\\n")
        if desc:
            lines.append(f"DESCRIPTION:{desc}")

        lines.append("END:VEVENT")

    lines.append("END:VCALENDAR")
    return "\r\n".join(lines) + "\r\n"



# =========================
# MAIN
# =========================

def main():
    if not (BASE and USERNAME and PASSWORD):
        raise RuntimeError("Missing BASE/USERNAME/PASSWORD. Check your .env file.")

    def dump_frames(page, label):
        print(label)
        for fr in page.frames:
            print(" -", fr.url)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)  # set True after it works
        context = browser.new_context()
        page = context.new_page()

        # 1) Load the troop site (frameset)
        page.goto(f"{BASE}/Index.htm", wait_until="domcontentloaded")
        dump_frames(page, "FRAME URLS (initial):")

        # 2) Get content frame helper
        def get_content_frame(expected_url_substring: str | None = None):
            """
            Return the best candidate content frame.
            If expected_url_substring is provided, prefer a frame whose URL contains it.
            """
            last_urls = None
            for _ in range(60):  # up to ~30 seconds
                frames = page.frames
                urls = [fr.url for fr in frames]
                last_urls = urls

                # Prefer a frame matching expected URL substring
                if expected_url_substring:
                    for fr in frames:
                        if expected_url_substring.lower() in (fr.url or "").lower():
                            try:
                                fr.locator("body").count()
                                return fr
                            except:
                                pass

                # Otherwise: choose first non-main frame that has a body
                for fr in frames[1:]:
                    try:
                        fr.locator("body").count()
                        return fr
                    except:
                        pass

                page.wait_for_timeout(500)

            raise RuntimeError(f"Content frame never became available. Frames seen: {last_urls}")

        # ✅ 3) Acquire the content frame BEFORE using it
        content = get_content_frame("Redirect.htm")
        print("Using content frame URL:", content.url)

        # ✅ 4) Click Log On
        try:
            content.get_by_role("link", name=re.compile(r"Log\s*On", re.I)).click(timeout=8000)
        except:
            content.locator("a", has_text=re.compile(r"Log\s*On", re.I)).first.click(timeout=8000)

        page.wait_for_timeout(500)
        dump_frames(page, "FRAME URLS (after clicking Log On):")


        # 5) Force open menu and go to events list page
        content = get_content_frame()
        try:
            content.evaluate("togglemenu();")
        except:
            pass

        events_list_url = "https://www.troopwebhost.org/FormList.aspx?Menu_Item_ID=45936&Stack=1"
        print("Navigating to events list:", events_list_url)
        page.goto(events_list_url, wait_until="domcontentloaded", timeout=20000)
        page.wait_for_timeout(800)

        content = get_content_frame("FormList.aspx")
        html = content.content()
        with open("events_list.html", "w", encoding="utf-8") as f:
            f.write(html)

        links = [x for x in extract_event_links_from_form_list(html)
                 if "Form_ID=182" in x["url"] and "ID=" in x["url"]]
        print(f"Found {len(links)} event detail links")

        # 6) Visit each detail page and extract best-available date/time/location
        events = []
        for i, x in enumerate(links, start=1):
            url = x["url"]
            m_id = re.search(r"ID=(\d+)", url)
            event_id = m_id.group(1) if m_id else "unknown"
            title_guess = x["title"] or f"(untitled) ID={event_id}"
            print(f"[{i}/{len(links)}] Fetching: {title_guess}")

            # Navigate to the detail URL.
            # (Do it on the page, not the old frame handle; frames can detach.)
            page.goto(url, wait_until="domcontentloaded", timeout=20000)
            page.wait_for_timeout(600)

            # Find the frame that actually contains the FormDetail content
            content = get_content_frame("FormDetail.aspx")
            content.locator("body").wait_for(timeout=15000)

            detail_html = content.content()
            soup = BeautifulSoup(detail_html, "html.parser")

            text = " ".join(soup.get_text("\n", strip=True).split())

            # Try to find a date/time in the detail page text (best-effort)
            # Common patterns: 01/17/26, 01/17/2026, 1/17/26, etc.
            m = re.search(r"(\d{1,2}/\d{1,2}/\d{2,4})", text)
            dtstart = None

            if m:
                # If we find a date, parse it
                try:
                    dtstart = dtparser.parse(m.group(1), dayfirst=False, yearfirst=False)
                    dtstart = dtstart.replace(tzinfo=LOCAL_TZ)
                except:
                    dtstart = None

            # Fallback: use date in title if present
            if not dtstart:
                m2 = re.search(r"\((\d{2})/(\d{2})/(\d{2,4})\)", title_guess)
                if m2:
                    mm, dd, yy = m2.group(1), m2.group(2), m2.group(3)
                    year = int(yy)
                    if year < 100:
                        year += 2000
                    dtstart = datetime(year, int(mm), int(dd), 0, 0, tzinfo=LOCAL_TZ)

            if not dtstart:
                # Can't place it on a calendar without a date
                continue

            # Assume all-day unless we can find a time
            all_day = True
            dtend = dtstart + timedelta(days=1)

            # Try to find time like "7:00 PM" or "19:00"
            tm = re.search(r"(\d{1,2}:\d{2}\s*(AM|PM))", text, re.I)
            if tm:
                try:
                    # Parse date + time together
                    dtstart2 = dtparser.parse(dtstart.strftime("%m/%d/%Y") + " " + tm.group(1))
                    dtstart = dtstart2.replace(tzinfo=LOCAL_TZ)
                    dtend = dtstart + timedelta(hours=2)  # default duration
                    all_day = False
                except:
                    pass

            # Best-effort location: look for a "Location:" label
            loc = ""
            lm = re.search(r"Location:\s*(.+?)(?:\s{2,}|$)", text, re.I)
            if lm:
                loc = lm.group(1).strip()

            summary = re.sub(r"\s*\(\d{1,2}/\d{1,2}/\d{2,4}\)\s*$", "", title_guess).strip()

            events.append({
                "summary": summary,
                "dtstart": dtstart,
                "dtend": dtend,
                "description": text[:2000],  # cap so ICS doesn't get huge
                "location": loc,
                "url": url,
                "all_day": all_day,
            })

        ics = build_ics(events)
        with open(OUTPUT_ICS, "w", encoding="utf-8") as f:
            f.write(ics)

        print(f"Wrote {OUTPUT_ICS} with {len(events)} events")

        context.close()
        browser.close()


if __name__ == "__main__":
    main()

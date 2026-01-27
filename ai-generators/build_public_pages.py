# build_public_pages.py
import sys
import os
import yaml
import json
import re
from datetime import datetime

# =========================
# Utilities
# =========================
def escape_html(text):
    if not isinstance(text, str):
        return ""
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def slugify(text):
    """Generate URL-friendly slug from text"""
    if not text:
        return "item"
    text = re.sub(r'[^a-zA-Z0-9\s-]', '', str(text))
    text = re.sub(r'[\s]+', '-', text.strip().lower())
    return text or "item"

def load_data(filepath):
    if not filepath or not os.path.exists(filepath):
        if filepath:
            print(f"🔍 File not found: {filepath}")
        return []
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read().strip()
            if not content:
                print(f"⚠️ File is empty: {filepath}")
                return []
            if filepath.endswith(('.yaml', '.yml')):
                data = yaml.safe_load(content) or []
                return data if isinstance(data, list) else [data]
            elif filepath.endswith('.json'):
                data = json.loads(content) or []
                return data if isinstance(data, list) else [data]
    except Exception as e:
        print(f"❌ Failed to load {filepath}: {e}")
        return []
    print(f"⚠️ Unsupported file type: {filepath}")
    return []

def _first_nonempty(*vals):
    for v in vals:
        if isinstance(v, str) and v.strip():
            return v.strip()
        if isinstance(v, (int, float)):  # allow numeric fields (e.g., postal code)
            return str(v)
        if isinstance(v, dict) and "@value" in v and isinstance(v["@value"], str) and v["@value"].strip():
            return v["@value"].strip()
    return ""

def _as_list(val):
    if val is None:
        return []
    if isinstance(val, list):
        return [str(x).strip() for x in val if str(x).strip()]
    if isinstance(val, str) and val.strip():
        return [s.strip() for s in val.split(",") if s.strip()]
    return []

def _title_from_filename(path):
    base = os.path.splitext(os.path.basename(path))[0]
    return base.replace("-", " ").replace("_", " ").strip().title()

def _is_placeholder_title(text):
    if not isinstance(text, str) or not text.strip():
        return True
    t = text.strip().lower()
    return (
        t in {"service", "unnamed service", "untitled", "n/a", "na", "tbd"}
        or bool(re.fullmatch(r"(service|item|entry)\s*\d+", t))
    )

def _guess_description(obj):
    return _first_nonempty(
        obj.get("description"),
        obj.get("summary"),
        obj.get("details"),
        obj.get("body"),
        obj.get("content"),
        obj.get("answer"),
        obj.get("copy"),
    )

def _guess_price(obj):
    return _first_nonempty(
        obj.get("price"),
        obj.get("price_range"),
        obj.get("starting_price"),
        obj.get("min_price"),
        obj.get("cost"),
        obj.get("fee"),
    ) or "Contact for pricing"

def _bullet_points(obj):
    """Try to produce a few crisp bullets from common fields."""
    feats = _as_list(obj.get("features") or obj.get("benefits") or obj.get("highlights"))
    specs = _as_list(obj.get("specialties") or obj.get("capabilities"))
    areas = _as_list(obj.get("service_areas") or obj.get("areas") or obj.get("locations_served"))
    bullets = []
    for f in feats[:3]:
        bullets.append(f)
    if not bullets:
        for s in specs[:3]:
            bullets.append(s)
    if areas:
        bullets.append("Service areas: " + ", ".join(areas[:5]))
    # de-dupe while preserving order
    seen = set()
    uniq = []
    for b in bullets:
        if b.lower() not in seen:
            uniq.append(b)
            seen.add(b.lower())
    return uniq[:4]

# =========================
# Normalization helpers for Contact data
# =========================
FIELD_ALIASES = {
    "entity_name": ["entity_name", "organization", "org_name", "company", "name"],
    "contact_person": ["contact_person", "contact", "contact_name", "primary_contact", "attention"],
    "email": ["email", "contact_email", "email_address", "mail"],
    "phone": ["phone", "telephone", "tel", "phone_number", "contact_number"],
    "address_street": ["address_street", "streetAddress", "street", "address1", "address_line_1", "address_line"],
    "address_city": ["address_city", "city", "addressLocality"],
    "address_state": ["address_state", "state", "addressRegion", "province"],
    "address_postal_code": ["address_postal_code", "postalCode", "zip", "zipCode", "postcode"],
    "hours": ["hours", "openingHours", "opening_hours", "business_hours"],
    "map_embed_url": ["map_embed_url", "map", "map_iframe"],
    "google_maps_url": ["google_maps_url", "maps_url", "map_url"],
    "latitude": ["geo_latitude", "latitude", "lat"],
    "longitude": ["geo_longitude", "longitude", "lng", "lon"],
    "website": ["website", "url", "homepage"],
    "sameAs": ["sameAs", "same_as", "social", "social_links"],
}

def _alias_get(d: dict, canon_key: str):
    """Fetch a value by canonical key using FIELD_ALIASES (and nested geo, contactPoint)."""
    if not isinstance(d, dict):
        return None
    if canon_key in d and (d[canon_key] or d[canon_key] == 0):
        return d[canon_key]
    for k in FIELD_ALIASES.get(canon_key, []):
        if k in d and (d[k] or d[k] == 0):
            return d[k]
    if canon_key in ("latitude", "longitude"):
        geo = d.get("geo") or {}
        if isinstance(geo, dict):
            if canon_key == "latitude":
                return geo.get("latitude")
            else:
                return geo.get("longitude")
    if canon_key in ("phone", "email"):
        cp = d.get("contactPoint") or d.get("contact_point")
        if isinstance(cp, dict):
            if canon_key == "phone":
                return _first_nonempty(cp.get("telephone"), cp.get("phone"))
            else:
                return _first_nonempty(cp.get("email"))
    return None

def _format_address_from_components(loc: dict):
    line1 = _first_nonempty(_alias_get(loc, "address_street"))
    line2 = _first_nonempty(loc.get("address2"), loc.get("address_line_2"), loc.get("suite"))
    city  = _first_nonempty(_alias_get(loc, "address_city"))
    state = _first_nonempty(_alias_get(loc, "address_state"))
    zipc  = _first_nonempty(_alias_get(loc, "address_postal_code"))
    parts = [line1, line2, ", ".join([p for p in [city, state] if p]) if city or state else None, zipc]
    return " ".join([p for p in parts if p]).strip()

def _format_address(addr, loc):
    """Accepts string/dict or composes from components."""
    if isinstance(addr, str) and addr.strip():
        return addr.strip()
    if isinstance(addr, dict):
        line1 = _first_nonempty(addr.get("streetAddress"), addr.get("address1"), addr.get("addressLine1"))
        line2 = _first_nonempty(addr.get("address2"), addr.get("addressLine2"), addr.get("suite"))
        city  = _first_nonempty(addr.get("addressLocality"), addr.get("city"))
        state = _first_nonempty(addr.get("addressRegion"), addr.get("state"))
        zipc  = _first_nonempty(addr.get("postalCode"), addr.get("zip"), addr.get("zipCode"))
        parts = [line1, line2, ", ".join([p for p in [city, state] if p]) if city or state else None, zipc]
        return " ".join([p for p in parts if p]).strip()
    return _format_address_from_components(loc)

def _extract_hours(loc):
    hours = _first_nonempty(_alias_get(loc, "hours"))
    if hours:
        return hours
    spec = loc.get("openingHoursSpecification") or loc.get("opening_hours_specification")
    if isinstance(spec, list) and spec:
        rows = []
        for r in spec:
            if not isinstance(r, dict):
                continue
            day = _first_nonempty(r.get("dayOfWeek"), r.get("day"), r.get("weekday"))
            if isinstance(day, list) and day:
                day = day[0]
            if isinstance(day, str) and "/" in day:
                day = day.rsplit("/", 1)[-1]
            opens  = _first_nonempty(r.get("opens"), r.get("openingTime"))
            closes = _first_nonempty(r.get("closes"), r.get("closingTime"))
            if day and (opens or closes):
                rows.append(f"{day}: {opens or '—'} – {closes or '—'}")
        if rows:
            return "; ".join(rows)
    return ""

def _map_embed_src(loc, address):
    lat = _alias_get(loc, "latitude")
    lng = _alias_get(loc, "longitude")
    if isinstance(lat, (int, float)) and isinstance(lng, (int, float)):
        return f"https://www.google.com/maps?q={lat},{lng}&z=15&output=embed"
    map_url = _first_nonempty(_alias_get(loc, "map_embed_url"))
    gmaps   = _first_nonempty(_alias_get(loc, "google_maps_url"))
    if map_url:
        return map_url
    if gmaps:
        return gmaps
    if address:
        from urllib.parse import quote_plus
        return f"https://www.google.com/maps?q={quote_plus(address)}&output=embed"
    return ""

def _normalize_records(payload):
    """Support {locations:[...]}, [ ... ], or single object."""
    if isinstance(payload, list):
        return payload
    if isinstance(payload, dict):
        if isinstance(payload.get("locations"), list):
            return payload["locations"]
        return [payload]
    return []

# =========================
# Branding / meta driven by entity_name
# =========================
def _load_first_yaml_json(path_glob):
    import glob
    for p in glob.glob(path_glob):
        if os.path.isfile(p) and p.lower().endswith((".json", ".yaml", ".yml")):
            data = load_data(p)
            if data:
                return data[0] if isinstance(data, list) else data
    return None

def _discover_entity_name_from_other_schemas():
    probes = [
        "schemas/organization/*.*",
        "schemas/organizations/*.*",
        "schemas/company/*.*",
        "schemas/entity/*.*",
        "schemas/business/*.*",
        "schemas/reviews/*.*",
        "schemas/services/*.*",
        "schemas/locations/*.*",
    ]
    for pat in probes:
        obj = _load_first_yaml_json(pat)
        if not obj or not isinstance(obj, dict):
            continue
        candidate = _first_nonempty(
            obj.get("entity_name"),
            obj.get("name"),
            obj.get("legal_name"),
            obj.get("brand"),
            obj.get("company"),
            obj.get("organization"),
            obj.get("site_title"),
        )
        if candidate:
            return candidate
    return None

def load_org_meta():
    """
    Returns a dict with site-level branding pulled from schemas.
    {
      "name": <entity_name/name/etc>,
      "favicon": <path or url or None>,
      "logo": <path or url or None>
    }
    """
    meta = {"name": None, "favicon": None, "logo": None}
    candidate_dirs = [
        "schemas/organization", "schemas/organizations",
        "schemas/company", "schemas/entity", "schemas/business",
    ]
    import glob
    org_file = None
    for d in candidate_dirs:
        if os.path.isdir(d):
            cand = [p for p in glob.glob(os.path.join(d, "*.*")) if p.lower().endswith((".json",".yaml",".yml"))]
            if cand:
                org_file = cand[0]
                break

    org = None
    if org_file:
        data = load_data(org_file)
        org = data[0] if isinstance(data, list) else data

    if isinstance(org, dict):
        meta["name"] = _first_nonempty(
            org.get("entity_name"),
            org.get("name"),
            org.get("legal_name"),
            org.get("brand"),
            org.get("site_title"),
        )
        meta["logo"] = _first_nonempty(org.get("logo_url"), org.get("logo"))
        meta["favicon"] = _first_nonempty(org.get("favicon"), org.get("favicon_url"))

    if not meta["name"]:
        meta["name"] = _discover_entity_name_from_other_schemas()

    if not meta["name"]:
        repo_slug = os.getenv("GITHUB_REPOSITORY") or ""
        meta["name"] = repo_slug.split("/", 1)[-1].replace("-", " ").title() if repo_slug else "Site"

    return meta

# =========================
# HTML shell
# =========================
def generate_nav():
    return """
    <nav style="background: #2c3e50; padding: 1rem; margin-bottom: 2rem;">
        <ul style="list-style: none; display: flex; gap: 2rem; margin: 0; padding: 0; flex-wrap: wrap; justify-content: center;">
            <li><a href="index.html" style="color: white; text-decoration: none;">Home</a></li>
            <li><a href="about.html" style="color: white; text-decoration: none;">About</a></li>
            <li><a href="services.html" style="color: white; text-decoration: none;">Services</a></li>
            <li><a href="testimonials.html" style="color: white; text-decoration: none;">Testimonials</a></li>
            <li><a href="faqs.html" style="color: white; text-decoration: none;">FAQs</a></li>
            <li><a href="help.html" style="color: white; text-decoration: none;">Help</a></li>
            <li><a href="contact.html" style="color: white; text-decoration: none;">Contact</a></li>
        </ul>
    </nav>
    """

def generate_page(title, content):
    # Entity-driven branding for <title> and favicon
    org = load_org_meta()
    site_name = org.get("name") or "Site"
    page_title = f"{escape_html(site_name)} — {escape_html(title)}" if title else escape_html(site_name)
    favicon_href = org.get("favicon") or "favicon.ico"
    theme_color = "#2c3e50"

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>{page_title}</title>
    <meta name="application-name" content="{escape_html(site_name)}">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="theme-color" content="{theme_color}">
    <link rel="icon" href="{escape_html(favicon_href)}">
    <link rel="icon" type="image/png" sizes="32x32" href="icons/favicon-32.png">
    <link rel="icon" type="image/png" sizes="16x16" href="icons/favicon-16.png">
    <link rel="apple-touch-icon" sizes="180x180" href="icons/apple-touch-icon.png">
    <link rel="manifest" href="site.webmanifest">
    <style>
        body {{ font-family: -apple-system, BlinkMacSystemFont, sans-serif; max-width: 900px; margin: 0 auto; padding: 20px; line-height: 1.7; }}
        h1, h2, h3 {{ color: #2c3e50; }}
        a {{ color: #3498db; text-decoration: none; }}
        a:hover {{ text-decoration: underline; }}
        img {{ max-width: 100%; height: auto; }}
        .page-header {{ background: #ecf0f1; padding: 2rem; border-radius: 8px; margin-bottom: 2rem; text-align: center; }}
        .card {{ border: 1px solid #eee; padding: 1.5rem; border-radius: 8px; margin: 2rem 0; }}
        .badge {{ background: #3498db; color: white; padding: 0.25rem 0.5rem; border-radius: 4px; font-size: 0.9em; }}
    </style>
</head>
<body>
    {generate_nav()}
    <div class="page-header">
        <h1>{escape_html(title or site_name)}</h1>
    </div>
    {content}
    <footer style="margin-top: 4rem; padding-top: 2rem; border-top: 1px solid #eee; text-align: center; color: #7f8c8d;">
        <p>© {datetime.now().year} — Auto-generated from structured data. Last updated: {datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')}</p>
    </footer>
</body>
</html>"""

# =========================
# Pages
# =========================
def generate_contact_page():
    """
    Builds contact.html from schemas/locations/*.{json,yaml,yml}
    Always renders a top 'Quick Contact' card (name/email/phone) from the first location,
    then renders full location card(s) below.
    """
    locations_dir = "schemas/locations"
    print(f"🔍 Checking contact data in: {locations_dir}")
    if not os.path.exists(locations_dir):
        print(f"❌ Locations directory not found: {locations_dir} — skipping contact.html")
        return False

    def _extract_contact(loc):
        phone = _first_nonempty(_alias_get(loc, "phone"))
        email = _first_nonempty(_alias_get(loc, "email"))
        return phone, email

    def _extract_site_and_social(loc):
        website = _first_nonempty(_alias_get(loc, "website"))
        socials = _as_list(_alias_get(loc, "sameAs"))
        return website, socials

    items = []
    files_seen = records_seen = rendered = 0
    first_name = ""
    first_phone = ""
    first_email = ""

    for fname in sorted(os.listdir(locations_dir)):
        if not fname.lower().endswith((".json", ".yaml", ".yml")):
            continue
        files_seen += 1
        path = os.path.join(locations_dir, fname)
        data = load_data(path)
        if not data:
            continue

        for loc in _normalize_records(data):
            if not isinstance(loc, dict):
                continue
            records_seen += 1

            name   = _first_nonempty(_alias_get(loc, "entity_name"), loc.get("location_name"), "Location")
            person = _first_nonempty(_alias_get(loc, "contact_person"))
            phone, email = _extract_contact(loc)
            addr   = _format_address(loc.get("address"), loc)
            hours  = _extract_hours(loc)
            site, socials = _extract_site_and_social(loc)
            map_src = _map_embed_src(loc, addr)

            # Capture for Quick Contact (first record only)
            if not first_name:
                first_name = name or ""
            if not first_phone and phone:
                first_phone = phone
            if not first_email and email:
                first_email = email

            block = f"<div class='card'>"
            block += f"<h3>{escape_html(name)}</h3><p>"
            if person:
                block += f"<strong>Contact:</strong> {escape_html(person)}<br>"
            if addr:
                block += f"<strong>Address:</strong> {escape_html(addr)}<br>"
            if phone:
                block += f"<strong>Phone:</strong> <a href='tel:{escape_html(phone)}'>{escape_html(phone)}</a><br>"
            if email:
                block += f"<strong>Email:</strong> <a href='mailto:{escape_html(email)}'>{escape_html(email)}</a><br>"
            if hours:
                block += f"<strong>Hours:</strong> {escape_html(hours)}<br>"
            if site:
                block += f"<strong>Website:</strong> <a href='{escape_html(site)}' target='_blank' rel='nofollow'>{escape_html(site)}</a><br>"
            block += "</p>"

            if socials:
                block += "<p><strong>Find us:</strong> " + " • ".join(
                    f"<a href='{escape_html(s)}' target='_blank' rel='nofollow'>{escape_html(s)}</a>" for s in socials[:8]
                ) + "</p>"

            if map_src:
                block += f"""
                <div style="margin-top: 1rem;">
                    <iframe src="{escape_html(map_src)}" width="100%" height="320"
                            style="border:0; border-radius: 8px;" allowfullscreen loading="lazy"></iframe>
                </div>
                """

            block += "</div>"
            items.append(block)
            rendered += 1

    if not items:
        print(f"⚠️ No usable contact info found (scanned {files_seen} files, {records_seen} records). Skipping contact.html")
        return False

    # Intro + ALWAYS show Quick Contact (name + email + phone) from first record
    intro = "<p>We’d love to hear from you. Reach out using the details below or visit us at our offices.</p>"
    if first_name or first_phone or first_email:
        intro += "<div class='card'><h2>Quick Contact</h2>"
        if first_name:
            intro += f"<p><strong>{escape_html(first_name)}</strong></p>"
        if first_phone:
            intro += f"<p><strong>Phone:</strong> <a href='tel:{escape_html(first_phone)}'>{escape_html(first_phone)}</a></p>"
        if first_email:
            intro += f"<p><strong>Email:</strong> <a href='mailto:{escape_html(first_email)}'>{escape_html(first_email)}</a></p>"
        intro += "</div>"

    content = intro + "".join(items)
    with open("contact.html", "w", encoding="utf-8") as f:
        f.write(generate_page("Contact Us", content))

    print(f"✅ contact.html generated — {rendered} location card(s) from {files_seen} file(s), {records_seen} record(s)")
    return True

def generate_services_page():
    services_dir = "schemas/services"
    print(f"🔍 Checking services data in: {services_dir}")
    if not os.path.exists(services_dir):
        print(f"❌ Services directory not found: {services_dir} — skipping services.html")
        return False

    def _guess_title(obj, filename):
        candidate = _first_nonempty(
            obj.get("title"),
            obj.get("service_name"),
            obj.get("name"),
            obj.get("headline"),
            obj.get("service"),
            obj.get("offering"),
            obj.get("product_name"),
            obj.get("category"),
            obj.get("subtype"),
            obj.get("type"),
            obj.get("label"),
        )
        if _is_placeholder_title(candidate):
            kws = _as_list(obj.get("keywords"))
            if kws:
                candidate = " / ".join(kws[:2]).title()
        if _is_placeholder_title(candidate):
            candidate = _title_from_filename(filename)
        return candidate

    items = []
    files_processed = 0
    placeholders_fixed = 0

    for file in sorted(os.listdir(services_dir)):
        if not file.endswith((".json", ".yaml", ".yml")):
            continue
        filepath = os.path.join(services_dir, file)
        data = load_data(filepath)
        if not data:
            continue
        files_processed += 1

        records = data if isinstance(data, list) else [data]
        expanded = []
        for rec in records:
            if isinstance(rec, dict) and isinstance(rec.get("services"), list):
                expanded.extend(rec["services"])
            else:
                expanded.append(rec)

        for svc in expanded:
            if not isinstance(svc, dict):
                continue

            title_before = _first_nonempty(svc.get("title"), svc.get("service_name"), svc.get("name"))
            title = _guess_title(svc, filepath)
            if _is_placeholder_title(title_before) and not _is_placeholder_title(title):
                placeholders_fixed += 1

            description = _guess_description(svc) or ""
            price = _guess_price(svc)
            featured = bool(svc.get("featured") or svc.get("is_featured"))
            slug = svc.get("slug") or slugify(title)
            badge = '<span class="badge">Featured</span>' if featured else ''
            bullets = _bullet_points(svc)

            bullet_html = ""
            if bullets:
                bullet_html = "<ul>" + "".join(f"<li>{escape_html(b)}</li>" for b in bullets) + "</ul>"

            items.append(f"""
            <div class="card" id="{escape_html(slug)}">
                <h2>{escape_html(title)} {badge}</h2>
                {'<p>' + escape_html(description) + '</p>' if description else ''}
                {bullet_html}
                <p><strong>Starting at:</strong> {escape_html(price)}</p>
                <a href="#{slug}" style="display: inline-block; margin-top: 1rem;">🔗 Permalink</a>
            </div>
            """)

    if not items:
        print("⚠️ No valid services found — skipping services.html")
        return False

    with open("services.html", "w", encoding="utf-8") as f:
        f.write(generate_page("Our Services", "".join(items)))

    if placeholders_fixed:
        print(f"✨ Polished {placeholders_fixed} placeholder service title(s).")
    print(f"✅ services.html generated ({len(items)} services from {files_processed} file(s))")
    return True

def generate_testimonials_page():
    reviews_dir = "schemas/reviews"
    print(f"🔍 Checking testimonials data in: {reviews_dir}")
    if not os.path.exists(reviews_dir):
        print(f"❌ Reviews directory not found: {reviews_dir} — skipping testimonials.html")
        return False

    items = []
    for file in os.listdir(reviews_dir):
        if file.endswith((".json", ".yaml", ".yml")):
            filepath = os.path.join(reviews_dir, file)
            rev_data = load_data(filepath)
            if not rev_data:
                continue
            for rev in (rev_data if isinstance(rev_data, list) else [rev_data]):
                author = rev.get('customer_name') or rev.get('author') or 'Anonymous'
                entity = rev.get('entity_name') or ''
                quote = rev.get('review_body') or rev.get('quote') or rev.get('review_title') or 'No review text provided.'
                rating = int(rev.get('rating', 5))
                date = rev.get('date') or ''
                star_display = '★' * rating + '☆' * (5 - rating)
                items.append(f"""
                <blockquote class="card" style="font-style: italic;">
                    <p>“{escape_html(quote)}”</p>
                    <footer style="margin-top: 1rem; font-style: normal;">
                        — {escape_html(author)}{f', {escape_html(entity)}' if entity else ''}
                        {f'<br/><small>{escape_html(date)}</small>' if date else ''}
                    </footer>
                    <div style="margin-top: 0.5rem; color: #f39c12;">{star_display}</div>
                </blockquote>
                """)

    if not items:
        print("⚠️ No valid testimonials found — skipping testimonials.html")
        return False

    with open("testimonials.html", "w", encoding="utf-8") as f:
        f.write(generate_page("Testimonials", "".join(items)))
    print(f"✅ testimonials.html generated ({len(items)} testimonials)")
    return True

def generate_index_page():
    """Home: show 'Welcome to {Entity}' in the visible H1 and keep <title> = '{Entity} — Welcome'."""
    org = load_org_meta()
    site_name = org.get("name") or "Site"

    links = [
        ("About Us", "about.html"),
        ("Our Services", "services.html"),
        ("Testimonials", "testimonials.html"),
        ("FAQs", "faqs.html"),
        ("Help Center", "help.html"),
        ("Contact Us", "contact.html"),
        ("Browse All Schema Files", "#files"),
    ]
    quick_links = "\n".join(
        f'<li style="margin: 0.5rem 0;"><a href="{url}" style="font-size: 1.1em; font-weight: 500;">{escape_html(name)}</a></li>'
        for name, url in links
    )

    file_links = []
    repo_slug = os.getenv('GITHUB_REPOSITORY')
    if not repo_slug:
        print("❌ ERROR: GITHUB_REPOSITORY environment variable not set!")
        sys.exit(1)

    base_url = f"https://raw.githubusercontent.com/{repo_slug}/main"
    print(f"🌐 Base URL for schema files: {base_url}")

    for root, dirs, files in os.walk("schemas"):
        for file in files:
            if file.endswith((".json", ".yaml", ".yml", ".md", ".llm")):
                filepath = os.path.join(root, file).replace("\\", "/")
                full_url = f"{base_url}/{filepath}"
                display_path = filepath.replace("schemas/", "")
                file_links.append(f'<li><a href="{full_url}" target="_blank">{escape_html(display_path)}</a></li>')

    content = f"""
    <p>Welcome to our AI-optimized data hub. Below are quick links to key sections, or browse all machine-readable files.</p>
    <h2>🚀 Quick Navigation</h2>
    <ul style="list-style: none; padding: 0;">
        {quick_links}
    </ul>
    <h2 id="files">📁 All Schema Files</h2>
    <ul>
        {''.join(sorted(file_links))}
    </ul>
    """

    # Pass a visible header that includes the entity name.
    with open("index.html", "w", encoding="utf-8") as f:
        f.write(generate_page(f"Welcome to {site_name}", content))
    print("✅ index.html generated")
    return True

def generate_about_page():
    # Locate org file (if any)
    candidate_dirs = [
        "schemas/organization",
        "schemas/organizations",
        "schemas/company",
        "schemas/entity",
        "schemas/business",
    ]
    org_dir = next((d for d in candidate_dirs if os.path.isdir(d)), None)

    org = None
    picked_path = None
    if org_dir:
        cand = [f for f in os.listdir(org_dir) if f.endswith(('.json', '.yaml', '.yml'))]
        if cand:
            picked_path = os.path.join(org_dir, cand[0])
            data = load_data(picked_path)
            if data:
                org = data[0] if isinstance(data, list) else data

    # Gather summary from other schemas
    services_dir = "schemas/services"
    locations_dir = "schemas/locations"
    reviews_dir = "schemas/reviews"

    # Count services
    service_titles = []
    if os.path.isdir(services_dir):
        for file in os.listdir(services_dir):
            if not file.endswith((".json", ".yaml", ".yml")):
                continue
            for rec in (load_data(os.path.join(services_dir, file)) or []):
                if isinstance(rec, dict) and isinstance(rec.get("services"), list):
                    for s in rec["services"]:
                        if isinstance(s, dict):
                            title = _first_nonempty(s.get("title"), s.get("service_name"), s.get("name"))
                            if _is_placeholder_title(title):
                                title = None
                            service_titles.append(title or _title_from_filename(file))
                elif isinstance(rec, dict):
                    title = _first_nonempty(rec.get("title"), rec.get("service_name"), rec.get("name"))
                    if _is_placeholder_title(title):
                        title = None
                    service_titles.append(title or _title_from_filename(file))
    service_count = len(service_titles)

    # Locations / service areas
    service_areas = set()
    phone = email = ""
    if os.path.isdir(locations_dir):
        for file in os.listdir(locations_dir):
            if not file.endswith((".json", ".yaml", ".yml")):
                continue
            for loc in (load_data(os.path.join(locations_dir, file)) or []):
                if not isinstance(loc, dict):
                    continue
                for area in _as_list(loc.get("service_areas") or loc.get("areas")):
                    service_areas.add(area)
                if not phone:
                    phone = _first_nonempty(_alias_get(loc, "phone"))
                if not email:
                    email = _first_nonempty(_alias_get(loc, "email"))

    # Reviews: average rating
    ratings = []
    if os.path.isdir(reviews_dir):
        for file in os.listdir(reviews_dir):
            if not file.endswith((".json", ".yaml", ".yml")):
                continue
            for rev in (load_data(os.path.join(reviews_dir, file)) or []):
                if isinstance(rev, dict):
                    try:
                        r = float(rev.get("rating"))
                        if r > 0:
                            ratings.append(r)
                    except Exception:
                        pass
    avg_rating = (sum(ratings) / len(ratings)) if ratings else None

    # Build an org fallback if missing
    if not org:
        repo_slug = os.getenv("GITHUB_REPOSITORY") or ""
        fallback_name = repo_slug.split("/", 1)[-1].replace("-", " ").title() if repo_slug else "Our Company"
        org = {
            "entity_name": fallback_name,
            "name": fallback_name,
            "description": "",
            "mission": "",
            "vision": "",
            "logo_url": "",
            "website": "",
        }

    # Compose page
    parts = []
    display_name = _first_nonempty(org.get("entity_name"), org.get("name")) or "About Us"
    logo_url = _first_nonempty(org.get("logo_url"), org.get("logo"))
    if logo_url:
        parts.append(f'<img src="{escape_html(logo_url)}" alt="{escape_html(display_name)}" style="max-height: 120px; margin-bottom: 2rem;">')

    desc = _first_nonempty(org.get("description"))
    if not desc:
        desc = f"{display_name} is a professional firm serving our community with high-quality services and a client-first approach."
    parts.append(f"<p>{escape_html(desc)}</p>")

    facts = []
    facts.append(f"<strong>Services offered:</strong> {service_count}")
    if avg_rating is not None:
        stars = "★" * int(round(avg_rating)) + "☆" * (5 - int(round(avg_rating)))
        facts.append(f"<strong>Average rating:</strong> {avg_rating:.1f} {stars}")
    if service_areas:
        facts.append(f"<strong>Service areas:</strong> {escape_html(', '.join(sorted(list(service_areas))[:8]))}")
    if phone:
        facts.append(f"<strong>Phone:</strong> {escape_html(phone)}")
    if email:
        facts.append(f'<strong>Email:</strong> <a href="mailto:{escape_html(email)}">{escape_html(email)}</a>')

    parts.append('<div class="card"><h2>Facts at a Glance</h2><ul>' +
                 "".join(f"<li>{row}</li>" for row in facts) + "</ul></div>")

    if org.get("mission"):
        parts.append(f"<h2>Our Mission</h2><p>{escape_html(org.get('mission'))}</p>")
    if org.get("vision"):
        parts.append(f"<h2>Our Vision</h2><p>{escape_html(org.get('vision'))}</p>")

    website = _first_nonempty(org.get("website"), org.get("url"))
    same_as = _as_list(org.get("sameAs") or org.get("same_as"))
    if website or same_as:
        links = []
        if website:
            links.append(f'<li><a href="{escape_html(website)}" target="_blank" rel="nofollow">Website</a></li>')
        for s in same_as[:12]:
            links.append(f'<li><a href="{escape_html(s)}" target="_blank" rel="nofollow">{escape_html(s)}</a></li>')
        parts.append("<h2>Links</h2><ul>" + "".join(links) + "</ul>")

    parts.append(f"""
    <div class="card">
        <h2>Ready to Talk?</h2>
        <p>Have a project in mind or need guidance? We’re here to help.</p>
        <p><a href="contact.html">Contact us</a> to get started.</p>
    </div>
    """)

    with open("about.html", "w", encoding="utf-8") as f:
        f.write(generate_page(display_name, "\n".join(parts)))

    if picked_path:
        print(f"✅ about.html generated from {picked_path}")
    else:
        print("✅ about.html generated from synthesized data")
    return True

def generate_faq_page():
    faq_dir = "schemas/faqs"
    print(f"🔍 Checking FAQs in: {faq_dir}")
    if not os.path.exists(faq_dir):
        print(f"❌ FAQ directory not found: {faq_dir} — skipping faqs.html")
        return False

    items = []
    for file in os.listdir(faq_dir):
        if file.endswith((".json", ".yaml", ".yml")):
            filepath = os.path.join(faq_dir, file)
            faq_data = load_data(filepath)
            if not faq_data:
                continue
            for item in (faq_data if isinstance(faq_data, list) else [faq_data]):
                question = (item.get('question') or '').strip()
                answer = (item.get('answer') or '').strip()
                if not question:
                    continue
                items.append(f"""
                <div class="card">
                    <h3 style="margin: 0 0 0.5rem 0;">{escape_html(question)}</h3>
                    <p>{escape_html(answer)}</p>
                </div>
                """)

    if not items:
        print("⚠️ No valid FAQs found — skipping faqs.html")
        return False

    with open("faqs.html", "w", encoding="utf-8") as f:
        f.write(generate_page("Frequently Asked Questions", "".join(items)))
    print(f"✅ faqs.html generated ({len(items)} FAQs)")
    return True

def generate_help_articles_page():
    help_dir = "schemas/help-articles"
    print(f"🔍 Looking for help articles in: {help_dir}")
    if not os.path.exists(help_dir):
        print(f"❌ Folder not found: {help_dir}")
        return False

    files_found = [f for f in os.listdir(help_dir) if f.endswith(".md")]
    print(f"📄 Found {len(files_found)} .md files: {files_found[:5]}")

    if not files_found:
        print("⚠️ No .md files found — skipping help.html")
        return False

    articles = []
    for file in files_found:
        filepath = os.path.join(help_dir, file)
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()

        title = None
        body_lines = []
        in_frontmatter = False
        frontmatter_done = False

        for line in content.splitlines():
            if line.strip() == "---" and not frontmatter_done:
                if not in_frontmatter:
                    in_frontmatter = True
                else:
                    in_frontmatter = False
                    frontmatter_done = True
                continue

            if in_frontmatter and not frontmatter_done:
                if line.lower().startswith("title:"):
                    title = line.split(":", 1)[1].strip()
            else:
                body_lines.append(line)

        if not title:
            title = file.replace(".md", "").replace("-", " ").title()

        html_lines = []
        for line in body_lines:
            if line.startswith("## "):
                html_lines.append(f"<h2>{escape_html(line[3:])}</h2>")
            elif line.startswith("# "):
                html_lines.append(f"<h1>{escape_html(line[2:])}</h1>")
            elif line.startswith(("- ", "* ")):
                html_lines.append(f"<p>• {escape_html(line[2:])}</p>")
            elif line.strip() == "":
                html_lines.append("<br/>")
            else:
                html_lines.append(f"<p>{escape_html(line)}</p>")

        article_html = f"""
        <div class="card">
            <h2>{escape_html(title)}</h2>
            {''.join(html_lines)}
        </div>
        """
        articles.append(article_html)

    with open("help.html", "w", encoding="utf-8") as f:
        f.write(generate_page("Help Center", "".join(articles)))
    print(f"✅ help.html generated ({len(articles)} articles)")
    return True

# =========================
# Entry point
# =========================
def find_repo_root():
    """Find a directory that contains 'schemas' by walking up from script dir."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    cur = script_dir
    for _ in range(4):
        if os.path.isdir(os.path.join(cur, "schemas")):
            return cur
        parent = os.path.dirname(cur)
        if parent == cur:
            break
        cur = parent
    return script_dir  # fallback

if __name__ == "__main__":
    print("🚀 STARTING build_public_pages.py — GENERIC VERSION FOR ANY REPO")

    REPO_ROOT = find_repo_root()
    os.chdir(REPO_ROOT)
    print(f"✅ WORKING DIRECTORY SET TO: {REPO_ROOT}")

    if not os.path.exists("schemas"):
        print("❌ FATAL: schemas/ folder not found at repo root")
        sys.exit(1)
    else:
        print(f"📁 schemas/ contents: {os.listdir('schemas')[:10]}")

    open(".nojekyll", "w").close()
    print("✅ Created .nojekyll file for GitHub Pages")

    html_files = ["index.html", "about.html", "services.html", "testimonials.html", "faqs.html", "help.html", "contact.html"]
    for f in html_files:
        if os.path.exists(f):
            os.remove(f)
            print(f"🗑️ Deleted old {f} — forcing rebuild")

    page_generators = [
        ("index.html", generate_index_page),
        ("about.html", generate_about_page),
        ("services.html", generate_services_page),
        ("testimonials.html", generate_testimonials_page),
        ("faqs.html", generate_faq_page),
        ("help.html", generate_help_articles_page),
        ("contact.html", generate_contact_page),
    ]

    any_success = False
    for filename, generator in page_generators:
        try:
            success = generator()
            if success:
                print(f"✅ {filename} generated successfully")
                any_success = True
        except Exception as e:
            print(f"❌ {filename} generation failed: {e}")

    if not any_success:
        print("⚠️ No pages generated — check your schemas/* folders and filenames")
    else:
        print("\n🎉 BUILD COMPLETE — site ready for GitHub Pages deployment")

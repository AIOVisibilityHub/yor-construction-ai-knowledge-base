import os
import pandas as pd
import json
import re
import sys

# ============================================================
# NO-DUPES / OVERWRITE generator
#
# Fixes the "-1" duplicate files on re-run by:
#  1) Writing to deterministic filenames (slug + ext) and OVERWRITING if present.
#  2) Skipping duplicate slugs WITHIN the same run (same sheet) instead of appending -1.
#
# Optional:
#  --clean  -> remove previously-generated files under schemas/* for supported sections
#             before writing new ones (helps remove orphans).
#
# Supports BOTH:
#  - legacy workbook sheets: entity_info, Services, Team, Press/News Mentions, Awards & Certifications, etc.
#  - newer/legal workbook sheets: Business Info, Practice Areas, Lawyers, Media Mentions, Awards, Certifications,
#    Accreditations, Help Articles, FAQs, Reviews, Case Studies, Locations
#
# Normalizes a few columns so build_public_pages.py can render consistently:
#  - Help Articles: article / article_content / content -> Markdown body
#  - Reviews: review -> review_body (and quote fallback)
#  - Locations: address_postal -> address_postal_code ; open_hours -> hours
# ============================================================

def slugify(text: str) -> str:
    """Generate clean, URL-friendly slug from text"""
    if text is None:
        return "untitled"
    text = str(text).strip()
    if not text:
        return "untitled"
    text = re.sub(r"[^a-zA-Z0-9\s-]", "", text)
    text = re.sub(r"[\s]+", "-", text.strip().lower())
    return text or "untitled"


def _as_str(v):
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    return str(v).strip()


def _is_blank(v):
    return v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == ""


def get_first(row, keys, default=""):
    """Return the first non-empty value from row for any of the provided keys."""
    for k in keys:
        if k in row and not _is_blank(row[k]):
            return row[k]
    return default


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Strip whitespace from column names for stable access."""
    df.columns = [str(c).strip() for c in df.columns]
    return df


def deterministic_path(output_dir: str, base_slug: str, ext: str) -> str:
    """
    Deterministic filename: <slug><ext>
    OVERWRITES if it already exists (prevents "-1" files on rerun).
    """
    base_slug = slugify(base_slug)
    filename = f"{base_slug}{ext}"
    return os.path.join(output_dir, filename)


def write_json(path: str, data: dict):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False, default=str)


def write_md(path: str, title: str, slug: str, body: str, extra_frontmatter: dict | None = None):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    extra_frontmatter = extra_frontmatter or {}
    with open(path, "w", encoding="utf-8") as f:
        f.write("---\n")
        if title:
            f.write(f"title: {title}\n")
        f.write(f"slug: {slug}\n")
        for k, v in extra_frontmatter.items():
            if v is None or v == "":
                continue
            f.write(f"{k}: {v}\n")
        f.write("---\n\n")
        f.write(body or "")


def clean_output_dirs(canonical_output: dict):
    """
    Removes previously-generated files in known schema output dirs.
    Only deletes .json and .md in those folders, and leaves folders intact.
    """
    removed = 0
    for _, out_dir in canonical_output.items():
        if not os.path.isdir(out_dir):
            continue
        for fname in os.listdir(out_dir):
            if fname.lower().endswith((".json", ".md")):
                try:
                    os.remove(os.path.join(out_dir, fname))
                    removed += 1
                except Exception:
                    pass
    print(f"🧽 Cleaned {removed} previously-generated file(s) under schemas/*.")


def main(input_file="templates/AI-Visibility-Master-Template.xlsx", clean=False):
    print(f"📂 Opening Excel file: {input_file}")

    if not os.path.exists(input_file):
        print(f"❌ FATAL: Excel file not found at {input_file}")
        sys.exit(1)
    else:
        print(f"✅ Excel file confirmed at: {input_file}")

    try:
        xlsx = pd.ExcelFile(input_file)
        print(f"📄 Available sheets in workbook: {xlsx.sheet_names}")
    except Exception as e:
        print(f"❌ Failed to load Excel file: {e}")
        sys.exit(1)

    # Canonical sheet keys -> output dirs
    canonical_output = {
        "organization": "schemas/organization",
        "services": "schemas/services",
        "products": "schemas/products",
        "faqs": "schemas/faqs",
        "help_articles": "schemas/help-articles",
        "reviews": "schemas/reviews",
        "locations": "schemas/locations",
        "team": "schemas/team",
        "awards": "schemas/awards",
        "press": "schemas/press",
        "case_studies": "schemas/case-studies",
    }

    # Sheet aliases (supports both old and new/legal names)
    SHEET_ALIASES = {
        "organization": [
            "entity_info",
            "Business Info",
            "Business information",
            "Organization",
            "Company",
            "Firm Info",
        ],
        "services": [
            "Services",
            "Practice Areas",
            "Practice areas",
            "Service Areas",
            "Medical Specialties",
        ],
        "products": [
            "Products",
        ],
        "faqs": [
            "FAQs",
            "FAQ",
        ],
        "help_articles": [
            "Help Articles",
            "Help articles",
            "Articles",
            "Guides",
            "Blog",
        ],
        "reviews": [
            "Reviews",
            "Testimonials",
        ],
        "locations": [
            "Locations",
            "Offices",
        ],
        "team": [
            "Team",
            "Lawyers",
            "Attorneys",
            "Providers",
            "Staff",
        ],
        "awards": [
            "Awards & Certifications",
            "Awards",
            "Certifications",
            "Accreditations",
            "Licenses",
            "Awards, Certifications, Accreditations",
        ],
        "press": [
            "Press/News Mentions",
            "Media Mentions",
            "Press",
            "News",
            "Media",
        ],
        "case_studies": [
            "Case Studies",
            "Case studies",
            "Matters",
            "Results",
        ],
    }

    def norm_sheet(s: str) -> str:
        return re.sub(r"\s+", " ", str(s).strip().lower())

    alias_lookup = {}
    for canon, aliases in SHEET_ALIASES.items():
        for a in aliases:
            alias_lookup[norm_sheet(a)] = canon

    if clean:
        clean_output_dirs(canonical_output)

    processed_any = False

    for actual_sheet in xlsx.sheet_names:
        canon = alias_lookup.get(norm_sheet(actual_sheet))
        if not canon:
            print(f"⚠️ Skipping unsupported sheet: {actual_sheet}")
            continue

        output_dir = canonical_output[canon]
        os.makedirs(output_dir, exist_ok=True)

        print(f"\n📄 Processing sheet: {actual_sheet}  →  {canon}  →  {output_dir}")

        df = xlsx.parse(actual_sheet)
        df = normalize_columns(df)

        if df.empty:
            print(f"⚠️ Sheet '{actual_sheet}' is empty — skipping")
            continue

        print(f"🧹 Cleaned column names: {list(df.columns)}")

        processed_count = 0
        seen_slugs = set()  # prevents duplicates WITHIN the same run for a given sheet

        # ----------------------------
        # ORGANIZATION (usually 1 row)
        # ----------------------------
        if canon == "organization":
            row_obj = None
            for _, r in df.iterrows():
                if not r.dropna().empty:
                    row_obj = r
                    break

            if row_obj is None:
                print("⚠️ No usable rows in organization sheet — skipping")
            else:
                row = row_obj

                business_name = _as_str(get_first(row, ["business_name", "entity_name", "name", "company_name", "firm_name"]))
                main_website_url = _as_str(get_first(row, ["main_website_url", "website", "url"]))
                logo_url = _as_str(get_first(row, ["logo_url", "logo", "logoUrl"]))
                short_description = _as_str(get_first(row, ["short_description", "description", "tagline"]))
                long_description = _as_str(get_first(row, ["long_description", "about", "about_text"]))

                same_as = []
                for k in [
                    "facebook_url", "instagram_url", "linkedin_url", "twitter_url", "x_url",
                    "youtube_url", "tiktok_url", "pinterest_url", "yelp_url", "bbb_url",
                    "avvo_url", "martindale_url", "other_profiles"
                ]:
                    v = get_first(row, [k])
                    if _is_blank(v):
                        continue
                    vv = str(v).strip()
                    if k == "other_profiles" and "," in vv:
                        for part in [p.strip() for p in vv.split(",") if p.strip()]:
                            same_as.append(part)
                    else:
                        same_as.append(vv)

                org = {}
                for col in df.columns:
                    v = row.get(col)
                    if pd.isna(v):
                        continue
                    if hasattr(v, "item"):
                        v = v.item()
                    org[col] = v

                if business_name:
                    org["entity_name"] = business_name
                if main_website_url:
                    org["website"] = main_website_url
                    org["url"] = main_website_url
                if logo_url:
                    org["logo_url"] = logo_url
                if short_description and "description" not in org:
                    org["description"] = short_description
                if long_description:
                    org["about"] = long_description
                if same_as:
                    org["sameAs"] = same_as

                path = os.path.join(output_dir, "organization.json")
                try:
                    write_json(path, org)  # overwrite
                    print(f"✅ Generated (overwrite): {path}")
                    processed_count += 1
                    processed_any = True
                except Exception as e:
                    print(f"❌ Failed to write {path}: {e}")

            print(f"📊 Total processed in '{actual_sheet}': {processed_count} items")
            continue

        # ----------------------------
        # HELP ARTICLES → Markdown
        # ----------------------------
        if canon == "help_articles":
            for idx, row in df.iterrows():
                if row.dropna().empty:
                    continue

                title = _as_str(get_first(row, ["title", "article_title", "name", "headline"]))
                slug = _as_str(get_first(row, ["slug", "article_slug"]))
                body = _as_str(get_first(row, ["article_content", "article", "content", "body", "markdown"]))

                if not slug:
                    slug = slugify(title) if title else f"article-{idx+1}"
                slug = slugify(slug)

                if slug in seen_slugs:
                    print(f"↩️ Skipping duplicate slug in sheet (kept first): {slug}")
                    continue
                seen_slugs.add(slug)

                path = deterministic_path(output_dir, slug, ".md")

                try:
                    write_md(
                        path=path,
                        title=title,
                        slug=slug,
                        body=body,
                        extra_frontmatter={
                            "date": _as_str(get_first(row, ["date", "published_date", "publish_date"]))
                        }
                    )
                    print(f"✅ Generated (overwrite): {path}")
                    processed_count += 1
                    processed_any = True
                except Exception as e:
                    print(f"❌ Failed to write {path}: {e}")

            print(f"📊 Total processed in '{actual_sheet}': {processed_count} items")
            continue

        # ----------------------------
        # FAQs → JSON
        # ----------------------------
        if canon == "faqs":
            for idx, row in df.iterrows():
                if row.dropna().empty:
                    continue

                question = _as_str(get_first(row, ["question", "q", "faq_question", "title"]))
                answer = _as_str(get_first(row, ["answer", "a", "faq_answer", "response", "content"]))
                slug = _as_str(get_first(row, ["slug", "faq_id", "id"]))

                if not question:
                    question = f"Untitled FAQ {idx+1}"
                if not slug:
                    slug = slugify(question)
                slug = slugify(slug)

                if slug in seen_slugs:
                    print(f"↩️ Skipping duplicate slug in sheet (kept first): {slug}")
                    continue
                seen_slugs.add(slug)

                path = deterministic_path(output_dir, slug, ".json")
                data = {"question": question, "answer": answer}

                try:
                    write_json(path, data)
                    print(f"✅ Generated (overwrite): {path}")
                    processed_count += 1
                    processed_any = True
                except Exception as e:
                    print(f"❌ Failed to write {path}: {e}")

            print(f"📊 Total processed in '{actual_sheet}': {processed_count} items")
            continue

        # ----------------------------
        # SERVICES → JSON (supports Practice Areas too)
        # ----------------------------
        if canon == "services":
            for idx, row in df.iterrows():
                if row.dropna().empty:
                    continue

                service_name = _as_str(get_first(row, ["service_name", "practice_area", "practice_area_name", "name", "title"]))
                slug = _as_str(get_first(row, ["slug", "service_id", "id"]))
                description = _as_str(get_first(row, ["description", "service_description", "summary"]))
                price_range = _as_str(get_first(row, ["price_range", "priceRange"]))
                license_number = _as_str(get_first(row, ["license_number", "license"]))
                bar_number = _as_str(get_first(row, ["bar_number", "barNumber"]))
                npi_number = _as_str(get_first(row, ["npi_number", "npiNumber"]))
                certification_body = _as_str(get_first(row, ["certification_body", "certification"]))

                if not service_name:
                    service_name = f"Service {idx+1}"
                if not slug:
                    slug = slugify(service_name)
                slug = slugify(slug)

                if slug in seen_slugs:
                    print(f"↩️ Skipping duplicate slug in sheet (kept first): {slug}")
                    continue
                seen_slugs.add(slug)

                path = deterministic_path(output_dir, slug, ".json")

                data = {
                    "name": service_name,
                    "description": description,
                }
                if price_range:
                    data["priceRange"] = price_range
                if license_number:
                    data["license"] = license_number
                if bar_number:
                    data["barNumber"] = bar_number
                if npi_number:
                    data["npiNumber"] = npi_number
                if certification_body:
                    data["certification"] = certification_body

                for col in df.columns:
                    if col in data:
                        continue
                    v = row.get(col)
                    if pd.isna(v):
                        continue
                    if hasattr(v, "item"):
                        v = v.item()
                    data[col] = v

                try:
                    write_json(path, data)
                    print(f"✅ Generated (overwrite): {path}")
                    processed_count += 1
                    processed_any = True
                except Exception as e:
                    print(f"❌ Failed to write {path}: {e}")

            print(f"📊 Total processed in '{actual_sheet}': {processed_count} items")
            continue

        # ----------------------------
        # TEAM → JSON (supports Lawyers sheet)
        # ----------------------------
        if canon == "team":
            for idx, row in df.iterrows():
                if row.dropna().empty:
                    continue

                member_name = _as_str(get_first(row, ["member_name", "lawyer_name", "attorney_name", "name", "full_name"]))
                if not member_name:
                    fn = _as_str(get_first(row, ["first_name", "firstname"]))
                    ln = _as_str(get_first(row, ["last_name", "lastname"]))
                    member_name = " ".join([p for p in [fn, ln] if p]).strip()

                slug = _as_str(get_first(row, ["slug", "member_id", "lawyer_id", "id"]))
                role = _as_str(get_first(row, ["role", "title", "position"]))
                bio = _as_str(get_first(row, ["bio", "description", "about", "summary"]))
                license_number = _as_str(get_first(row, ["license_number", "license"]))
                bar_number = _as_str(get_first(row, ["bar_number", "barNumber"]))
                npi_number = _as_str(get_first(row, ["npi_number", "npiNumber"]))

                if not member_name:
                    member_name = f"Member {idx+1}"
                if not slug:
                    slug = slugify(member_name)
                slug = slugify(slug)

                if slug in seen_slugs:
                    print(f"↩️ Skipping duplicate slug in sheet (kept first): {slug}")
                    continue
                seen_slugs.add(slug)

                path = deterministic_path(output_dir, slug, ".json")

                data = {
                    "name": member_name,
                    "role": role,
                    "description": bio,
                }

                if license_number:
                    data["license"] = license_number
                if bar_number:
                    data["barNumber"] = bar_number
                if npi_number:
                    data["npiNumber"] = npi_number

                for col in df.columns:
                    if col in data:
                        continue
                    v = row.get(col)
                    if pd.isna(v):
                        continue
                    if hasattr(v, "item"):
                        v = v.item()
                    data[col] = v

                try:
                    write_json(path, data)
                    print(f"✅ Generated (overwrite): {path}")
                    processed_count += 1
                    processed_any = True
                except Exception as e:
                    print(f"❌ Failed to write {path}: {e}")

            print(f"📊 Total processed in '{actual_sheet}': {processed_count} items")
            continue

        # ----------------------------
        # REVIEWS → JSON (normalize review -> review_body)
        # ----------------------------
        if canon == "reviews":
            for idx, row in df.iterrows():
                if row.dropna().empty:
                    continue

                title = _as_str(get_first(row, ["review_title", "title", "headline"]))
                body = _as_str(get_first(row, ["review_body", "review", "quote", "testimonial", "content"]))
                slug = _as_str(get_first(row, ["slug", "review_id", "id"]))
                rating = get_first(row, ["rating", "stars"])
                date = _as_str(get_first(row, ["date", "review_date"]))

                if not slug:
                    slug = slugify(title) if title else f"review-{idx+1}"
                slug = slugify(slug)

                if slug in seen_slugs:
                    print(f"↩️ Skipping duplicate slug in sheet (kept first): {slug}")
                    continue
                seen_slugs.add(slug)

                path = deterministic_path(output_dir, slug, ".json")

                data = {}
                for col in df.columns:
                    v = row.get(col)
                    if pd.isna(v):
                        continue
                    if hasattr(v, "item"):
                        v = v.item()
                    data[col] = v

                if title:
                    data["review_title"] = title
                if body:
                    data["review_body"] = body
                    if "quote" not in data:
                        data["quote"] = body
                if not _is_blank(rating):
                    data["rating"] = rating
                if date:
                    data["date"] = date

                try:
                    write_json(path, data)
                    print(f"✅ Generated (overwrite): {path}")
                    processed_count += 1
                    processed_any = True
                except Exception as e:
                    print(f"❌ Failed to write {path}: {e}")

            print(f"📊 Total processed in '{actual_sheet}': {processed_count} items")
            continue

        # ----------------------------
        # LOCATIONS → JSON (normalize postal/hours)
        # ----------------------------
        if canon == "locations":
            for idx, row in df.iterrows():
                if row.dropna().empty:
                    continue

                name = _as_str(get_first(row, ["location_name", "name", "office_name", "title"]))
                slug = _as_str(get_first(row, ["slug", "location_id", "id"]))
                if not slug:
                    slug = slugify(name) if name else f"location-{idx+1}"
                slug = slugify(slug)

                if slug in seen_slugs:
                    print(f"↩️ Skipping duplicate slug in sheet (kept first): {slug}")
                    continue
                seen_slugs.add(slug)

                path = deterministic_path(output_dir, slug, ".json")

                data = {}
                for col in df.columns:
                    v = row.get(col)
                    if pd.isna(v):
                        continue
                    if hasattr(v, "item"):
                        v = v.item()
                    data[col] = v

                address_postal = _as_str(get_first(row, ["address_postal", "postal", "zip", "postal_code", "address_postal_code"]))
                open_hours = _as_str(get_first(row, ["open_hours", "hours", "opening_hours"]))
                if address_postal and "address_postal_code" not in data:
                    data["address_postal_code"] = address_postal
                if open_hours and "hours" not in data:
                    data["hours"] = open_hours
                if name and "location_name" not in data:
                    data["location_name"] = name

                try:
                    write_json(path, data)
                    print(f"✅ Generated (overwrite): {path}")
                    processed_count += 1
                    processed_any = True
                except Exception as e:
                    print(f"❌ Failed to write {path}: {e}")

            print(f"📊 Total processed in '{actual_sheet}': {processed_count} items")
            continue

        # ----------------------------
        # GENERIC HANDLER (press, awards, case studies, products, etc.)
        # ----------------------------
        for idx, row in df.iterrows():
            if row.dropna().empty:
                continue

            id_field = _as_str(get_first(row, [
                "slug", "id",
                "service_id", "product_id", "faq_id", "review_id", "location_id",
                "case_id", "press_id",
                "name", "title", "headline"
            ]))
            if not id_field:
                id_field = f"item-{idx+1}"
            slug = slugify(id_field)

            if slug in seen_slugs:
                print(f"↩️ Skipping duplicate slug in sheet (kept first): {slug}")
                continue
            seen_slugs.add(slug)

            path = deterministic_path(output_dir, slug, ".json")

            data = {}
            for col in df.columns:
                v = row.get(col)
                if pd.isna(v):
                    continue
                if hasattr(v, "item"):
                    v = v.item()
                data[col] = v

            if canon == "press":
                t = _as_str(get_first(row, ["title", "mention_title", "headline"]))
                if t and "title" not in data:
                    data["title"] = t

            try:
                write_json(path, data)
                print(f"✅ Generated (overwrite): {path}")
                processed_count += 1
                processed_any = True
            except Exception as e:
                print(f"❌ Failed to write {path}: {e}")

        print(f"📊 Total processed in '{actual_sheet}': {processed_count} items")

    if not processed_any:
        print("\n⚠️ No supported sheets were processed. Check sheet names vs aliases.")
        print("Supported sheet aliases include:", sorted({a for v in SHEET_ALIASES.values() for a in v}))
        sys.exit(2)

    print("\n🎉 All files generated successfully (no -1 duplicates).")


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Generate schema files from Excel (no -1 duplicates; overwrite).")
    parser.add_argument("--input", type=str, default="templates/AI-Visibility-Master-Template.xlsx",
                        help="Path to input Excel file")
    parser.add_argument("--clean", action="store_true",
                        help="Delete previously generated .json/.md files under known schemas/* folders before writing")
    args = parser.parse_args()
    main(args.input, clean=args.clean)

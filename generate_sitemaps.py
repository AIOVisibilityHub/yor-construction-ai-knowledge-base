# generate_sitemaps.py
import os
import re
import glob
import subprocess
from datetime import datetime
from xml.etree.ElementTree import Element, SubElement, ElementTree
from urllib.parse import quote

# ---------------------------
# Repo root / env helpers
# ---------------------------
def _run(cmd):
    try:
        return subprocess.check_output(cmd, stderr=subprocess.DEVNULL).decode("utf-8").strip()
    except Exception:
        return ""

def find_repo_root():
    """Walk up from this file until we find a directory containing either 'schemas' or '.git'."""
    cur = os.path.dirname(os.path.abspath(__file__))
    for _ in range(6):
        if os.path.isdir(os.path.join(cur, "schemas")) or os.path.isdir(os.path.join(cur, ".git")):
            return cur
        parent = os.path.dirname(cur)
        if parent == cur:
            break
        cur = parent
    return os.path.dirname(os.path.abspath(__file__))

def get_repo_slug():
    """
    Returns 'owner/repo' using (in order):
    - GITHUB_REPOSITORY env
    - git remote origin url
    """
    env_repo = os.getenv("GITHUB_REPOSITORY")
    if env_repo and "/" in env_repo:
        return env_repo

    origin = _run(["git", "config", "--get", "remote.origin.url"])
    m = re.search(r"github\.com[:/](?P<owner>[^/]+)/(?P<repo>[^/\.]+)", origin or "")
    if m:
        return f"{m.group('owner')}/{m.group('repo')}"

    # Last resort: read .git/config
    try:
        with open(".git/config", "r", encoding="utf-8") as f:
            cfg = f.read()
        m = re.search(r"github\.com[:/](?P<owner>[^/]+)/(?P<repo>[^/\.]+)", cfg)
        if m:
            return f"{m.group('owner')}/{m.group('repo')}"
    except Exception:
        pass

    raise RuntimeError("Unable to determine owner/repo. Set GITHUB_REPOSITORY or add a git remote origin.")

def get_branch_name():
    """
    Returns a branch/ref name using (in order):
    - GITHUB_REF_NAME env (Actions)
    - current git branch
    - 'main' fallback
    """
    ref = os.getenv("GITHUB_REF_NAME")
    if ref:
        return ref
    ref = _run(["git", "rev-parse", "--abbrev-ref", "HEAD"])
    return ref if ref and ref != "HEAD" else "main"

def get_raw_base_url(repo_slug: str, ref: str) -> str:
    # raw URLs for files inside the repo (json/yaml/md/etc.)
    # Later we append "/<path>", so no trailing slash here.
    return f"https://raw.githubusercontent.com/{repo_slug}/{quote(ref)}"

def get_pages_base_url(repo_slug: str) -> str:
    """
    If a CNAME file exists, use that as the site root.
    Otherwise project pages: https://<owner>.github.io/<repo>
    """
    cname_path = "CNAME"
    if os.path.isfile(cname_path):
        try:
            host = open(cname_path, "r", encoding="utf-8").read().strip()
            if host:
                host = host.replace("http://", "").replace("https://", "").strip("/")
                return f"https://{host}"
        except Exception:
            pass
    owner, repo = repo_slug.split("/", 1)
    return f"https://{owner}.github.io/{repo}"

# ---------------------------
# File discovery
# ---------------------------
def find_generated_files(patterns=None):
    """
    Returns repo-relative paths for data files we want in the AI sitemap.
    """
    if patterns is None:
        patterns = [
            "schemas/**/*.json", "schemas/**/*.yaml", "schemas/**/*.yml",
            "schemas/**/*.md", "schemas/**/*.llm",
        ]
    paths = set()
    for pat in patterns:
        matches = glob.glob(pat, recursive=True)
        for p in matches:
            if os.path.isfile(p):
                paths.add(p.replace("\\", "/"))
    paths = sorted(paths)
    print(f"🔎 AI files discovered: {len(paths)}")
    if paths:
        print("   e.g.", paths[:3])
    else:
        print("   (No matches found. Working dir:", os.getcwd(), ")")
        print("   schemas/ exists:", os.path.isdir("schemas"))
    return paths

def find_public_pages(extra_glob=False):
    """
    Public HTML pages we want in the standard sitemap.
    By default includes the core set; if extra_glob=True, also include any *.html at repo root.
    """
    core = [
        "index.html", "about.html", "services.html",
        "testimonials.html", "faqs.html", "help.html", "contact.html",
    ]
    pages = [p for p in core if os.path.exists(p)]
    if extra_glob:
        for p in glob.glob("*.html"):
            if p not in pages:
                pages.append(p)
    pages = sorted(pages)
    print(f"🔎 Public pages discovered: {len(pages)}")
    if pages:
        print("   e.g.", pages[:5])
    else:
        print("   (No HTML pages found in", os.getcwd(), ")")
    return pages

# ---------------------------
# Sitemap writer (pretty-printed)
# ---------------------------
import xml.dom.minidom, io

def write_sitemap(urls, out_file):
    urlset = Element("urlset", attrib={
        "xmlns": "http://www.sitemaps.org/schemas/sitemap/0.9"
    })
    now = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    for u in urls:
        url_el = SubElement(urlset, "url")
        SubElement(url_el, "loc").text = u
        SubElement(url_el, "lastmod").text = now

    rough = io.BytesIO()
    ElementTree(urlset).write(rough, encoding="utf-8", xml_declaration=True)
    pretty = xml.dom.minidom.parseString(rough.getvalue()).toprettyxml(indent="  ")
    with open(out_file, "w", encoding="utf-8") as outf:
        outf.write(pretty)
    print(f"📝 Wrote {out_file} with {len(urls)} URL(s)")

# ---------------------------
# Main
# ---------------------------
def main():
    import argparse
    ap = argparse.ArgumentParser(description="Generate sitemaps for GitHub-hosted data & pages.")
    ap.add_argument("--repo", help="Override 'owner/repo' (default: auto-detect)")
    ap.add_argument("--ref", help="Override branch/ref (default: auto-detect)")
    ap.add_argument("--raw-base", help="Override raw base URL")
    ap.add_argument("--pages-base", help="Override GitHub Pages base URL (or custom domain)")
    ap.add_argument("--skip-ai", action="store_true", help="Skip ai-sitemap.xml")
    ap.add_argument("--skip-pages", action="store_true", help="Skip sitemap.xml (HTML pages)")
    ap.add_argument("--include-all-html", action="store_true", help="Also include any *.html at repo root")
    args = ap.parse_args()

    # Ensure we run from repo root
    repo_root = find_repo_root()
    os.chdir(repo_root)
    print(f"📂 Working directory: {repo_root}")

    repo_slug = args.repo or get_repo_slug()
    ref = args.ref or get_branch_name()
    raw_base = args.raw_base or get_raw_base_url(repo_slug, ref)
    pages_base = args.pages_base or get_pages_base_url(repo_slug)

    print(f"📦 repo: {repo_slug}")
    print(f"🌿 ref : {ref}")
    print(f"📡 raw : {raw_base}")
    print(f"🌐 pages: {pages_base}")

    wrote_any = False

    # AI sitemap → raw machine-readable files
    if not args.skip_ai:
        data_files = find_generated_files()
        ai_urls = [f"{raw_base}/{p}" for p in data_files]
        write_sitemap(ai_urls, "ai-sitemap.xml")
        wrote_any = wrote_any or bool(ai_urls)

    # Standard sitemap → public HTML pages
    if not args.skip_pages:
        html_pages = find_public_pages(extra_glob=args.include_all_html)
        page_urls = [f"{pages_base}/{p}" for p in html_pages]
        write_sitemap(page_urls, "sitemap.xml")
        wrote_any = wrote_any or bool(page_urls)

    if not wrote_any:
        print("⚠️ No URLs found for either sitemap. Check:")
        print("   • Workflow order (generate data/pages BEFORE this step)")
        print("   • Working directory (script should be run at repo root)")
        print("   • That 'schemas/' and HTML files actually exist on the branch being built")

    print("✅ Done.")

if __name__ == "__main__":
    main()

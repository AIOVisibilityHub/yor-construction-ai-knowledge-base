# generate_robots.py
import os, re, subprocess
from urllib.parse import quote

def _run(cmd):
    try:
        return subprocess.check_output(cmd, stderr=subprocess.DEVNULL).decode().strip()
    except Exception:
        return ""

def get_repo_slug():
    env_repo = os.getenv("GITHUB_REPOSITORY")
    if env_repo and "/" in env_repo:
        return env_repo
    origin = _run(["git", "config", "--get", "remote.origin.url"])
    m = re.search(r"github\.com[:/](?P<owner>[^/]+)/(?P<repo>[^/\.]+)", origin or "")
    if m:
        return f"{m.group('owner')}/{m.group('repo')}"
    raise RuntimeError("Unable to detect repository slug")

def get_branch():
    ref = os.getenv("GITHUB_REF_NAME") or _run(["git", "rev-parse", "--abbrev-ref", "HEAD"])
    return ref if ref and ref != "HEAD" else "main"

def get_urls(repo_slug, branch):
    owner, repo = repo_slug.split("/", 1)
    raw_base = f"https://raw.githubusercontent.com/{repo_slug}/{quote(branch)}"
    pages_base = f"https://{owner}.github.io/{repo}"
    return raw_base, pages_base

def generate_robots():
    repo_slug = get_repo_slug()
    branch = get_branch()
    raw_base, pages_base = get_urls(repo_slug, branch)

    lines = [
        "User-agent: *",
        "Allow: /schemas/",
        "Allow: /ai-sitemap.xml",
        f"Sitemap: {raw_base}/ai-sitemap.xml",
        "",
        "# Explicitly invite AI crawlers",
        "User-agent: GPTBot",
        "Allow: /",
        "",
        "User-agent: ChatGPT-User",
        "Allow: /",
        "",
        "User-agent: PerplexityBot",
        "Allow: /",
        "",
        "User-agent: YouBot",
        "Allow: /",
        "",
        "User-agent: Claude-Web",
        "Allow: /",
        "",
        "User-agent: CCBot  # Common Crawl → feeds many AI models",
        "Allow: /",
        "",
        "User-agent: FacebookBot  # Meta AI",
        "Allow: /",
        "",
        "User-agent: anthropic-ai",
        "Allow: /",
        "",
        "# Generic fallback and sitemap for GitHub Pages site",
        "User-agent: *",
        "Allow: /",
        f"Sitemap: {pages_base}/sitemap.xml",
        "",
    ]

    with open("robots.txt", "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"✅ robots.txt written for {repo_slug} ({branch})")

if __name__ == "__main__":
    generate_robots()

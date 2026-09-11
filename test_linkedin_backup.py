import os
import re
import time
import requests
import pandas as pd

from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor, as_completed


# ============================================================
# CONFIGURATION
# ============================================================

BASE_URL = (
    "https://www.linkedin.com/jobs-guest/jobs/api/"
    "seeMoreJobPostings/search"
)

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/139.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}

# LinkedIn returns approximately 10 cards per request
RESULTS_PER_PAGE = 10

# Only postings from the last 48 hours (was 24h)
POSTED_WITHIN_SECONDS = 172800

# LinkedIn workplace-type filter: 1 = On-site, 2 = Remote, 3 = Hybrid
WORKPLACE_TYPE_FILTER = "2"

# LinkedIn job-type filter: F = Full-time, P = Part-time, C = Contract,
# T = Temporary, V = Volunteer, I = Internship, O = Other.
# This is the main lever for "projects, not jobs" - it asks LinkedIn
# to only return contract/freelance-type postings in the first place.
JOB_TYPE_FILTER = "C"

# Number of LinkedIn pages to check for each keyword
PAGES_PER_SEARCH = 3

# Delay between requests
REQUEST_DELAY = 2

# Whether to open each ambiguous listing's own page and check its
# actual "Workplace type" field / job description text, instead
# of trusting only the short text shown on the search results
# card. This adds one extra request per ambiguous listing, so it
# increases total runtime - set to False to skip this step.
VERIFY_REMOTE_WITH_JD = True

# Whether to also force a JD-page check of the "Employment type"
# field for EVERY listing (not just the remote-ambiguous ones), to
# confirm it really says "Contract" rather than trusting
# JOB_TYPE_FILTER alone. LinkedIn's server-side filters are known to
# leak the "wrong" type through sometimes (the same issue you saw
# with remote/hybrid). This is the main quality lever: every listing
# that ends up in the output has BOTH remote and contract status
# confirmed straight from LinkedIn's own fields, not just text
# heuristics. Defaults to True now that SEARCHES is trimmed (fewer
# raw listings to check, so the added requests stay manageable) -
# set to False if you'd rather trade some quality back for speed.
VERIFY_CONTRACT_WITH_JD = True

# Delay each worker thread waits around its own detail request in
# the JD-verification pass below. Kept slightly higher than
# REQUEST_DELAY since these hit a different, more sensitive
# endpoint (single job pages) many times in a row.
DETAIL_REQUEST_DELAY = 2

# How many detail-page requests to run in parallel during the JD
# verification pass. This is what actually fixes the "reading the
# JDs takes forever" problem - instead of one request every
# DETAIL_REQUEST_DELAY seconds, up to this many run at once.
# Keep this modest (3-6): too high and LinkedIn's guest endpoints
# are more likely to start throttling or blocking the requests.
MAX_WORKERS = 4

# Output folder
OUTPUT_DIR = "output"

OUTPUT_FILE = os.path.join(
    OUTPUT_DIR,
    "linkedin_remote_contract_projects_last_48h.xlsx"
)


# ============================================================
# SEARCH CATEGORIES
# ============================================================
#
# Trimmed from 67 keywords down to ~32. The dropped ones were
# near-synonyms of a term that's still in the list (e.g. "Software
# Developer" vs "Software Engineer", "ML Engineer" vs "Machine
# Learning Engineer") - LinkedIn's own search already matches
# related titles, so searching both mostly just re-fetches the
# same postings twice under two different search_keyword labels.
# A few generic, sales-flavored titles were also dropped entirely
# ("Solutions Engineer", "Technical Consultant") since they skew
# toward pre-sales/consulting roles rather than hands-on IT
# contract work. This roughly halves total requests and raw
# listings fetched, without losing category coverage.

SEARCHES = [

    # --------------------------------------------------------
    # SOFTWARE DEVELOPMENT
    # --------------------------------------------------------

    ("Software Engineer", "Remote"),
    ("Application Developer", "Remote"),

    # --------------------------------------------------------
    # WEB DEVELOPMENT
    # --------------------------------------------------------

    ("Full Stack Developer", "Remote"),
    ("Frontend Developer", "Remote"),
    ("Backend Developer", "Remote"),

    # --------------------------------------------------------
    # JAVASCRIPT / TYPESCRIPT
    # --------------------------------------------------------

    ("React Developer", "Remote"),
    ("Node.js Developer", "Remote"),
    ("TypeScript Developer", "Remote"),

    # --------------------------------------------------------
    # PYTHON / BACKEND
    # --------------------------------------------------------

    ("Python Developer", "Remote"),
    ("Django Developer", "Remote"),

    # --------------------------------------------------------
    # MOBILE
    # --------------------------------------------------------

    ("Mobile App Developer", "Remote"),
    ("Android Developer", "Remote"),
    ("iOS Developer", "Remote"),

    # --------------------------------------------------------
    # AI / MACHINE LEARNING
    # --------------------------------------------------------

    ("AI Engineer", "Remote"),
    ("Machine Learning Engineer", "Remote"),
    ("Generative AI Engineer", "Remote"),
    ("Computer Vision Engineer", "Remote"),

    # --------------------------------------------------------
    # DATA
    # --------------------------------------------------------

    ("Data Scientist", "Remote"),
    ("Data Engineer", "Remote"),

    # --------------------------------------------------------
    # DEVOPS / CLOUD
    # --------------------------------------------------------

    ("DevOps Engineer", "Remote"),
    ("Cloud Engineer", "Remote"),
    ("Site Reliability Engineer", "Remote"),
    ("AWS Developer", "Remote"),

    # --------------------------------------------------------
    # QA / TESTING
    # --------------------------------------------------------

    ("QA Engineer", "Remote"),
    ("QA Automation Engineer", "Remote"),

    # --------------------------------------------------------
    # CYBERSECURITY
    # --------------------------------------------------------

    ("Cybersecurity Engineer", "Remote"),
    ("Security Engineer", "Remote"),

    # --------------------------------------------------------
    # DATABASE
    # --------------------------------------------------------

    ("Database Engineer", "Remote"),
    ("SQL Developer", "Remote"),

    # --------------------------------------------------------
    # SYSTEMS / INTEGRATION
    # --------------------------------------------------------

    ("Systems Engineer", "Remote"),
    ("Integration Engineer", "Remote"),
    ("API Developer", "Remote"),
]


# ============================================================
# EXTRACT JOB ID
# ============================================================

def extract_job_id(card):

    # LinkedIn normally has:
    # data-entity-urn="urn:li:jobPosting:123456789"

    urn = card.get("data-entity-urn", "")

    match = re.search(
        r"jobPosting:(\d+)",
        urn
    )

    if match:
        return match.group(1)

    # Fallback
    html = str(card)

    match = re.search(
        r"jobPosting:(\d+)",
        html
    )

    if match:
        return match.group(1)

    return ""


# ============================================================
# EXTRACT JOB CARD
# ============================================================

def parse_job_card(card, keywords):

    title_element = card.select_one(
        ".base-search-card__title"
    )

    company_element = card.select_one(
        ".base-search-card__subtitle"
    )

    location_element = card.select_one(
        ".job-search-card__location"
    )

    link_element = card.select_one(
        "a.base-card__full-link"
    )

    date_element = card.select_one(
        ".job-search-card__listdate"
    )

    if not title_element:
        return None

    title = title_element.get_text(
        " ",
        strip=True
    )

    company = ""

    if company_element:
        company = company_element.get_text(
            " ",
            strip=True
        )

    location = ""

    if location_element:
        location = location_element.get_text(
            " ",
            strip=True
        )

    url = ""

    if link_element:
        url = link_element.get(
            "href",
            ""
        )

    # Remove tracking parameters
    if url:
        url = url.split("?")[0]

    posted = ""

    if date_element:
        posted = date_element.get_text(
            " ",
            strip=True
        )

    job_id = extract_job_id(card)

    return {
        "job_id": job_id,
        "title": title,
        "company": company,
        "location": location,
        "posted": posted,
        "url": url,
        "search_keyword": keywords,
        "source": "LinkedIn",
    }


# ============================================================
# WORKPLACE TYPE + JOB TYPE SAFETY-NET FILTERS
# ============================================================
#
# f_WT=2 and f_JT=C already ask LinkedIn for remote, contract-only
# postings, but LinkedIn's own filters are known to still let some
# hybrid/on-site/full-time postings through - these fields and the
# displayed location/title text are separate on LinkedIn's side, so
# a listing can be mistagged. This is a client-side safety net on
# top of the API filters, checking the visible location and title
# text for words that indicate the listing is NOT what we asked for.

NON_REMOTE_KEYWORDS = [
    "hybrid",
    "on-site",
    "onsite",
    "on site",
]

# Words in a title that indicate a listing is NOT contract/project
# work, even though f_JT=C asked LinkedIn to only return contract
# postings. Only checked against the title, since employment type
# isn't shown in the location text.
NON_CONTRACT_KEYWORDS = [
    "full-time",
    "full time",
    "permanent",
    "part-time",
    "part time",
    "internship",
    "temporary",
    "volunteer",
]


def is_remote_only(location_text, title_text=""):

    combined = f"{location_text} {title_text}".lower()

    return not any(
        keyword in combined
        for keyword in NON_REMOTE_KEYWORDS
    )


def is_likely_contract(title_text):

    # Benefit of the doubt: only reject here if the title itself
    # explicitly says something non-contract. Everything else is
    # left to JOB_TYPE_FILTER (and, optionally, the JD check).

    title = title_text.lower()

    return not any(
        keyword in title
        for keyword in NON_CONTRACT_KEYWORDS
    )


def needs_jd_check(location_text, title_text=""):

    # If the search-results card already explicitly says "Remote"
    # in the location OR title text, that's a confident signal on
    # its own - no need to spend an extra request confirming it.
    # Only listings that DIDN'T already declare themselves
    # non-remote (filtered out above) and DON'T already say
    # "Remote" anywhere are ambiguous enough to be worth checking
    # against the full JD.

    combined = f"{location_text} {title_text}".lower()

    return "remote" not in combined


def fetch_job_details(job_id):

    # Fetches a single job's own page from LinkedIn's guest API,
    # so we can check its real "Workplace type" / "Employment type"
    # fields and full job description, rather than just the short
    # text shown on the search results card.

    if not job_id:
        return "", "", ""

    url = f"https://www.linkedin.com/jobs-guest/jobs/api/jobPosting/{job_id}"

    try:

        response = requests.get(
            url,
            headers=HEADERS,
            timeout=30,
        )

        response.raise_for_status()

    except requests.RequestException as e:

        print(
            f"    Could not fetch job {job_id}: {e}"
        )

        return "", "", ""

    soup = BeautifulSoup(
        response.text,
        "html.parser"
    )

    description_element = soup.find(
        "div",
        class_="show-more-less-html__markup"
    )

    description_text = ""

    if description_element:
        description_text = description_element.get_text(
            " ",
            strip=True
        )

    # LinkedIn exposes workplace type and employment type as their
    # own criteria rows (alongside seniority level, job function,
    # etc.) - use them directly when present.
    workplace_type = ""
    employment_type = ""

    for item in soup.find_all("li", class_="description__job-criteria-item"):

        label_element = item.find("h3")
        value_element = item.find("span")

        if not label_element or not value_element:
            continue

        label = label_element.get_text(strip=True).lower()
        value = value_element.get_text(strip=True)

        if "workplace" in label:
            workplace_type = value

        elif "employment type" in label:
            employment_type = value

    return description_text, workplace_type, employment_type


def classify_remote_from_jd(description_text, workplace_type_field):

    # Returns True (confirmed remote), False (confirmed NOT
    # remote), or None (couldn't tell from what we fetched).
    #
    # Trust order:
    #   1. LinkedIn's own explicit "Workplace type" field, if present.
    #   2. Wording found in the actual job description text.

    if workplace_type_field:

        field = workplace_type_field.lower()

        if "remote" in field:
            return True

        if any(keyword in field for keyword in NON_REMOTE_KEYWORDS):
            return False

    if not description_text:
        return None

    text = description_text.lower()

    if any(keyword in text for keyword in NON_REMOTE_KEYWORDS):
        return False

    if "remote" in text or "work from home" in text or "work from anywhere" in text:
        return True

    return None


def classify_contract_from_jd(employment_type_field):

    # Returns True (confirmed contract), False (confirmed NOT
    # contract), or None (couldn't tell - the field was empty or
    # unrecognized). Unlike workplace type, employment type isn't
    # reliably mentioned in the job description body, so this only
    # trusts LinkedIn's own "Employment type" criteria field.

    if not employment_type_field:
        return None

    field = employment_type_field.lower()

    if "contract" in field:
        return True

    if any(keyword in field for keyword in NON_CONTRACT_KEYWORDS):
        return False

    return None


def verify_listing(job_id):

    # Runs inside a worker thread during the JD-verification pass.
    # Fetches one listing's detail page and returns both verdicts
    # (remote, contract) computed from it.

    time.sleep(DETAIL_REQUEST_DELAY)

    description_text, workplace_type_field, employment_type_field = fetch_job_details(
        job_id
    )

    remote_verdict = classify_remote_from_jd(
        description_text,
        workplace_type_field,
    )

    contract_verdict = classify_contract_from_jd(
        employment_type_field
    )

    return remote_verdict, contract_verdict


# ============================================================
# SCRAPE ONE SEARCH
# ============================================================

def scrape_linkedin_jobs(
    keywords,
    location,
    pages=PAGES_PER_SEARCH
):

    jobs = []

    for page in range(pages):

        # IMPORTANT:
        # LinkedIn is returning 10 jobs per page.
        start = page * RESULTS_PER_PAGE

        params = {
            "keywords": keywords,
            "location": location,

            # Pagination
            "start": start,

            # Remote listings only
            "f_WT": WORKPLACE_TYPE_FILTER,

            # Contract/project listings only (not full-time jobs)
            "f_JT": JOB_TYPE_FILTER,

            # Last POSTED_WITHIN_SECONDS only (48h by default now)
            "f_TPR": f"r{POSTED_WITHIN_SECONDS}",
        }

        print(
            f"Searching: {keywords} | "
            f"{location} | "
            f"page {page + 1}"
        )

        try:

            response = requests.get(
                BASE_URL,
                params=params,
                headers=HEADERS,
                timeout=30,
            )

            response.raise_for_status()

        except requests.RequestException as e:

            print(
                f"  Request failed: {e}"
            )

            continue

        soup = BeautifulSoup(
            response.text,
            "html.parser"
        )

        cards = soup.select(
            "li"
        )

        found = 0

        for card in cards:

            job = parse_job_card(
                card,
                keywords
            )

            if not job:
                continue

            jobs.append(job)

            found += 1

        print(
            f"  Found {found} jobs"
        )

        # If LinkedIn gives no results,
        # don't unnecessarily request more pages.
        if found == 0:
            break

        time.sleep(
            REQUEST_DELAY
        )

    return jobs


# ============================================================
# MAIN
# ============================================================

def main():

    all_jobs = []

    print()
    print("=" * 70)
    print("LINKEDIN REMOTE CONTRACT/PROJECT SCRAPER")
    print(f"FILTER: REMOTE + CONTRACT, LAST {POSTED_WITHIN_SECONDS // 3600} HOURS")
    print("=" * 70)
    print()

    # --------------------------------------------------------
    # Run searches
    # --------------------------------------------------------

    for keywords, location in SEARCHES:

        jobs = scrape_linkedin_jobs(
            keywords,
            location,
            pages=PAGES_PER_SEARCH,
        )

        all_jobs.extend(
            jobs
        )

        time.sleep(
            REQUEST_DELAY
        )

    # --------------------------------------------------------
    # Check results
    # --------------------------------------------------------

    if not all_jobs:

        print()
        print(
            "No listings found."
        )

        return

    df = pd.DataFrame(
        all_jobs
    )

    # --------------------------------------------------------
    # Remove empty URLs
    # --------------------------------------------------------

    df = df[
        df["url"].notna()
    ]

    df = df[
        df["url"] != ""
    ]

    # --------------------------------------------------------
    # Keep remote-only, contract-only listings (text-based pass)
    # --------------------------------------------------------

    before_remote_filter = len(df)

    df = df[
        df.apply(
            lambda row: is_remote_only(row["location"], row["title"])
            and is_likely_contract(row["title"]),
            axis=1,
        )
    ]

    removed_non_remote = before_remote_filter - len(df)

    # --------------------------------------------------------
    # Double-check ambiguous listings against their real JD
    # --------------------------------------------------------
    # A listing is "ambiguous" if either:
    #   - VERIFY_REMOTE_WITH_JD is on and its location/title text
    #     didn't already say "Remote" (needs_jd_check), or
    #   - VERIFY_CONTRACT_WITH_JD is on (in which case every
    #     listing is checked, since employment type never shows
    #     up in the card text at all)
    #
    # These checks now run concurrently (MAX_WORKERS at a time)
    # instead of one request every DETAIL_REQUEST_DELAY seconds,
    # which is what was making this step so slow before.

    removed_by_jd_check = 0

    if (VERIFY_REMOTE_WITH_JD or VERIFY_CONTRACT_WITH_JD) and len(df) > 0:

        df = df.reset_index(drop=True)

        ambiguous_mask = df.apply(
            lambda row: (
                VERIFY_REMOTE_WITH_JD
                and needs_jd_check(row["location"], row["title"])
            )
            or VERIFY_CONTRACT_WITH_JD,
            axis=1,
        )

        ambiguous_indices = df.index[ambiguous_mask].tolist()
        ambiguous_count = len(ambiguous_indices)

        print()
        print(
            f"Checking {ambiguous_count} ambiguous listing(s) "
            f"against their full JD ({MAX_WORKERS} at a time)..."
        )

        verdicts = {}
        checked = 0

        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:

            future_to_idx = {
                executor.submit(verify_listing, df.loc[idx, "job_id"]): idx
                for idx in ambiguous_indices
            }

            for future in as_completed(future_to_idx):

                idx = future_to_idx[future]

                try:
                    remote_verdict, contract_verdict = future.result()
                except Exception as e:
                    print(f"    Could not verify a listing: {e}")
                    remote_verdict, contract_verdict = None, None

                verdicts[idx] = (remote_verdict, contract_verdict)

                checked += 1

                if checked % 10 == 0:
                    print(f"  Checked {checked}/{ambiguous_count}")

        keep_flags = []

        for idx in df.index:

            if idx not in verdicts:
                keep_flags.append(True)
                continue

            remote_verdict, contract_verdict = verdicts[idx]

            # Treat "couldn't tell" (None) as keep - we don't want
            # to silently drop a legitimate listing just because we
            # couldn't confirm it from the JD page.
            keep_flags.append(
                remote_verdict is not False and contract_verdict is not False
            )

        before_jd_filter = len(df)

        df = df[keep_flags]

        removed_by_jd_check = before_jd_filter - len(df)

    # --------------------------------------------------------
    # Remove duplicates
    # --------------------------------------------------------

    before = len(df)

    df = df.drop_duplicates(
        subset=["url"],
        keep="first"
    )

    after = len(df)

    duplicates_removed = (
        before - after
    )

    # --------------------------------------------------------
    # Reorder columns
    # --------------------------------------------------------

    columns = [
        "job_id",
        "title",
        "company",
        "location",
        "posted",
        "url",
        "search_keyword",
        "source",
    ]

    df = df[
        columns
    ]

    # --------------------------------------------------------
    # Create output folder
    # --------------------------------------------------------

    os.makedirs(
        OUTPUT_DIR,
        exist_ok=True
    )

    # --------------------------------------------------------
    # Save Excel
    # --------------------------------------------------------

    # FIX: use a LOCAL variable (output_path) instead of
    # reassigning OUTPUT_FILE inside this function.
    #
    # Why: in Python, assigning to a name anywhere inside a
    # function makes that name local to the WHOLE function -
    # even on lines that run before the assignment. The old
    # code had "OUTPUT_FILE = fallback_file" inside main(),
    # which made every earlier use of OUTPUT_FILE in this
    # function a reference to an as-yet-unset local variable,
    # causing:
    #   UnboundLocalError: local variable 'OUTPUT_FILE'
    #   referenced before assignment
    output_path = OUTPUT_FILE

    try:

        df.to_excel(
            output_path,
            index=False,
        )

    except PermissionError:

        # If Excel is already open, use a timestamped file.
        timestamp = time.strftime(
            "%Y%m%d_%H%M%S"
        )

        output_path = os.path.join(
            OUTPUT_DIR,
            f"linkedin_jobs_last_24h_{timestamp}.xlsx"
        )

        df.to_excel(
            output_path,
            index=False,
        )

    # --------------------------------------------------------
    # Summary
    # --------------------------------------------------------

    print()
    print("=" * 70)
    print(
        f"NOT REMOTE/NOT CONTRACT (TEXT FILTER) REMOVED: {removed_non_remote}"
    )
    print(
        f"NOT REMOTE/NOT CONTRACT (CONFIRMED VIA JD): {removed_by_jd_check}"
    )
    print(
        f"TOTAL RAW JOBS: {before}"
    )
    print(
        f"DUPLICATES REMOVED: {duplicates_removed}"
    )
    print(
        f"TOTAL UNIQUE JOBS: {after}"
    )
    print("=" * 70)

    print()
    print(
        df[
            [
                "title",
                "company",
                "location",
                "posted",
            ]
        ]
        .head(20)
        .to_string(index=False)
    )

    print()
    print(
        f"Saved to: {os.path.abspath(output_path)}"
    )


if __name__ == "__main__":
    main()
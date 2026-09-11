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

# Last 7 days (604800 seconds) - wider window = more results
POSTED_WITHIN_SECONDS = 172800

# LinkedIn workplace-type filter: 1 = On-site, 2 = Remote, 3 = Hybrid
WORKPLACE_TYPE_FILTER = "2"

# LinkedIn job-type filter: F = Full-time, P = Part-time, C = Contract,
# T = Temporary, V = Volunteer, I = Internship, O = Other.
JOB_TYPE_FILTER = "C"

# Number of LinkedIn pages to check for each keyword (more = more results)
PAGES_PER_SEARCH = 5

# Delay between requests
REQUEST_DELAY = 2

# Whether to open each listing's own page and check its
# actual "Workplace type" field / job description text.
VERIFY_REMOTE_WITH_JD = True

# Whether to also force a JD-page check of the "Employment type"
# field for EVERY listing, to confirm it really says "Contract".
VERIFY_CONTRACT_WITH_JD = True

# STRICT MODE: when True, any listing where we CAN'T confirm
# remote status from the JD page is DROPPED (not kept).
# This is the key fix - previously None (unknown) was treated
# as "keep", which let hybrid/onsite jobs leak through.
STRICT_REMOTE_ONLY = True

# Delay each worker thread waits around its own detail request.
DETAIL_REQUEST_DELAY = 2

# How many detail-page requests to run in parallel.
MAX_WORKERS = 4

# Output folder
OUTPUT_DIR = "output"

OUTPUT_FILE = os.path.join(
    OUTPUT_DIR,
    "linkedin_remote_IT_contract_last_7d.xlsx"
)


# ============================================================
# IT-ONLY TITLE FILTER
# ============================================================
#
# This is the main filter: "only IT related posts".
# Every listing's title is checked against these keywords.
# If NONE of them appear in the title, the listing is dropped.
#
# This catches the problem where a search for "Python Developer"
# returns a "Data Entry Clerk" or "Account Manager" that
# happens to mention Python somewhere.

IT_TITLE_KEYWORDS = [
    # Software / Development
    "software", "developer", "development", "programmer", "coder",
    "engineer", "engineering", "architect",
    # Web
    "frontend", "front-end", "front end",
    "backend", "back-end", "back end",
    "full stack", "full-stack", "fullstack",
    "web developer", "web engineer",
    # Languages / Frameworks
    "python", "java", "javascript", "typescript", "react",
    "angular", "vue", "node.js", "nodejs", "node js",
    "django", "flask", "fastapi", "spring", "dotnet", ".net",
    "c#", "c++", "golang", "go developer", "rust", "ruby",
    "rails", "php", "laravel", "swift", "kotlin",
    "flutter", "react native",
    # Mobile
    "mobile", "android", "ios",
    # AI / ML / Data
    "ai ", " ai", "artificial intelligence",
    "machine learning", "ml ", " ml",
    "deep learning", "nlp", "natural language",
    "computer vision", "generative ai", "gen ai",
    "llm", "large language model",
    "data scientist", "data engineer", "data analyst",
    "data science", "data engineering",
    "agent", "agentic",
    # Cloud / DevOps / Infra
    "devops", "dev ops", "sre", "site reliability",
    "cloud", "aws", "azure", "gcp", "google cloud",
    "infrastructure", "platform engineer",
    "kubernetes", "k8s", "docker", "terraform",
    "ci/cd", "ci cd",
    # Cybersecurity
    "security", "cybersecurity", "cyber security",
    "infosec", "penetration", "soc analyst",
    # Database
    "database", "dba", "sql", "nosql", "mongodb",
    "postgresql", "mysql", "redis",
    # QA / Testing
    "qa", "quality assurance", "test engineer",
    "test automation", "sdet", "automation engineer",
    # IT / Systems
    "it ", " it ", "information technology",
    "systems engineer", "system engineer",
    "network engineer", "integration engineer",
    "api developer", "api engineer",
    "embedded", "firmware",
    # Blockchain / Web3
    "blockchain", "web3", "smart contract", "solidity",
    # UI/UX (technical design)
    "ui developer", "ux engineer", "ui/ux",
    # Tech lead / Management (still technical)
    "tech lead", "technical lead", "engineering manager",
    "cto", "vp engineering",
]

# Titles that are DEFINITELY not IT, even if they contain
# an IT keyword by accident (e.g., "Security Guard" contains
# "security"). These are checked FIRST.
NON_IT_TITLE_KEYWORDS = [
    "security guard", "security officer", "security patrol",
    "physical security",
    "nurse", "nursing", "medical", "clinical", "physician",
    "therapist", "dental", "pharmacy", "pharmacist",
    "teacher", "instructor", "tutor", "professor", "lecturer",
    "teaching",
    "driver", "delivery", "warehouse", "forklift",
    "cashier", "retail", "store manager", "store associate",
    "barista", "cook", "chef", "restaurant",
    "janitor", "custodian", "cleaning",
    "receptionist", "administrative assistant",
    "real estate", "realtor", "mortgage",
    "insurance agent", "insurance sales",
    "account manager", "account executive",
    "sales representative", "sales associate", "sales manager",
    "recruiter", "recruiting", "talent acquisition",
    "human resources", "hr manager", "hr specialist",
    "marketing manager", "marketing specialist",
    "social media manager", "content writer",
    "graphic designer",
    "legal", "paralegal", "attorney", "lawyer",
    "accountant", "bookkeeper", "tax preparer",
    "construction", "electrician", "plumber", "hvac",
    "mechanic", "technician",
    "data entry", "typist",
]


def is_it_related_title(title):
    """
    Returns True only if the job title is IT/tech related.
    Checks against NON_IT blacklist first, then IT whitelist.
    """
    title_lower = title.lower()

    # Step 1: reject if it matches any non-IT keyword
    for keyword in NON_IT_TITLE_KEYWORDS:
        if keyword in title_lower:
            return False

    # Step 2: accept only if it matches at least one IT keyword
    for keyword in IT_TITLE_KEYWORDS:
        if keyword in title_lower:
            return True

    # Step 3: no IT keyword found = not IT
    return False


# ============================================================
# SEARCH CATEGORIES
# ============================================================

SEARCHES = [

    # SOFTWARE DEVELOPMENT
    ("Software Engineer", ""),
    ("Software Developer", ""),
    ("Application Developer", ""),

    # WEB DEVELOPMENT
    ("Full Stack Developer", ""),
    ("Frontend Developer", ""),
    ("Backend Developer", ""),
    ("Web Developer", ""),

    # JAVASCRIPT / TYPESCRIPT
    ("React Developer", ""),
    ("Node.js Developer", ""),
    ("TypeScript Developer", ""),
    ("Angular Developer", ""),
    ("Vue.js Developer", ""),

    # PYTHON / BACKEND
    ("Python Developer", ""),
    ("Django Developer", ""),
    ("Java Developer", ""),
    (".NET Developer", ""),

    # MOBILE
    ("Mobile App Developer", ""),
    ("Android Developer", ""),
    ("iOS Developer", ""),
    ("Flutter Developer", ""),
    ("React Native Developer", ""),

    # AI / MACHINE LEARNING / AGENTS
    ("AI Engineer", ""),
    ("AI Developer", ""),
    ("Machine Learning Engineer", ""),
    ("Generative AI Engineer", ""),
    ("Computer Vision Engineer", ""),
    ("NLP Engineer", ""),
    ("LLM Engineer", ""),
    ("AI Agent Developer", ""),
    ("Deep Learning Engineer", ""),
    ("Data Scientist", ""),

    # DATA
    ("Data Engineer", ""),
    ("Data Analyst", ""),

    # DEVOPS / CLOUD
    ("DevOps Engineer", ""),
    ("Cloud Engineer", ""),
    ("Site Reliability Engineer", ""),
    ("AWS Developer", ""),
    ("Azure Developer", ""),
    ("Platform Engineer", ""),

    # QA / TESTING
    ("QA Engineer", ""),
    ("QA Automation Engineer", ""),
    ("SDET", ""),

    # CYBERSECURITY
    ("Cybersecurity Engineer", ""),
    ("Security Engineer", ""),

    # DATABASE
    ("Database Engineer", ""),
    ("SQL Developer", ""),

    # SYSTEMS / INTEGRATION
    ("Systems Engineer", ""),
    ("Integration Engineer", ""),
    ("API Developer", ""),
    ("Embedded Engineer", ""),
    ("Blockchain Developer", ""),
]


# ============================================================
# EXTRACT JOB ID
# ============================================================

def extract_job_id(card):
    urn = card.get("data-entity-urn", "")
    match = re.search(r"jobPosting:(\d+)", urn)
    if match:
        return match.group(1)
    html = str(card)
    match = re.search(r"jobPosting:(\d+)", html)
    if match:
        return match.group(1)
    return ""


# ============================================================
# EXTRACT JOB CARD
# ============================================================

def parse_job_card(card, keywords):

    title_element = card.select_one(".base-search-card__title")
    company_element = card.select_one(".base-search-card__subtitle")
    location_element = card.select_one(".job-search-card__location")
    link_element = card.select_one("a.base-card__full-link")
    date_element = card.select_one(".job-search-card__listdate")

    if not title_element:
        return None

    title = title_element.get_text(" ", strip=True)

    company = ""
    if company_element:
        company = company_element.get_text(" ", strip=True)

    location = ""
    if location_element:
        location = location_element.get_text(" ", strip=True)

    url = ""
    if link_element:
        url = link_element.get("href", "")
    if url:
        url = url.split("?")[0]

    posted = ""
    if date_element:
        posted = date_element.get_text(" ", strip=True)

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

NON_REMOTE_KEYWORDS = [
    "hybrid",
    "on-site",
    "onsite",
    "on site",
    "in-office",
    "in office",
]

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
    title = title_text.lower()
    return not any(
        keyword in title
        for keyword in NON_CONTRACT_KEYWORDS
    )


def needs_jd_check(location_text, title_text=""):
    combined = f"{location_text} {title_text}".lower()
    return "remote" not in combined


def fetch_job_details(job_id):

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
        print(f"    Could not fetch job {job_id}: {e}")
        return "", "", ""

    soup = BeautifulSoup(response.text, "html.parser")

    description_element = soup.find(
        "div",
        class_="show-more-less-html__markup"
    )

    description_text = ""
    if description_element:
        description_text = description_element.get_text(" ", strip=True)

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

    if not employment_type_field:
        return None

    field = employment_type_field.lower()

    if "contract" in field:
        return True

    if any(keyword in field for keyword in NON_CONTRACT_KEYWORDS):
        return False

    return None


def verify_listing(job_id):

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

def scrape_linkedin_jobs(keywords, location, pages=PAGES_PER_SEARCH):

    jobs = []

    for page in range(pages):

        start = page * RESULTS_PER_PAGE

        params = {
            "keywords": keywords,
            "start": start,
            "f_WT": WORKPLACE_TYPE_FILTER,
            "f_JT": JOB_TYPE_FILTER,
            "f_TPR": f"r{POSTED_WITHIN_SECONDS}",
        }

        # Only add location param if non-empty
        # (empty = worldwide search, more results)
        if location:
            params["location"] = location

        print(
            f"Searching: {keywords} | "
            f"{'Worldwide' if not location else location} | "
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
            print(f"  Request failed: {e}")
            continue

        soup = BeautifulSoup(response.text, "html.parser")
        cards = soup.select("li")

        found = 0

        for card in cards:
            job = parse_job_card(card, keywords)
            if not job:
                continue
            jobs.append(job)
            found += 1

        print(f"  Found {found} jobs")

        if found == 0:
            break

        time.sleep(REQUEST_DELAY)

    return jobs


# ============================================================
# MAIN
# ============================================================

def main():

    all_jobs = []

    print()
    print("=" * 70)
    print("LINKEDIN REMOTE IT CONTRACT/PROJECT SCRAPER")
    print(f"FILTER: REMOTE + CONTRACT + IT-ONLY, LAST {POSTED_WITHIN_SECONDS // 3600} HOURS")
    print(f"STRICT REMOTE: {STRICT_REMOTE_ONLY}")
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

        all_jobs.extend(jobs)
        time.sleep(REQUEST_DELAY)

    # --------------------------------------------------------
    # Check results
    # --------------------------------------------------------

    if not all_jobs:
        print()
        print("No listings found.")
        return

    df = pd.DataFrame(all_jobs)
    total_raw = len(df)

    # --------------------------------------------------------
    # Remove empty URLs
    # --------------------------------------------------------

    df = df[df["url"].notna()]
    df = df[df["url"] != ""]

    # --------------------------------------------------------
    # FILTER 1: IT-only titles
    # --------------------------------------------------------

    before_it_filter = len(df)

    df = df[df["title"].apply(is_it_related_title)]

    removed_non_it = before_it_filter - len(df)

    print()
    print(
        f"[FILTER 1] Removed {removed_non_it} non-IT titles "
        f"({before_it_filter} -> {len(df)})"
    )

    # --------------------------------------------------------
    # FILTER 2: Remote-only + contract-only (text-based pass)
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

    print(
        f"[FILTER 2] Removed {removed_non_remote} non-remote/non-contract "
        f"by card text ({before_remote_filter} -> {len(df)})"
    )

    # --------------------------------------------------------
    # FILTER 3: Verify via JD page (strict mode)
    # --------------------------------------------------------

    removed_by_jd_check = 0

    if (VERIFY_REMOTE_WITH_JD or VERIFY_CONTRACT_WITH_JD) and len(df) > 0:

        df = df.reset_index(drop=True)

        # In strict mode, check ALL listings, not just ambiguous ones
        if STRICT_REMOTE_ONLY:
            ambiguous_indices = df.index.tolist()
        else:
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
            f"[FILTER 3] Checking {ambiguous_count} listing(s) "
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

            if STRICT_REMOTE_ONLY:
                # STRICT: only keep if CONFIRMED remote (True).
                # None (unknown) and False both get dropped.
                remote_ok = (remote_verdict is True)
            else:
                # Lenient: only drop if CONFIRMED not remote (False).
                remote_ok = (remote_verdict is not False)

            # Contract: still lenient - only drop if confirmed NOT contract
            contract_ok = (contract_verdict is not False)

            keep_flags.append(remote_ok and contract_ok)

        before_jd_filter = len(df)

        df = df[keep_flags]

        removed_by_jd_check = before_jd_filter - len(df)

        print(
            f"[FILTER 3] Removed {removed_by_jd_check} "
            f"non-remote/non-contract via JD check "
            f"({before_jd_filter} -> {len(df)})"
        )

    # --------------------------------------------------------
    # Remove duplicates
    # --------------------------------------------------------

    before = len(df)

    df = df.drop_duplicates(
        subset=["url"],
        keep="first"
    )

    after = len(df)

    duplicates_removed = before - after

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

    df = df[columns]

    # --------------------------------------------------------
    # Create output folder
    # --------------------------------------------------------

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # --------------------------------------------------------
    # Save Excel
    # --------------------------------------------------------

    output_path = OUTPUT_FILE

    try:
        df.to_excel(output_path, index=False)

    except PermissionError:

        timestamp = time.strftime("%Y%m%d_%H%M%S")

        output_path = os.path.join(
            OUTPUT_DIR,
            f"linkedin_IT_remote_contract_{timestamp}.xlsx"
        )

        df.to_excel(output_path, index=False)

    # --------------------------------------------------------
    # Summary
    # --------------------------------------------------------

    print()
    print("=" * 70)
    print(f"TOTAL RAW SCRAPED:                    {total_raw}")
    print(f"REMOVED (NOT IT-RELATED TITLE):       {removed_non_it}")
    print(f"REMOVED (HYBRID/ONSITE BY CARD TEXT): {removed_non_remote}")
    print(f"REMOVED (CONFIRMED VIA JD PAGE):      {removed_by_jd_check}")
    print(f"DUPLICATES REMOVED:                   {duplicates_removed}")
    print(f"FINAL UNIQUE IT REMOTE CONTRACTS:     {after}")
    print("=" * 70)

    print()
    print(
        df[["title", "company", "location", "posted"]]
        .head(20)
        .to_string(index=False)
    )

    print()
    print(f"Saved to: {os.path.abspath(output_path)}")


if __name__ == "__main__":
    main()

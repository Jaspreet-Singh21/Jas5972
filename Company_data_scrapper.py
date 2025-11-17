from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import pandas as pd
import time
import random
import re
import os
from datetime import datetime

#input
input_file = "naukri_companies_all.xlsx"
output_file = "naukri_company_jobs.xlsx"

options = Options()
options.add_argument("--headless")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")
options.add_argument(
    "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0.0.0 Safari/537.36"
)
driver = webdriver.Chrome(options=options)

def clean_excel_text(value):
    """Remove illegal characters that Excel doesn't support."""
    if isinstance(value, str):
        return re.sub(r"[\x00-\x08\x0b-\x0c\x0e-\x1f]", "", value)
    return value


def safe_save_excel(data, filename, backup_interval=False):
    """Safely save DataFrame to Excel and create timestamped backups."""
    df = pd.DataFrame(data)
    df = df.map(clean_excel_text)

    # Save main file
    df.to_excel(filename, index=False)

    # Create timestamped backup every N intervals
    if backup_interval:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_name = f"backup_{os.path.splitext(filename)[0]}_{timestamp}.xlsx"
        df.to_excel(backup_name, index=False)
        print(f"[backup] 💾 Backup saved as '{backup_name}'")

    print(f"Saved {len(df)} rows to '{filename}'")


#read
companies = pd.read_excel(input_file)
print(f"[info] Loaded {len(companies)} company URLs from '{input_file}'")

job_data = []

#main
for idx, row in companies.iterrows():
    base_company_url = row["Full URL"]
    company_name = row["Company Name"]

    print(f"\n[info] ({idx+1}/{len(companies)}) Fetching jobs for {company_name}")
    total_jobs = 0

    first_url = (
        base_company_url
        if "pageNo=" in base_company_url
        else re.sub(r"(&pageNo=\d+)?$", "&pageNo=1", base_company_url)
    )

    print(f"[page-check] Checking first page: {first_url}")
    driver.get(first_url)
    time.sleep(random.uniform(5, 8))
    soup = BeautifulSoup(driver.page_source, "html.parser")

    # find total pages
    max_page = 1
    pag_wrapper = soup.find("div", class_="paginationWrapper")
    if pag_wrapper:
        pages_div = pag_wrapper.find("div", class_=re.compile(r"\bpages\b"))
        if pages_div:
            page_nums = []
            for a in pages_div.find_all("a"):
                txt = (a.text or "").strip()
                if txt.isdigit():
                    page_nums.append(int(txt))
                else:
                    dt = a.get("data-test", "") or a.get("data-testid", "")
                    m = re.search(r"pageNumber-(\d+)", dt)
                    if m:
                        page_nums.append(int(m.group(1)))
            if page_nums:
                max_page = max(page_nums)

    print(f"[pagination] Detected pages for {company_name}: {max_page}")

#loop
    for page_no in range(1, max_page + 1):
        if "pageNo=" in base_company_url:
            url = re.sub(r"pageNo=\d+", f"pageNo={page_no}", base_company_url)
        else:
            sep = "&" if "?" in base_company_url else "?"
            url = f"{base_company_url}{sep}pageNo={page_no}"

        print(f"[page] Fetching page {page_no}: {url}")
        driver.get(url)
        time.sleep(random.uniform(5, 8))

        soup = BeautifulSoup(driver.page_source, "html.parser")

        # Extract industry (once per page)
        industry_div = soup.find("div", class_="chips typ-14Medium")
        industry = industry_div.text.strip() if industry_div else "N/A"

        # Extract total job count (for reference)
        job_header = soup.find("h2", class_="headJobs")
        total_text = job_header.text.strip() if job_header else ""
        total_match = re.search(r"(\d+)\s+Job openings", total_text)
        total_expected = int(total_match.group(1)) if total_match else None

        # Find job listings
        job_articles = soup.find_all(
            "article", {"data-test": "tupleContainer", "class": "jobTuple"}
        )
        if not job_articles:
            print(f"[page-empty] No job cards found on page {page_no} for {company_name}")
            continue

        for job in job_articles:
            try:
                title_tag = job.find("a", {"data-test": "tupleTitle"})
                job_title = title_tag.text.strip() if title_tag else "N/A"
                job_url = title_tag.get("href", "") if title_tag else "N/A"

                exp_tag = job.find("span", {"data-test": "experience"})
                experience = exp_tag.text.strip() if exp_tag else "N/A"

                sal_tag = job.find("span", {"data-test": "salary"})
                salary = sal_tag.text.strip() if sal_tag else "N/A"

                loc_tag = job.find("span", {"data-test": "location"})
                location = loc_tag.text.strip() if loc_tag else "N/A"

                desc_tag = job.find("div", class_="fs13 ellipsis")
                description = desc_tag.text.strip() if desc_tag else "N/A"

                tags = [li.text.strip() for li in job.select(".jobtags li")] or ["N/A"]
                tags_combined = ", ".join(tags)

                job_data.append({
                    "Company": company_name,
                    "Industry": industry,
                    "Job Title": job_title,
                    "Experience": experience,
                    "Salary": salary,
                    "Location": location,
                    "Description": description,
                    "Skills / Tags": tags_combined,
                    "Job URL": job_url
                })
                total_jobs += 1
            except Exception as e:
                print(f"Error parsing job for {company_name}: {e}")

        print(f"Page {page_no} done ({len(job_articles)} jobs). Total: {total_jobs}")
        time.sleep(random.uniform(4, 6))

#compare check
    if total_expected and total_jobs < total_expected:
        print(f"{company_name}: Expected {total_expected} jobs, scraped {total_jobs}.")
    else:
        print(f"Completed {company_name} ({total_jobs} jobs).")

    if (idx + 1) % 100 == 0:
        safe_save_excel(job_data, output_file, backup_interval=True)

#save
driver.quit()
safe_save_excel(job_data, output_file)
print(f"\n Extracted {len(job_data)} jobs from {len(companies)} companies.")

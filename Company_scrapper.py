from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import pandas as pd
import time
import random

#input
base_url = "https://www.naukri.com"
template_url = "https://www.naukri.com/companies-hiring-in-india?src=gnbCompanies_homepage_srch&pageNo={}&liveAge=15"

options = Options()
options.add_argument("--headless")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")
options.add_argument(
    "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

driver = webdriver.Chrome(options=options)
all_companies = []
page_no = 1

print("[info] Starting Naukri company scraper...")

while True:
    url = template_url.format(page_no)
    print(f"\n[info] Fetching page {page_no}: {url}")

    driver.get(url)
    time.sleep(random.uniform(5, 8))

    soup = BeautifulSoup(driver.page_source, "html.parser")
    wrapper = soup.find("div", class_="tupple-wrapper")

    if not wrapper:
        print(f"[warn] No 'tupple-wrapper' found on page {page_no}. Stopping.")
        break

    company_divs = wrapper.find_all("div", class_="freeTuple")
    if not company_divs:
        print(f"[info] No more companies found on page {page_no}. Finished scraping.")
        break

    for div in company_divs:
        try:
            comp_id = div.get("id", "").strip()
            a = div.find("a", class_="titleAnchor")
            href = a.get("href", "").replace("&amp;", "&")
            name = a.text.strip()
            slug = href.split("/")[1] if "/" in href else ""

            all_companies.append({
                "Company ID": comp_id,
                "Slug": slug,
                "Relative URL": href,
                "Full URL": base_url + href,
                "Company Name": name
            })
        except Exception as e:
            print(f"[warn] Skipping entry: {e}")

    print(f"[info] ✅ Page {page_no} → {len(company_divs)} companies scraped so far ({len(all_companies)} total).")

    # polite random delay before next page
    time.sleep(random.uniform(4, 6))
    page_no += 1

driver.quit()

# --- SAVE TO EXCEL ---
df = pd.DataFrame(all_companies)
df.to_excel("naukri_companies_all.xlsx", index=False)
print(f"\nCompleted! Scraped {len(df)} companies across {page_no-1} pages.")
print("Saved to naukri_companies_all.xlsx")

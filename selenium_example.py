from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import csv


def get_driver(headless: bool= True) -> webdriver.Chrome:
    opts = webdriver.ChromeOptions()
    if headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=opts)


def scrape_first_page() -> list[dict]:
    driver = get_driver(headless=True)
    try:
        driver.get("https://quotes.toscrape.com/")

        wait = WebDriverWait(driver, 10)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.quote")))
        cards = driver.find_elements(By.CSS_SELECTOR, "div.quote")


        rows = []
        for q in cards:
            text = q.find_element(By.CSS_SELECTOR, "span.text").text
            author = q.find_element(By.CSS_SELECTOR, "small.author").text
            tags = [t.text for t in q.find_elements(By.CSS_SELECTOR, "div.tags a.tag")]
            rows.append({"quote": text, "author": author, "tags": ", ".join(tags)})
        return rows
    finally:
        driver.quit()


def save_csv(rows: list[dict], out_path: Path):
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["quote", "author", "tags"])
        w.writeheader()
        w.writerows(rows)


if __name__ == "__main__":
    data = scrape_first_page()
    out = Path("data/processed/quotes_page1.csv")
    save_csv(data, out)
    print(f"Ok -> {out} ({len(data)} filas)")

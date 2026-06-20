import requests
from bs4 import BeautifulSoup
import time

def scrape_books(target_count=50):
    """
    Task 1: Web Scraping using BeautifulSoup
    Scrapes book data from http://books.toscrape.com
    """
    print(f"Starting web scraping to collect {target_count} books...")
    
    books = []
    page = 1
    rating_map = {
        "one": 1, "two": 2, "three": 3, "four": 4, "five": 5
    }
    
    while len(books) < target_count:
        url = f"http://books.toscrape.com/catalogue/page-{page}.html"
        print(f"Crawling: {url}")
        
        try:
            res = requests.get(url, timeout=10)
            if res.status_code != 200:
                print(f"Failed to retrieve page {page}, status code: {res.status_code}")
                break
                
            soup = BeautifulSoup(res.text, "html.parser")
            articles = soup.find_all("article", class_="product_pod")
            if not articles:
                break
                
            for article in articles:
                if len(books) >= target_count:
                    break
                    
                try:
                    h3 = article.find("h3")
                    a = h3.find("a")
                    title = a.get("title") or a.text.strip()
                    
                    rating_p = article.find("p", class_="star-rating")
                    rating_classes = rating_p.get("class") if rating_p else []
                    rating_word = "Three"
                    for c in rating_classes:
                        if c != "star-rating":
                            rating_word = c
                            break
                    rating_num = rating_map.get(rating_word.lower(), 3)
                    
                    price_p = article.find("p", class_="price_color")
                    price_text = price_p.text if price_p else "£0.00"
                    price_float = float(price_text.replace("£", "").replace("Â", "").strip())
                    
                    avail_p = article.find("p", class_="instock availability")
                    avail_text = avail_p.text.strip() if avail_p else "In stock"
                    
                    href = a.get("href")
                    link = "http://books.toscrape.com/catalogue/" + href.replace("../", "")
                    
                    books.append({
                        "title": title,
                        "rating_word": rating_word,
                        "rating_num": rating_num,
                        "price": price_float,
                        "availability": avail_text,
                        "link": link
                    })
                except Exception as e:
                    print(f"Error parsing article: {e}")
            page += 1
            time.sleep(0.3)
            
        except requests.RequestException as e:
            print(f"Request failed: {e}")
            break
            
    print(f"Scraped {len(books)} books successfully.")
    return books

if __name__ == "__main__":
    scraped_data = scrape_books(50)
    # Output first few to verify
    for book in scraped_data[:5]:
        print(book)

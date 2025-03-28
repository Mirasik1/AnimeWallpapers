import os
import requests
from bs4 import BeautifulSoup


class VideoDownloader:
    def __init__(self, url, filename):
        self.url = url
        self.filename = filename
        self.headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36"
        }

    def get_existing_size(self):
        return os.path.getsize(self.filename) if os.path.exists(self.filename) else 0

    def download(self):
        file_size = self.get_existing_size()
        self.headers["Range"] = f"bytes={file_size}-"

        with requests.get(self.url, headers=self.headers, stream=True) as response:
            if response.status_code in (200, 206):
                with open(self.filename, "ab") as file:
                    for chunk in response.iter_content(chunk_size=1024):
                        if chunk:
                            file.write(chunk)
                print(f"Видео скачано как {self.filename}")
            else:
                print(f"Ошибка скачивания: {response.status_code}")


class VideoScraper:
    def __init__(self, page_url):
        self.page_url = page_url
        self.headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36"
        }

    def get_video_sources(self):
        response = requests.get(self.page_url, headers=self.headers)
        if response.status_code != 200:
            print(f"Ошибка {response.status_code}")
            return []

        soup = BeautifulSoup(response.text, "html.parser")
        video_tag = soup.find("video", id="my-player")
        if not video_tag:
            print("Видео не найдено")
            return []

        sources = [
            (source.get("res"), source.get("src"))
            for source in video_tag.find_all("source")
            if source.get("src")
        ]
        return sources


if __name__ == "__main__":
    scraper = VideoScraper("https://jut.su/dr-stoune/season-4/episode-2.html")
    video_links = scraper.get_video_sources()

    if video_links:
        best_video = max(video_links, key=lambda x: int(x[0] or 0))
        downloader = VideoDownloader(best_video[1], "video.mp4")
        downloader.download()

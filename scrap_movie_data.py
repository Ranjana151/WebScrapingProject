import requests
from bs4 import BeautifulSoup
import openpyxl

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = "Top Rated Movie"
sheet.append(["Movie Rank", "Movie Name", "Release Year", "IMDB Rating"])

url = "https://www.imdb.com/chart/top/"
r = requests.get(url).text

soup = BeautifulSoup(r, 'html.parser')
movies = soup.find('tbody', class_="lister-list").find_all('tr')
for movie in movies:
    movie_rank = movie.find('td', class_="titleColumn").get_text(strip=True).split('.')[0]
    movie_name = movie.find('td', class_="titleColumn").a.text
    movie_release_year = movie.find('td', class_="titleColumn").span.text.strip('()')
    movie_rating = movie.find('td', class_="ratingColumn imdbRating").strong.text
    """print(
        f"Movie Rank:{movie_rank} Movie Name:{movie_name} Movie Year :{movie_release_year} Movie Rating {movie_rating}")"""
    sheet.append([movie_rank, movie_name, movie_release_year, movie_rating])

excel.save("IMDB MOVIES RATINGS.xlsx")

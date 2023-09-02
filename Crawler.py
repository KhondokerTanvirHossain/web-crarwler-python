import requests
from bs4 import BeautifulSoup
from xlwt import *


def func_movie_content(movie_soup):
    movie_content_soup = movie_soup.find('p', {
        'data-qa': 'movie-info-synopsis'
    })
    if movie_content_soup is None:
        return "404 NOT FOUND!"
    else:
        return movie_content_soup.getText().strip()


def func_movie_review(movie_soup):
    movie_review_soup = movie_soup.find('span', {
        'data-qa': 'critics-consensus'
    })
    if movie_review_soup is None:
        return "404 NOT FOUND!"
    else:
        return movie_review_soup.getText().strip()


def func_movie_category(movie_soup):
    movie_category_soup = movie_soup.find('p', {
        'class': 'info',
        'data-qa': 'score-panel-subtitle'
    })
    if movie_category_soup is None:
        return "404 NOT FOUND!"
    else:
        return movie_category_soup.getText().strip()


if __name__ == '__main__':
    url = "https://www.rottentomatoes.com/top/bestofrt/"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.132 Safari/537.36 QIHU 360SE'
    }
    f = requests.get(url, headers=headers)
    movies_lst = []
    soup = BeautifulSoup(f.content, 'lxml')
    blocks = soup.find_all('span', {
        'class': 'p--small'
    })
    num = 0
    line = 1
    workbook = Workbook(encoding='utf-8')
    table = workbook.add_sheet('data')
    table.write(0, 0, 'Number')
    table.write(0, 1, 'Name')
    table.write(0, 2, 'Url')
    table.write(0, 3, 'Information')
    table.write(0, 4, 'Introduction')
    table.write(0, 5, 'Review')
    for block in blocks:
        movie_name = block.getText().strip()
        movie_url = 'https://www.rottentomatoes.com/m/' + movie_name.replace(' ', '_').lower()
        # print(movie_url)
        # movies_lst.append(movie_url)
        movie_f = requests.get(movie_url, headers=headers)
        movie_soup = BeautifulSoup(movie_f.content, 'lxml')
        movie_content = func_movie_content(movie_soup)
        movie_review = func_movie_review(movie_soup)
        movie_category = func_movie_category(movie_soup)
        # print(movie_content)
        table.write(line, 0, num)
        table.write(line, 1, movie_name)
        table.write(line, 2, movie_url)
        table.write(line, 3, movie_category)
        table.write(line, 4, movie_content)
        table.write(line, 5, movie_review)
        line += 1
        num += 1
    workbook.save('movies_top100.xls')

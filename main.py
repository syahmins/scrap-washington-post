import requests
from bs4 import BeautifulSoup
import xlsxwriter
from datetime import datetime

# tautan yang akan diambil datanya
html_doc = requests.get('https://www.washingtonpost.com/global-opinions/?itid=sf_opinions_subnav')

# tampilkan isi halaman dalam format html
soup = BeautifulSoup(html_doc.text, 'html5lib')


def buat_excel():
    # Buat judul kolom
    judul_kolom = ['No.', 'Title', 'Author', 'Summary', 'URL']
    row = 3

    # Buat berkas baru .xlsx dan worksheet baru di dalamnya
    workbook = xlsxwriter.Workbook('opinions_washingtonpost.xlsx')
    worksheet = workbook.add_worksheet()

    # Atur lebar kolom di Excel
    worksheet.set_column(0, 0, 5)  # Nomor urut
    worksheet.set_column(1, 1, 60)  # Title
    worksheet.set_column(2, 2, 25)  # Author name
    worksheet.set_column(3, 3, 60)  # Summary
    worksheet.set_column(4, 4, 30)  # URL

    # buat judul kolom di Excel
    for col, title in enumerate(judul_kolom):
        worksheet.write(row, col, title)

    # data untuk kolom nomor, judul dan url
    list_url = soup.find_all('a', attrs={'data-pb-field': 'headlines.basic'})
    row = 3

    # tanggal dan waktu data diambil
    waktu = 'Scraping date/time: ' + str(datetime.now())
    worksheet.write(0, 0, waktu)

    for url in list_url:
        row += 1

        worksheet.write(row, 0, row)
        worksheet.write(row, 1, url.text)
        worksheet.write(row, 4, url.get('href'))

    # data untuk kolom penulis
    penulis_all = soup.find_all('span', class_={'author vcard'})
    row = 3  # reset pengaturan baris

    for penulis in penulis_all:
        row += 1
        worksheet.write(row, 2, penulis.text)

    # data untuk kolom ringkasan
    ringkasan_all = soup.find_all('div', class_={'blurb normal normal-style'})

    row = 3  # reset pengaturan baris
    for ringkasan in ringkasan_all:
        row += 1
        worksheet.write(row, 3, ringkasan.text)
        # isi_ringkasan = ringkasan.text

    workbook.close()

    print('Data sudah selesai diekspor ke Excel...')


if __name__ == '__main__':
    buat_excel()

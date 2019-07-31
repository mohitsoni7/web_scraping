import io
import requests
import xlsxwriter
from bs4 import BeautifulSoup
from pprint import pprint

if __name__ == '__main__':
    # url = 'https://www.youtube.com/playlist?list=PL-osiE80TeTtoQCKZ03TU5fNfx2UY6U4p'
    # url = 'https://www.youtube.com/playlist?list=PL-osiE80TeTskrapNbzXhwoFUiLCjGgY7'
    url = 'https://www.youtube.com/playlist?list=PL-osiE80TeTt2d9bfVyTiXJA-UTHn6WwU'
    url = 'https://www.youtube.com/playlist?list=PL-osiE80TeTt2d9bfVyTiXJA-UTHn6WwU'
    source = requests.get(url).text
    soup = BeautifulSoup(source, 'lxml')

    tutorials_title = soup.find('h1', class_='pl-header-title').text.strip()
    content_div = soup.body.find('div', id='content')


    excel_data = []
    max_width_title = 0
    max_width_link = 0

    # video_td = content_div.find('td', class_='pl-video-title')
    all_video_td = content_div.find_all('td', class_='pl-video-title')
    for idx, item in enumerate(all_video_td, 1):
        video_link = item.a.get('href')
        video_link = video_link.split('&')[0]
        video_link = f'https://youtube.com{video_link}'
        video_title = item.text.strip().split('\n')[0]
        
        if len(video_title) > max_width_title:
            max_width_title = len(video_title)

        if len(video_link) > max_width_link:
            max_width_link = len(video_link)
        
        data = [idx, video_title, video_link]
        excel_data.append(data)
    
    # print('======================')
    # pprint(excel_data)
    # print('======================')

    # Create a workbook
    tutorials_title = '_'.join(tutorials_title.lower().split(' '))
    file_path = f'/home/mohit/Documents/learning/web_scraping/{tutorials_title}.xlsx'
    workbook = xlsxwriter.Workbook(file_path)

    heading_format = workbook.add_format({
                'align': 'center',
                'valign': 'vcenter',
                'bold': True})
    center_format = workbook.add_format({
                'align': 'center',
                'valign': 'vcenter'})
    left_format = workbook.add_format({
                'align': 'left',
                'valign': 'vcenter'})
    
    worksheet = workbook.add_worksheet()

    worksheet.merge_range('A1:D1', f"{tutorials_title.title()}", heading_format)
    worksheet.write('A2', 'S.No.', heading_format)
    worksheet.write('B2', 'Competency', heading_format)
    worksheet.write('C2', 'Source', heading_format)
    worksheet.write('D2', 'Duration', heading_format)

    worksheet.set_column('B:B', max_width_title)
    worksheet.set_column('C:C', max_width_link)

    row = 2
    col = 0

    for serial_no, competency, source in excel_data:
        worksheet.write(row, col, serial_no, center_format)
        worksheet.write(row, col+1, competency, left_format)
        worksheet.write(row, col+2, source, left_format)
        worksheet.write(row, col+3, '', left_format)
        row += 1
    
    # Close workbook
    workbook.close()
    print(f'Excel sheet created at: {file_path}')
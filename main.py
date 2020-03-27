from pdf2image import convert_from_path
import requests
import os
from sys import platform

if platform == "linux" or platform == "linux2":
    pass
    path = 'files/'
elif platform == "win32":
    poppler = r'C:\Users\Violet\PycharmProjects\gs-to-jpg-send-tlgrm\poppler\bin'
    path = r'C:\Users\Violet\PycharmProjects\gs-to-jpg-send-tlgrm\files\\'

# Google Spreadsheet
# Param export GS
# &format=pdf                   //export format
# &size=a4                      //A3/A4/A5/B4/B5/letter/tabloid/legal/statement/executive/folio
# &portrait=false               //true= Potrait / false= Landscape
# &scale=1                      //1= Normal 100% / 2= Fit to width / 3= Fit to height / 4= Fit to Page
# &top_margin=0.00              //All four margins must be set!
# &bottom_margin=0.00           //All four margins must be set!
# &left_margin=0.00             //All four margins must be set!
# &right_margin=0.00            //All four margins must be set!
# &gridlines=false              //true/false
# &printnotes=false             //true/false
# &pageorder=2                  //1= Down, then over / 2= Over, then down
# &horizontal_alignment=CENTER  //LEFT/CENTER/RIGHT
# &vertical_alignment=TOP       //TOP/MIDDLE/BOTTOM
# &printtitle=false             //true/false
# &sheetnames=false             //true/false
# &fzr=false                    //true/false
# &fzc=false                    //true/false
# &attachment=false             //true/false
gs_id = '1QJRxp...gVVBIw'
gs_url = 'https://docs.google.com/spreadsheets/d/' + gs_id + '/export?format=pdf&gid='
gs_param = '&size=a5&portrait=false&vertical_alignment=MIDDLE&horizontal_alignment=CENTER'
gs_table = [581871535, 1884134963, 1426883343]  # gid

# Telegram
bot = 'TOKEN'
proxy = {'https': 'socks5h://LOGIN:PASS@IP:1080'}


def main(gid):
    # Download Google Spreadsheet
    response = requests.get(gs_url + str(gid) + gs_param)
    with open(path + 'test-table.pdf', 'wb') as out_file:
        for chunk in response:
            out_file.write(chunk)

    # Google Spreadsheet pdf to jpg
    convert_from_path(path + 'test-table.pdf', poppler_path=poppler, output_folder=path,
                      output_file='test-table', fmt='jpg')

    # Send Google spreadsheet jpg send from bot
    requests.post('https://api.telegram.org/bot' + bot + '/sendPhoto',
                  files={'photo': open(path + 'test-table0001-1.jpg', 'rb')},
                  data={'chat_id': 'id_tlgrm'}, proxies=proxy)


for gid in gs_table:
    main(gid)

os.remove(path + 'test-table.pdf')
os.remove(path + 'test-table0001-1.jpg')

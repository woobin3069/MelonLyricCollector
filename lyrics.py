import win32com.client as win32
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import numpy as np
import time
from selenium.common.exceptions import NoSuchElementException

#Get lyrics
driver = webdriver.Chrome('C:/Users/USER/Downloads/chromedriver_win32/chromedriver.exe')
driver.get("https://www.melon.com/artist/song.htm?artistId=672375#params%5BlistType%5D=A&params%5BorderBy%5D=ISSUE_DATE&params%5BartistId%5D=672375&po=pageObj&startIndex=1")

fail = []
lyric = []


def check_exists_by_link_class(text, num):
    try:
        driver.find_element_by_class_name(text)
    except NoSuchElementException:
        fail.append(num)
        return False
    return True

time.sleep(5)
lyric_link = []
details = driver.find_elements_by_class_name("btn_icon_detail")
print(len(details))

for i in range(50):
    lyric_link.append(details[i].get_attribute("href"))
    lyric_link[i] = lyric_link[i].split("'")
    del lyric_link[i][0]
    del lyric_link[i][1]
    lyric_link[i] = "".join(lyric_link[i])

title = []
lyrics = []
for i in range(50):
    driver.get("https://www.melon.com/song/detail.htm?songId=" + lyric_link[i])
    if check_exists_by_link_class("button_more", i):
        song_name = driver.find_element_by_class_name("song_name")
        title.append(song_name.text)
        more_btn = driver.find_element_by_class_name("button_more")
        more_btn.click()
        lyric_song = driver.find_element_by_class_name("lyric")
        lyrics.append(lyric_song.text)

for i in range(len(lyrics)):
    lyric.append(lyrics[i].split('\n'))
print(lyric)



#Generate hwp document
hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
hwp.Open('e:/BTSlyrics/bts_lyrics.hwp', "HWP", None)

field_list = [i for i in hwp.GetFieldList().split('')]
hwp.Run('SelectAll')
hwp.Run('Copy')
hwp.MovePos(3)

for i in range(len(lyric)):
    hwp.Run('Paste')
    hwp.MovePos(3)

print(field_list)

for page in range(len(lyric)):
    if page in fail:
        continue
    lyric[page] = "[NEXT_LINE]".join(lyric[page])
    hwp.PutFieldText(f'{field_list[0]}{{{{{page}}}}}',
                     title[page])
    hwp.PutFieldText(f'{field_list[1]}{{{{{page}}}}}',
                     lyric[page])
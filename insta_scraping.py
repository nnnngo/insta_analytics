import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
import re
import requests
from typing import Dict
from bs4 import BeautifulSoup
import datetime

class SocialMedias():

    @classmethod
    def get_statuses(cls, op) -> Dict:
        statuses = {}
        if op == 1:
            print("取得したい投稿のurlを入力してください")
            url = input()
            statuses = cls.__get_instagram_post_statuses(url)
            return statuses
        elif op == 2:
            print("取得したいユーザーのurlを入力してください")
            url = input()
            statuses = cls.__get_instagram_statuses(url)
            statuses['url'] = url
            return statuses

    @classmethod
    def __get_instagram_post_statuses(cls, url) -> Dict:
        statuses = {}
        response = requests.get(url)
        soup = BeautifulSoup(response.content, 'lxml')
        js = soup.find("script", text=re.compile("window._sharedData")).text
        data = json.loads(js[js.find("{"):js.rfind("}")+1]);
        statuses['url'] = url
        statuses['like'] = data['entry_data']['PostPage'][0]['graphql']['shortcode_media']['edge_media_preview_like']['count']
        statuses['time'] = datetime.datetime.fromtimestamp(int(data['entry_data']['PostPage'][0]['graphql']['shortcode_media']['taken_at_timestamp'])).strftime('%Y/%m/%d')
        return statuses

    @classmethod
    def __get_instagram_statuses(cls, url: str) -> Dict:
        statuses = {}

        # ToDo: error handling
        response = requests.get(url)

        soup = BeautifulSoup(response.content, 'lxml')
        js = soup.find("script", text=re.compile("window._sharedData")).text
        data = json.loads(js[js.find("{"):js.rfind("}")+1]);
        print(data['entry_data']['ProfilePage'][0]['graphql']['user']['username'])
        statuses['name'] = data['entry_data']['ProfilePage'][0]['graphql']['user']['username']
        statuses['address'] = data['entry_data']['ProfilePage'][0]['graphql']['user']['business_address_json']
        statuses['following_count'] = data['entry_data']['ProfilePage'][0]['graphql']['user']['edge_follow']['count']
        statuses['follower_count'] = data['entry_data']['ProfilePage'][0]['graphql']['user']['edge_followed_by']['count']
        edge = data['entry_data']['ProfilePage'][0]['graphql']['user']['edge_owner_to_timeline_media']['edges']
        statuses['engagement'] = 0
        for i in range(0, 5):
            statuses['engagement'] += edge[i]['node']['edge_media_to_comment']['count'] + edge[i]['node']['edge_liked_by']['count']
        if i!= 0:
            statuses['engagement'] /= i + 1
            statuses['engagement'] /= statuses['follower_count']
            statuses['engagement'] = round(statuses['engagement'] * 100,2)
        return statuses

    @staticmethod
    def __get_element_by_class(soup: BeautifulSoup, class_name: str) -> BeautifulSoup:
        return soup.find(attrs={'class': re.compile('^' + class_name + '$')})

def main():
    social_medias = SocialMedias()
    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']
    print("jsonファイル名を入力してください。(拡張子不要)")
    jsonfile = input()
    credentials = ServiceAccountCredentials.from_json_keyfile_name(jsonfile + ".json", scope)
    gc = gspread.authorize(credentials)
    print("開きたいシート名を入力してください")
    gcname = input()
    wks = gc.open(gcname).sheet1
    print("1.投稿　2.ユーザー")
    op = int(input())
    if op == 1:
        print("書き込みたい投稿数を入力してください")
        num = int(input())
        num2 = 0
        cell_list = wks.range('B1:B100')
        for cell in cell_list:
            if len(cell.value) > 0:#セルに何か書き込まれている場合
                num2 += 1
            else:
                break
        for i in range(2, num+2):
            i += num2 - 1
            instaSt = social_medias.get_statuses(op)
            wks.update_acell('D'+ str(i), instaSt['like'])
            wks.update_acell('B'+ str(i), instaSt['url'])
            wks.update_acell('C'+ str(i), instaSt['time'])
            i -= num2 - 1
    elif op == 2:
        print("書き込むインフルエンサーの人数を入力してください")
        num = int(input())
        num2 = 0
        cell_list = wks.range('B1:B100')
        for cell in cell_list:
            if len(cell.value) > 0:#セルに何か書き込まれている場合
                num2 += 1
            else:
                break
        for i in range(2, num + 2):
            i += num2 - 1
            instaSt = social_medias.get_statuses(op)
            wks.update_acell('D'+ str(i), instaSt['follower_count'])
            wks.update_acell('C'+ str(i), instaSt['url'])
            wks.update_acell('B'+ str(i), instaSt['name'])
            wks.update_acell('F'+ str(i), instaSt['address'])
            wks.update_acell('E'+ str(i), instaSt['engagement'])
            i -= num2 - 1

if __name__ == "__main__":
    main()

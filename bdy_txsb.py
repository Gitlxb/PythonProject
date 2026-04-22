# -- coding: utf-8 --
# @Time : 2025-05-16 15:55
# @Author : 贝特利
# @Email : 1356087739@qq.com
# @File : bdy_txsb.py
# @Software: PyCharm


import requests
import json


def main():
    url = "https://aip.baidubce.com/oauth/2.0/token?grant_type=client_credentials&client_id=AROBDhdOUPjJ0tXImXEaNR6m&client_secret=D7Z9nWVeNnXQAeixFQ3ecdvBYpO9Jj9m"

    payload = ""
    headers = {
        'Content-Type': 'application/json',
        'Accept': 'application/json'
    }

    response = requests.request("POST", url, headers=headers, data=payload)

    print(response.text)


if __name__ == '__main__':
    main()
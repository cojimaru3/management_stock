#!/usr/bin/env python
# -*- coding: utf-8 -*-

import requests

api_token = 'xoxb-1784983863632-1774938969443-91OiJqQFfoLVrlPUE7Uz7YVv'
channel = 'dividend_stocks_channel'

class SlackManager:
    def __init__(self):
        self.api_token = api_token
        self.channel = channel

    def upload_file(self,file,file_name):
        files = {'file':open(file,'rb')}
        param = {
            'token' : self.api_token,
            'channels' : self.channel,
            'filename' : file_name,
            'initial_comment' : 'file upload',
            'title' : 'dividend stocks'
        }
        res = requests.post(url="https://slack.com/api/files.upload",
            params=param,
            files=files)
        print(res.json())
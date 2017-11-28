#!/usr/bin/python
"""
今目标网页端的相关操作

它可以实现以下功能：
* 自动写日志，并可以上传附件(可选)
* 写好的日志，可以添加到主线上
* 下载所有的主线，并包括其中的附件
* 自动阅读权限下所有员工的日志内容
* 通过网页端对客户端发送信息
* 自动签到
* 向今目标客户端发送消息

以上功能，可以配合windows的定时任务完成

Email: sbmzhcn@gmail.com
"""

import re
import time
import random
import math
import os
import mimetypes
import datetime
import json
import ctypes
from urllib.parse import urljoin, unquote, parse_qs
import unittest
import doctest
import requests


class JinGoal():
    base_url = "https://web.jingoal.com/"
    login_url = "https://sso.jingoal.com/oauth/authorize"

    def __init__(self, username, password, sha_pass=None):
        self.username = username
        self.password = password
        self.sesstion = requests.session()
        self.user_id = None
        self.company_id = None
        self.company_code = None
        self.sha_passwd = self.sha1(password)
        if sha_pass:
            self.sha_passwd = sha_pass
        self.login()

    def request(self, http_method, path, **kwargs):
        response = self.sesstion.request(
            http_method,
            urljoin(self.base_url, path), **kwargs)
        return response

    def _get_user_config(self, code, params=None):
        self.sesstion.headers['Content-Type'] = 'application/json'
        if code:
            params = {
                "code": code,
                "ouri": self.base_url,
            }
        result = self.request(
            "GET", "mgt/workbench/v1/get_userconfig", params=params)
        return result.json()

    def _login(self):
        """
        >>> print("test doctest")
        test doctest
        """
        self.sesstion = requests.session()
        logindata = {
            'password': self.sha_passwd,
            'login_type': 'default',
            'username': self.username,
            'identify': '',
        }
        params = {
            "client_id": "jmbmgtweb",
            "response_type": "code",
            "state": "{access_count:1}",
            "locale": "zh_CN",
            "redirect_uri": self.base_url,
        }
        print("Start trying login systems......")
        response = self.sesstion.post(self.login_url, params=params,
                                      data=logindata, allow_redirects=False)  # {"ok":true}
        result = response.json()
        # {'showIdentifyCode': False, 'error': 'BadCredentials'}
        if 'error' in result and result["error"] == "NeedIdentifyCode":
            print("!!! Login Failed !!!, Need Identify Code")
            return False
        elif "ok" in result:
            print("Login success, but need check it again.")
        else:
            print("Login system occurr an unexpected error: {0}".format(result))
            return False
        response = self.sesstion.get(self.login_url, params=params, allow_redirects=False)
        code = re.search('code=(.+?)&', response.headers["Location"]).group(1)
        return self.check_login_status(code)

    def check_login_status(self, code=None):
        user_config = self._get_user_config(code)
        if user_config["meta"]["code"] != 200:
            return False
        else:
            self.user_id = user_config["data"]["user_id"]
            encrypt_info = parse_qs(unquote(user_config["data"]["encrypt_info"]))
            print(encrypt_info)
            self.company_id = encrypt_info["memo"][0].split(",")[0]
            self.company_code = encrypt_info["memo"][0].split(",")[1]
            login_info = """{company_name}\nuser_name: {user_name}\nduty_name: {duty_name}\ndept_name: {dept_name}
\nuser_id: {user_id}\ncompany_id: {company_id}\n公司代码: {company_code}"""
            print("=======以下是你的登陆信息=======")
            print(login_info.format(**user_config["data"], company_id=self.company_id, company_code=self.company_code))
            print("========登陆信息显示结束========")
        with open("session.txt", 'w') as file_stream:
            file_stream.write(json.dumps(requests.utils.dict_from_cookiejar(self.sesstion.cookies)))
        return user_config

    def get(self, url, params=None, stream=False):
        return self.sesstion.get(url, params=params, stream=stream)

    def is_logged(self):
        if self.check_login_status() is False:
            return self._login()
        return True

    def login(self):
        try:
            if os.path.exists("session.txt"):
                with open("session.txt") as f:
                    cookie = f.read()
                    if cookie:
                        self.sesstion.cookies = requests.utils.cookiejar_from_dict(
                            json.loads(cookie))
            return self.is_logged()
        except Exception as e:
            print("Login exception: {}".format(e))
            return False

    def write_segment(self, path, content=None, attachs=None):
        """
        return {"id":145399864,"ok":true}
        """
        response = self.request("POST", path, json={
            "isAttach": True if attachs else False,
            "content": content,
            "attachs": attachs if attachs else ""
        }, allow_redirects=False, headers={"Content-Type": "application/json"})

        try:
            return response.json()
        except:
            return False

    def edit_segment(self, path, content=None, logid=None):
        data = {
            "content": content,
            "id": "0",
            "type": "2",
            "workLogId": logid
        }
        response = self.request("POST", path, json=data, allow_redirects=False)
        try:
            return response.json()
        except:
            return False

    def edit_work_record(self, content="", w_id="", logid=""):
        return self.request("POST", 'module/worklog/modify/saveLogSegment.do', json={
            "content": content,
            "id": w_id,
            "type": "2",
            "workLogId": logid
        })

    def write_worklog(self, content=None, files=None):
        """
        新建工作小结

        :content: 日志内容，可以是html格式.
        :files: 数组，包含文件路径.
        :return: {"id":145398715,"ok":true}
        """
        attachs = files = []
        for i, f in enumerate(files):
            try:
                filename, local_path, status, size = self.upload_file(f)
                attachs.append({
                    "id": i,
                    "deleted": False,
                    "fileName": filename,
                    "localPath": local_path,
                    "size": int(size)
                })
                print("Upload file {0} success: {1}....".format(
                    filename, local_path))
            except Exception as e:
                print("error when uploading file {0}, Error: {1}".format(f, e))
        return self.write_segment("module/worklog/editEndSegment.do", content, attachs=attachs)

    def write_log(self, content=""):
        return self.write_segment("module/worklog/editEndSegment.do", content)

    def edit_log(self, content=""):
        """
        编辑工作记录
        """
        return self.edit_segment("module/worklog/modify/saveLogSegment.do", content)

    def upload_file(self, file_path):
        """
        上传文件
        """
        from requests_toolbelt import MultipartEncoder
        idoo = self.mgtuid()
        filename = os.path.basename(file_path)
        content_type = mimetypes.guess_type(
            filename)[0] or 'application/octet-stream'
        m_data = MultipartEncoder(
            fields={
                "filename": "test.xlsx",
                'file': ("test.xlsx", open(file_path, 'rb'), content_type)
            }
        )
        response = self.request("POST", "module/mgtfileupload/post.do",
                                params={"idoo": idoo, "X-Progress-ID": idoo},
                                data=m_data,
                                headers={'Content-Type': m_data.content_type})
        result_group = re.search(
            r"parent.mgtFileuploadPage.setResult\(\"(.*?)\",(\w+),(\d+)\);", response.text)
        return filename, result_group.group(1), result_group.group(2), result_group.group(3)

    def reset_today_log(self):
        """
        清除当天日志中的附件
        当天日志内容设置为空，无法删除日志
        """
        response = self.request("GET", 'module/worklog/workLogInfo.do')
        # attachs = re.findall(r'showAttach_(\d+)', response.text)
        segments = re.findall(r'editSegment\(\'(\d{8,})\',\'(\d*)\',\'(\d)\'', response.text)
        current_worklog_id = "0"

        for log_id, seg_id, seg_type in segments:
            current_worklog_id = log_id
            self.delete_attach_byid(log_id)
            if seg_type == "3":
                self.write_worklog(content='')
            if seg_type == "2":
                # 如果有工作记录，只能让它内容为空
                self.edit_work_record(content="", w_id=seg_id, logid=log_id)
        return current_worklog_id

    def delete_attach_byid(self, file_id):
        return self.request("POST", "module/attach/{}/delete.do".format(file_id)).json()["ok"]

    def list_my_worklog(self):
        response = self.request("POST", "module/worklog/workLogToList4Igoal.do", data={
            "key": "",
            "listType": "myLog",
            "currentPage": 1
        }).text
        log_list = re.findall('<input.*?name="workLog".*?data-content="(.*?)"', response, re.MULTILINE | re.DOTALL)
        return log_list

    def mgtuid(self, user_id=None):
        user_id = self.user_id if self.user_id else None
        s_result = user_id if user_id else "kp"
        result = "{0}.{1:.0f}".format(s_result, time.time() * 1000)
        for _ in range(32):
            result += format(math.floor(random.random() * 16), "x").upper()
        return result

    def get_igoal_list(self):
        response = self.request("POST", "igoal/igoal/list.do", params={
            "ownerId": 0,
            "status": -1,
            "tagId": -1,
            "key": "",
            "deleted": False,
            "enc": 0,
            "currPage": 1
        }, headers={
            "Content-Type": "application/json",
            "X-Requested-With": "XMLHttpRequest"
        })
        return re.findall(r"viewIgoal\('(\d+)',0\)", response.text)

    def print_igoal_list(self):
        for igoal_id in self.get_igoal_list():
            frames = self.igoal_detail(igoal_id)
            title = frames["igoal"]["tagName"] + "-" + frames["igoal"]["igoal"]["name"]
            print(title, igoal_id, frames)

    def dowload_all_igoal(self):
        """
        下载所有的主线，并保存到本地
        """
        for igoal_id in self.get_igoal_list():
            frames = self.igoal_detail(igoal_id)
            title = frames["igoal"]["tagName"] + "-" + frames["igoal"]["igoal"]["name"]
            self.download_frames(frames, parent=title)

    def download_frames(self, frames, parent=""):
        for f in frames["frames"]:
            if "hasChilds" in f and f["hasChilds"] is True:
                dir_path = os.path.dirname(__file__) + "/dowload/" + parent + "/" + f["igoalFrame"]["title"] + "/"
                if not os.path.exists(dir_path):
                    os.makedirs(dir_path)
                childs = self.igoal_childs(f["id"])
                for item in childs["items"]:
                    item_detail = self.igoal_detail_item(item["id"])
                    source_url = re.search('src="(.*?)"', item_detail)
                    src_url = re.search('src=".*?url=(.*?)&[^"]*token=(.*?)"', item_detail)
                    if item["referType"] == "plan":
                        detail_url = "/webcd/plan/planDetail4Igoal.do"
                    elif item["referType"] == "task":
                        detail_url = "/webcd/task/taskDetail4Igoal.do"
                    elif item["referType"] == "log":
                        detail_url = "/module/worklog/logDetail4IgoalIframe.do"
                    elif item["referType"] == "doc":
                        detail_url = source_url.group(1)
                    elif item["referType"] == "memo":
                        detail_url = "/webxm/memo/memoDetail4IgoalRefer.do"

                    if item["referType"] != "doc":
                        r = self.request("GET", detail_url, params={"token": src_url.group(2), "url": src_url.group(1)})
                        print(item["item"]["title"])
                        file_title = item["ownerName"] + "-" + item["typeName"] + "-" + item["item"]["title"] + ".html"
                        file_title = re.sub(r'[^-a-zA-Z0-9\u4e00-\u9fff_.()/ ]+', '', file_title)
                        with open(dir_path + file_title, "w", encoding="utf-8") as html_file:
                            html_file.write(r.text.strip())
                        # 如果日志或者其它类型中有附件，下载之
                        files = re.findall('<a[^>]*title="(.+?)"[^>]*href=".+?(url=token&rid=.*?&mgtToken=.*?)&', r.text)
                        for attach_filename, attach_url in files:
                            if not os.path.exists(dir_path + attach_filename.strip()):
                                print("下载附件...", attach_filename, attach_url)
                                attach_steam = self.request("GET", "webcd/resources/?" + attach_url, stream=True)
                                with open(dir_path + attach_filename.strip(), 'wb') as fs:
                                    for chunk in attach_steam:
                                        if chunk:
                                            fs.write(chunk)
                    elif item["referType"] == "doc":
                        itemRefer = json.loads(item["item"]["itemRefer"])
                        payload_text = """callCount=1
page={page}
httpSessionId=
scriptSessionId=${{scriptSessionId}}229
c0-scriptName=dwrService
c0-methodName=getDWRIncludeURL
c0-id=0
c0-e1=number:{userid}
c0-param0=Object_Object:{{id:reference:c0-e1}}
c0-e2=string:%2Fmydoc%2Fsysdoc%2FiGoal%2FshowAttachDetail.jsp%3Ffile%3Dfalse%26url%3D{url}%26token%3D{token}%26online_ids_stamp%3D0
c0-param1=Array:[reference:c0-e2]
batchId=0""".format(page=detail_url, token=re.search('token=(.+)', detail_url).group(1), url=itemRefer["url"], userid=self.user_id)
                        rs = self.request("POST", "mgt2/dwr/call/plaincall/dwrService.getDWRIncludeURL.dwr?t=&uid={0}&                      cid=undefined".format(self.user_id), data=payload_text, headers={"Content-Type": "text/plain", "X-Requested-With": 'XMLHttpRequest'})
                        json_data = re.search(r's1\[0\]="(.+)"', rs.text)
                        parsed_text = json_data.group(1).encode('utf-8').decode('unicode_escape')
                        try:
                            download_url = re.search('href="(/mgt2/mydoc/sysdoc/iGoal/download.jsp.*?)"', parsed_text).group(1)
                            file_name = re.search('<a class="black01"[^>]*title="(.*?)"', parsed_text, re.MULTILINE).group(1)
                            if not os.path.exists(dir_path + file_name):
                                # 如果文件没有下载过，再下载
                                print("Dowloading file: {}".format(file_name))
                                steam = self.request("GET", download_url, stream=True)
                                with open(dir_path + file_name, 'wb') as fs:
                                    for chunk in steam:
                                        if chunk:
                                            fs.write(chunk)
                            else:
                                print("File {0} is exists, skip dowload.".format(file_name))
                        except AttributeError as e:
                            print("AttributeError", e, item)
                if "frames" in childs:
                    self.download_frames(childs, parent=parent + "/" + f["igoalFrame"]["title"] + "/")

    def add_igoal(self, log_id, **kwargs):
        """把日志添加到主线上 """
        self._login()
        # 把工作日志添加到主线上面
        # 2754921,2754796,2696046,2746926,2746913,2754850,2701856,3242493,2754906,2754834,2754807,2754869
        # 冯建会, 马晓东,  冯磊,    张鹏90, 王庆玲,  时鹏亮,  孙冰,   欧本林,   张扬,  万通,    陈忠营, 王翔
        # url = "https://web.jingoal.com/igoal/igoal/item/add.do"
        url_id = "RA1B1_{0}".format(log_id)
        token = self.check_right_igoal(url_id)
        frameId = kwargs["frameId"] if "frameId" in kwargs else 8090098
        igoalId = kwargs["igoalId"] if "igoalId" in kwargs else 1179576
        title = igoalId = kwargs["title"] if "title" in kwargs else "{0} 工作小结".format(
            datetime.datetime.now().strftime("%Y-%m-%d"))
        data = {
            "itemsData": [{
                "token": token,
                "frameId": frameId,  # 对应 张雷团队
                "igoalId": igoalId,  # 对应 主线 id
                "type": "log",
                "params": {
                    "url": url_id,
                    "file": "false"
                },
                "referred": {
                    "userId": str(self.user_id),
                    "date": ""
                },
                "title": title,
                "content": ""
            }],
            "affair": [False, False, False],
            "selectUserIds": "3242493,2701856"  # -1， 所有人, 欧本林，孙冰 3242493,2701856
        }
        response = self.request("POST", "igoal/igoal/item/add.do", json=data)
        return response.json()

    def delete_igoal(self, item_id):
        self.request("POST", "igoal/igoal/item/delete.do", data={"itemId": item_id}, headers={
            "Content-Type": "application/x-www-form-urlencoded"})
        return True

    def reset_today_igoal(self, log_id=None, frame_id=8090098):
        items = self.list_igoal_item(frame_id)
        for item in items:
            igoal_id = item["id"]
            item_refer = item["item"]["itemRefer"]
            create_time = datetime.datetime.fromtimestamp(
                item["item"]["createTime"] / 1e3)
            logid = re.search(r'(\d{4,})', item_refer).group(1)
            if log_id and int(log_id) == int(logid):
                self.delete_igoal(igoal_id)
            elif create_time.date() == datetime.datetime.now().date():
                self.delete_igoal(igoal_id)

    def list_igoal_item(self, frame_id=8090098):
        response = self.request("POST", "igoal/igoal/frame/childs.do", data={"frameId": frame_id}, headers={
            "Content-Type": "application/x-www-form-urlencoded"})
        print(response.text)
        return response.json()["items"]

    def igoal_detail(self, igoal_id):
        response = self.request("POST", "igoal/igoal/igoal/frames.do", data={"igoalId": igoal_id}, headers={
            "Content-Type": "application/x-www-form-urlencoded"})
        return response.json()

    def igoal_detail_item(self, item_id):
        return self.request("POST", "igoal/igoal/item/detail.do", data={"itemId": item_id}, headers={
            "Content-Type": "application/x-www-form-urlencoded"}).text

    def igoal_childs(self, frame_id):
        return self.request("POST", "igoal/igoal/frame/childs.do", data={"frameId": frame_id}, headers={
            "Content-Type": "application/x-www-form-urlencoded"}).json()

    def check_right_igoal(self, url_id):
        return self.request("POST", "module/worklog/checkRight4Igoal.do", data={"url": url_id, "file": "false"},
                            headers={'Content-Type': 'application/x-www-form-urlencoded'}).text.strip()

    def get_all_can_view_users(self):
        response = self.request("GET", "module/worklog/viewlog.do")
        return re.search(r"selectViewUsers\('(.*?)'\)", response.text).group(1).split(",")

    def get_all_users(self):
        """
        获取所有可查看的用户列表

        return:
            [(2755114, '小明', 'SEO二组'),]
        """
        response = self.request("GET", "module/tree/getTree.do", params={
            "treeType": "byDept",
            "permId": "logDeptcomment"
        })
        treeData = json.loads(response.text)
        children = treeData[0]["children"]
        all_users = []
        self.reserve_get_users(children, all_users)
        return all_users

    def reserve_get_users(self, children, all_users=None, parent=None):
        all_users = []
        # 递归获取所有用户
        for node in children:
            if node["isParent"]:
                self.reserve_get_users(
                    node["children"], all_users, node["name"])
            else:
                all_users.append((node["id"], node["name"], parent))

    def get_user_log(self, user_id, page=1):
        data = {
            "from": -1,
            "to": -1,
            "curr": page,
            "unread": True,
            "viewuserCurr": 1,
            "selectUserId": str(user_id),
            "fromViewLog": True,
            "type": -1
        }
        r = self.request("POST", "module/worklog/viewlogList.do", json=data)
        items_total = re.search(r"items_total:'(\d+)'", r.text).group(1)
        # all_id = re.search(r"viewlog.setListpn\('(.*?)','(\d+)'\);", r.text).group(1).split(",")
        # 获取所有未读的
        unread_id = []
        recent_logs = []
        matchs = re.findall(
            r'<tr[^>]*viewlog.detail\((\d+)[^>]*>[^>]*<td([^>]*)>', r.text, re.MULTILINE | re.DOTALL)
        recent_log_matchs = re.findall(
            r'<div>\W+<span class="bluea">([^>]*)</span>([^<]*).*?</div>\W+<div>', r.text, re.MULTILINE | re.DOTALL)
        for recent_log in recent_log_matchs:
            recent_logs.append((recent_log))
        for m in matchs:
            userid, is_bold = m
            if is_bold.strip():
                unread_id.append(userid)
        return items_total, unread_id, recent_logs

    def view_log(self, log_id):
        response = self.request("POST", "module/worklog/viewDetail.do", params={"workLogId": log_id})
        page_content = response.text.strip()
        date = re.search(r'<b class="p14 w">(.*?)</b>', page_content).group(1)
        # group = re.search(r'<span class="remark_9 bold p14">(.*?)</span>', page_content).group(1)
        contents = re.search(r'<td[^>]*mgtContent[^>]*>(.*?)</td>', page_content, re.MULTILINE | re.DOTALL)
        return date, contents.group(1).strip()

    def read_all_user_log(self):
        users = self.get_all_users()
        pad = " "
        for userid, name, group in users:
            print(
                "正读取 {userid} - {name} - {group} 的日志".format(userid=userid, name=name, group=group))
            unread_id = []
            items_total, unread_log_ids, recent_logs = self.get_user_log(
                userid)
            unread_id = unread_log_ids
            page_nums = math.ceil(int(items_total) / 15)
            # print(page_nums, unread_logs)
            if unread_log_ids:
                # 如果第一页有未读的，才遍历所有的日志
                for num in range(2, page_nums + 1):
                    items_total, unread_log_ids, recent_logs = self.get_user_log(
                        userid, page=num)
                    unread_id += unread_log_ids
                # 读取所有的未读日志
                print("{}发现未读的日志内容，正循环获取".format(pad * 2))

                def format_content(content):
                    # print(content)
                    format_content = ""
                    for line in content.split("<br>"):
                        if line.strip():
                            format_content += "{0}--{1}\n".format(
                                pad * 6, line)
                    return format_content
                for unid in unread_id:
                    date, content = self.view_log(unid)
                    print("{pad}{date}\n{content}".format(
                        date=date, content=format_content(content), pad=pad * 6))
                # break
            else:
                print("{}没有未读的日志内容".format(pad * 2))
                print("{}最近的日志内容概览:".format(pad * 4))
                for logs in recent_logs:
                    content = logs[1].strip().replace("\r", " ").replace("\n", " ")
                    print("{pad}{date}\n{pad}--{content}".format(date=logs[0],
                                                                 content=content, pad=pad * 6))

    def get_view_perm(self):
        # 获取有权限查看你日志的人员名单
        response = self.request("POST", "module/worklog/viewperm.do")
        perm_user = re.search(r'<div class="modal-body">\W+<div>([^>]*)</div>', response.text, re.MULTILINE).group(1).split("，")
        perm_user = [p.strip() for p in perm_user]
        return perm_user

    def get_unread_notification(self):
        # 获取未读的系统提醒
        r = self.request("POST", "module/notification/list.do", data={
            "currPage": 1,
            "title": '',
            "hasView": 0
        }, headers={"Content-Type": "application/x-www-form-urlencoded"})
        all_mid = re.findall(r'msgcenter.viewMsg\((\d+),this\);', r.text)
        for mid in all_mid:
            self.view_notification(int(mid.strip()))
        return all_mid

    def view_notification(self, nid):
        r = self.request("POST", "module/notification/viewNoti.do", data={"notiId": nid}, headers={
            "Content-Type": "application/x-www-form-urlencoded"})
        subject = re.search(
            r'<td class="remark_6 text-right">主题：</td>\W+<td>(.*?)</td>', r.text).group(1)
        content = re.search(
            r'<div class="contents">\W+<div[^>]*>(.*?)</div>', r.text).group(1)
        noti_url = re.search('url=(.*?)', r.text).group(1)
        print(subject, nid, content, noti_url)

    def get_companyorg(self):
        return self.request("GET", "mgt/rest/common/get_companyorg", params={"version": 1}).json()

    def get_users_bydeptid(self, deptid):
        return self.request("GET", "mgt/rest/common/get_users_bydeptid", params={"departmentId": deptid}).json()

    def connect_chat(self):
        url = "https://eim8.jingoal.com/"
        graceful = False
        username = self.username
        passwd = self.sha_passwd
        rnd_id = math.ceil(
            100000.5 + (((900000.49999) - (100000.5)) * random.random()))
        reqstr = "<body><iq xmlns='' type='get' id='{rnd_id}'><auth><username>{username}</username><password>{password}</password>"
        if graceful:
            reqstr += "<type>graceful</type></auth></iq></body>"
        else:
            reqstr += "</auth></iq></body>"
        payload = reqstr.format(
            rnd_id=rnd_id, username=username, password=passwd)
        r = self.sesstion.post(url, data=payload, headers={
            "Content-Type": 'text/xml; charset=utf-8', "Connection": 'keep-alive'})
        sid = re.search('sid="(.*?)"', r.text).group(1)
        return sid

    def send_msg(self, sid=None, msg="中国china"):
        sid = self.connect_chat()
        url = "https://eim8.jingoal.com/?{}".format(sid)
        data = '<body><message from="{user_id}@{company_id}/HTTP" type="chat" to="{user_id}@{company_id}/EIM" xmlns="" id="WT.1473164071454"><body>{message}</body></message></body>'
        data = data.format(message=msg, user_id=self.user_id, company_id=self.company_id).encode("utf-8")
        r = self.sesstion.post(url,
                               data=data,
                               headers={"Content-Type": 'text/xml; charset=utf-8', "Connection": 'keep-alive'})
        print("===", r.text)

    def add_attend(self, shift=1):
        # 签到
        response = self.request("POST", "attendance/attend/addAttend.do", data={
            "subId": 0,
            "shift": shift,  # 1 签到  2 签退
            "sure": True,
            "clockTypeStr": "windows",
        }, headers={"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"})
        if response.text.strip() == '0':
            print("add attend success.")
        return self.request("POST", "attendance/attend/takeAttendance.do", data={"isClient": False}).text

    def sha1(self, string):
        """
        编码密码的，不知道为什么这么复杂，从js转换过来的
        https://web.jingoal.com/user/js/dest/common.js
        """
        def rotateLeft(lValue, iShiftBits):
            a = ctypes.c_int32(lValue)
            a.value <<= iShiftBits
            return a.value | (lValue & 0xffffffff) >> 32 - iShiftBits

        def cvtHex(value):
            v = string = ""
            for i in range(-7, 1):
                v = ctypes.c_int32(value >> 4 * -i & 15).value
                string += format(v, 'x')
            return string

        def uTF8Encode(string):
            output = ""
            string = string.replace('\x0d\x0a', "\n")
            for s in string:
                c = ord(s)
                if c < 128:
                    output += chr(c)
                elif c > 127 and c < 2048:
                    output += chr(c >> 6 | 192)
                    output += chr(63 & c | 128)
                else:
                    output += chr(c >> 12 | 224)
                    output += chr(c >> 6 & 63 | 128)
                    output += chr(63 & c | 128)
            return output

        def charCode(char):
            return int(ord(char))
        H0 = 1732584193
        H1 = 4023233417
        H2 = 2562383102
        H3 = 271733878
        H4 = 3285377520
        string = uTF8Encode(string)
        stringLength = len(string)
        count = 0
        wordArray = []
        i = ""
        W = []
        while count < stringLength - 3:
            j = int(ord(string[count])) << 24 | int(ord(string[
                count + 1])) << 16 | int(ord(string[count + 2])) << 8 | int(ord(string[count + 3]))
            wordArray.append(j)
            count += 4

        if stringLength % 4 == 0:
            i = 2147483648
        elif stringLength % 4 == 1:
            i = charCode(string[stringLength - 1]) << 24 | 8388608
        elif stringLength % 4 == 2:
            i = charCode(
                string[stringLength - 2]) << 24 | charCode(string[stringLength - 1]) << 16 | 32768
        elif stringLength % 4 == 3:
            i = charCode(string[stringLength - 3]) << 24 | charCode(
                string[stringLength - 2]) << 16 | charCode(string[stringLength - 1]) << 8 | 128

        wordArray.append(i)
        while len(wordArray) % 16 != 14:
            wordArray.append(0)

        blockstart = 0
        wordArray.append(stringLength >> 29)
        wordArray.append(stringLength << 3 & 4294967295)
        while blockstart < len(wordArray):
            for i in range(16):
                W.append(wordArray[blockstart + i])
            for i in range(16, 80):
                W.append(rotateLeft(W[i - 3] ^ W[i - 8] ^ W[i - 14] ^ W[i - 16], 1))
            A = H0
            B = H1
            C = H2
            D = H3
            E = H4
            for i in range(20):
                tempValue = rotateLeft(
                    A, 5) + (B & C | ~B & D) + E + W[i] + 1518500249 & 4294967295
                E = D
                D = C
                C = rotateLeft(B, 30)
                B = A
                A = tempValue
            for i in range(20, 40):
                tempValue = rotateLeft(A, 5) + (B ^ C ^ D) + \
                    E + W[i] + 1859775393 & 4294967295
                E = D
                D = C
                C = rotateLeft(B, 30)
                B = A
                A = tempValue
            for i in range(40, 60):
                tempValue = rotateLeft(
                    A, 5) + (B & C | B & D | C & D) + E + W[i] + 2400959708 & 4294967295
                E = D
                D = C
                C = rotateLeft(B, 30)
                B = A
                A = tempValue
            for i in range(60, 80):
                tempValue = rotateLeft(A, 5) + (B ^ C ^ D) + \
                    E + W[i] + 3395469782 & 4294967295
                E = D
                D = C
                C = rotateLeft(B, 30)
                B = A
                A = tempValue

            H0 = ctypes.c_int32(H0 + A & 4294967295).value
            H1 = ctypes.c_int32(H1 + B & 4294967295).value
            H2 = ctypes.c_int32(H2 + C & 4294967295).value
            H3 = ctypes.c_int32(H3 + D & 4294967295).value
            H4 = ctypes.c_int32(H4 + E & 4294967295).value
            blockstart += 16
        tempValue = cvtHex(H0) + cvtHex(H1) + cvtHex(H2) + \
            cvtHex(H3) + cvtHex(H4)
        return tempValue.lower()


def is_holiday(day):
    response = requests.get("http://www.easybots.cn/api/holiday.php", params={"d": day})
    result = response.json()[day if isinstance(day, str) else day[0]]
    return True if result == "1" else False


def is_holiday_today():
    """
    判断今天是否为节假日
    return: bool
    """
    today = datetime.date.today().strftime('%Y%m%d')
    return is_holiday(today)


def get_current_week_num():
    """获得当前周数"""
    return datetime.datetime.now().date().isocalendar()[1]


def current_year():
    return datetime.datetime.now().year


class TestFunction(unittest.TestCase):
    jg = JinGoal("15639270312", "mypassword")

    def test_login(self):
        pass

    def test_reset_log(self):
        result = self.jg.reset_today_log()
        self.assertIsInstance(result, str)
        self.assertIsInstance(int(result), int)

    def test_write_log(self):
        print("write log")
        result = self.jg.write_worklog(content='test contents - 1')
        self.assertIsInstance(result, dict)
        self.assertEqual(result["ok"], True)
        self.assertIsInstance(result["id"], int)
        print("reset log")
        log_id = self.jg.reset_today_log()
        result = self.jg.add_igoal(log_id)
        print("add igoal result ", result)
        igoal_id = re.findall(r"RA1B1_(\d+)", result["data"][0]["item"]["itemRefer"])[0]
        result = self.jg.delete_igoal(igoal_id)
        print("delete igoal for ", igoal_id, result)
        print("list log")
        result = self.jg.list_my_worklog()
        print(result)
        self.assertIsInstance(result, list)

    def test_reset_today_igoal(self):
        result = self.jg.reset_today_igoal()
        print("reset igoal", result)

    def test_igoal(self):
        igoal_list = self.jg.get_igoal_list()
        self.assertIsInstance(igoal_list, list)
        self.jg.print_igoal_list()

    def test_sha1(self):
        sha1 = self.jg.sha1('123456')
        self.assertEqual(sha1, '7c4a8d09ca3762af61e59520943dc26494f8941b')

    def test_user_log(self):
        # self.jg.read_all_user_log()
        pass

    def test_notification(self):
        result = self.jg.get_unread_notification()
        print(result)

    def test_attend(self):
        self.jg.add_attend()

if __name__ == "__main__":
    doctest.testmod()
    unittest.main()

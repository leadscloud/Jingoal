from jingoal import JinGoal
import re
import os
import base64
import json


def dowload_contacts():
    def get_user_byid(userid, users):
        if not users:
            return users
        for user in users:
            if user["user_id"] == userid:
                return user
    def loop_process(nodes, level=0, users=None):
        for node in nodes:
            user_data = None
            # if users and "data" in users:
            #     user_data = get_user_byid(node["id"], users["data"])

            # if users:
            #     user_data = get_user_byid(node["id"], users["data"])

            # users = jg.get_users_bydeptid(node["id"])
            # print(users, users["meta"]["code"])
            # if users["meta"]["code"] == 200:
            #     user_data = get_user_byid(node["id"], users["data"])



            if "children" in node:
                # users = jg.get_users_bydeptid(node["id"])
                if users:
                    user_data = get_user_byid(node["id"], users["data"])
                if user_data:
                    print("{level}{name} - {id} - {gender} - {duty_name} - {email} - {mobile} - {corp_phone} - {login_name}".format(**node, level="----"*level, **user_data))
                else:
                    print("{level}{name} - {id}".format(**node, level="----"*level))

                if node["children"]:
                    # print(users)
                    loop_process(node["children"], level=level+1, users=jg.get_users_bydeptid(node["id"]))
            else:
                # users = jg.get_users_bydeptid(node["id"])
                if "data" in users:
                    user_data = get_user_byid(node["id"], users["data"])
                if user_data:
                    print("{level}{name} - {id} - {gender} - {duty_name} - {email} - {mobile} - {corp_phone} - {login_name}".format(**node, level="----"*level, **user_data))
                else:
                    print("{level}{name} - {id}".format(**node, level="----"*level))

    response = jg.get_companyorg()
    if response["meta"]["code"] == 200:
        org_nodes = response["data"]["org_nodes"]
        loop_process(org_nodes)


def download_frames(frames, parent=""):
    # frames = jg.igoal_detail(frame_id)
    for f in frames["frames"]:
        if "hasChilds" in f and f["hasChilds"] is True:
            dir_path = "dowload/" + parent + "/" + f["igoalFrame"]["title"] + "/"
            if not os.path.exists(dir_path):
                os.makedirs(dir_path)
            childs = jg.igoal_childs(f["id"])
            # print("childs", childs)
            # download_frames(childs)
            for item in childs["items"]:
                # print("-----------", item)
                item_detail = jg.igoal_detail_item(item["id"])
                source_url = re.search('src="(.*?)"', item_detail)
                src_url = re.search('src=".*?url=(.*?)&[^"]*token=(.*?)"', item_detail)
                # print(src_url.group(1), src_url.group(2))
                if item["referType"] == "plan":
                    detail_url = "https://web.jingoal.com/webcd/plan/planDetail4Igoal.do"
                elif item["referType"] == "task":
                    detail_url = "https://web.jingoal.com/webcd/task/taskDetail4Igoal.do"
                elif item["referType"] == "log":
                    detail_url = "https://web.jingoal.com/module/worklog/logDetail4IgoalIframe.do"
                elif item["referType"] == "doc":
                    detail_url = source_url.group(1)
                elif item["referType"] == "memo":
                    detail_url = "https://web.jingoal.com/webxm/memo/memoDetail4IgoalRefer.do"

                # r = jg.get("https://web.jingoal.com/webcd/plan/planDetail4Igoal.do?token={0}&url={1}".format(src_url.group(2), src_url.group(1)))
                if item["referType"] != "doc":
                    r = jg.get(detail_url, params={"token": src_url.group(2), "url": src_url.group(1)})
                    print(item["item"]["title"])
                    file_title = item["ownerName"] + "-" + item["typeName"] + "-" + item["item"]["title"] + ".html"
                    file_title = re.sub(r'[^-a-zA-Z0-9\u4e00-\u9fff_.()/ ]+', '', file_title)
                    file_title = file_title.replace("/", "-")
                    with open(dir_path + file_title, "w", encoding="utf-8") as html_file:
                        html_file.write(r.text.strip())
                    # 如果日志或者其它类型中有附件，下载之
                    files = re.findall('<a[^>]*title="(.+?)"[^>]*href=".+?(url=token&rid=.*?&mgtToken=.*?)&', r.text)
                    for attach_filename, attach_url in files:
                        if not os.path.exists(dir_path + attach_filename.strip()):
                            print("下载附件...", attach_filename, attach_url)
                            attach_steam = jg.request("GET", "https://web.jingoal.com/webcd/resources/?" + attach_url, stream=True)
                            with open(dir_path + attach_filename.strip(), 'wb') as fs:
                                for chunk in attach_steam:
                                    if chunk:
                                        fs.write(chunk)

                elif item["referType"] == "doc":
                    # print("download files .....")
                    # print(detail_url)
                    itemRefer = json.loads(item["item"]["itemRefer"])
                    # print(itemRefer)
                    # print(">>", re.search('token=(.+)', detail_url).group(1))
                    payload_text = """callCount=1
page={page}
httpSessionId=
scriptSessionId=${{scriptSessionId}}229
c0-scriptName=dwrService
c0-methodName=getDWRIncludeURL
c0-id=0
c0-e1=number:2755052
c0-param0=Object_Object:{{id:reference:c0-e1}}
c0-e2=string:%2Fmydoc%2Fsysdoc%2FiGoal%2FshowAttachDetail.jsp%3Ffile%3Dfalse%26url%3D{url}%26token%3D{token}%26online_ids_stamp%3D0
c0-param1=Array:[reference:c0-e2]
batchId=0""".format(page=detail_url, token=re.search('token=(.+)', detail_url).group(1), url=itemRefer["url"])
                    rs = jg.request("POST", 'https://web.jingoal.com/mgt2/dwr/call/plaincall/dwrService.getDWRIncludeURL.dwr?t=&uid=2755052&cid=undefined', data=payload_text, headers={"Content-Type": "text/plain", "X-Requested-With": 'XMLHttpRequest'})

                    json_data = re.search('s1\[0\]="(.+)"', rs.text)
                    # html_parser = html.parser.HTMLParser()
                    # unescaped = html_parser.unescape(rs.text)
                    parsed_text = json_data.group(1).encode('utf-8').decode('unicode_escape')
                    # print(parsed_text)
                    try:
                        download_url = re.search('href="(/mgt2/mydoc/sysdoc/iGoal/download.jsp.*?)"', parsed_text).group(1)
                        file_name = re.search('<a class="black01"[^>]*title="(.*?)"', parsed_text, re.MULTILINE).group(1)
                        # print(parsed_text)
                        if not os.path.exists(dir_path + file_name):
                            # 如果文件没有下载过，再下载
                            print("Dowloading file: {}".format(file_name))
                            steam = jg.request("GET", "https://web.jingoal.com" + download_url, stream=True)
                            with open(dir_path + file_name, 'wb') as fs:
                                for chunk in steam:
                                    if chunk: # filter out keep-alive new chunks
                                        fs.write(chunk)
                        else:
                            print("File {0} is exists, skip dowload.".format(file_name))
                        # 根据headers获取文件名，由于编码方式不一致，放弃
                        # d = steam.headers['Content-Disposition']
                        # print(d, steam.headers)
                        # fname = re.findall("filename=\"(.+)\"", d)[0]
                        # true_fname = re.search('=\?UTF-8\?B?(.*?)\?=', fname).group(1)
                        # true_fname = base64.b64decode(true_fname).decode("utf-8")
                        # print(">>>", true_fname)
                        # print(steam.headers, steam.request.url)


                    except AttributeError as e:
                        print("AttributeError", e, item)
            if "frames" in childs:
                download_frames(childs, parent=parent+"/"+f["igoalFrame"]["title"]+"/")


def get_all_igoal():
    r = jg.request("POST", "https://web.jingoal.com/igoal/igoal/list.do", params={
        "ownerId": 0,
        "status": -1,
        "tagId": -1,
        "key": "",
        "deleted": False,
        "enc": 0,
        "currPage":1
    }, headers={
        "Content-Type": "application/json",
        "X-Requested-With": "XMLHttpRequest"
    })
    igoal_ids = re.findall("viewIgoal\('(\d+)',0\)", r.text)
    print(igoal_ids)
    for igoal_id in igoal_ids:
        frames = jg.igoal_detail(igoal_id)
        title = frames["igoal"]["tagName"] + "-" + frames["igoal"]["igoal"]["name"]
        download_frames(frames, parent=title)

if __name__ == "__main__":
    jg = JinGoal("15639270312", "mypassword")
    get_all_igoal()
    # print(dowload_contacts())

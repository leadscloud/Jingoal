from jingoal import JinGoal
from GoogleAanlytics import ga_main
import datetime
import time

PROXY_HOST = '192.168.1.2'
PROXY_PORT = 1690

CURRENT_WEEK_LOG = """
上周工作总结：
===========

1. 每天网站发链接，优化网站
2. 检查网站内容，安排每天内容更新
3. 日常优化工作

"""

def get_site_traffic():
    # 首先自动获取流量数据并制作Excel文件
    return ga_main(PROXY_HOST, PROXY_PORT)

def write_week_log(log_content=""):
    def format_md_content(text):
        content = "<pre style=\"width: 850px;overflow:auto;\"><span style=\"font-family:simhei,Heiti SC Thin,Monaco,monospace;\">"
        content += "<br>".join(text.strip().splitlines())
        content += "</span></pre>"
        return content
    md_content, excel_filename = get_site_traffic()

    # 把每周工作总结附加上去
    if log_content:
        md_content = log_content + md_content
    else:
        md_content = CURRENT_WEEK_LOG + md_content

    #添加一个附注事项，可忽略
    md_content += "\n\n注: 如果显示较混乱请到日志详细页面查看，或者下载Excel查看数据"

    html_content = format_md_content(md_content)
    # 填写每周工作总结
    jg = JinGoal("15639270312", "mypassword")
    # jg._login()
    print("")
    print("*" * 40 + "以下为你的工作日志内容" + "*" * 40)
    print(md_content)
    print("*" * 40 + "*工作日志内容显示结束*" + "*" * 40)
    print("")
    # 先上传文件，没法同时写日志并上传
    log_id = jg.reset_today_log() # 先删除已经上传的所有文件，在重新写日志时这个需要操作
    jg.reset_today_igoal(log_id) # 重置今天已经添加到主线上的日志
    # print(log_id)
    result = jg.write_worklog(content=html_content, files=[excel_filename])
    if result:
        print("上传文件成功", result)
        log_id = result["id"]
    # 再写日志
    result = jg.write_worklog(content=html_content)
    if result:
        print("写日志成功", result)
    result = jg.add_igoal(log_id)
    print("添加日志到主线成功", result)
    for item in jg.list_igoal_item():
        igoal_id = item["id"]
        owener_name = item["ownerName"]
        read_status = item["read"]
        item_refer = item["item"]["itemRefer"]
        title = item["item"]["title"]
        create_time = datetime.datetime.fromtimestamp(item["item"]["createTime"]/1e3)
        # logid = re.search('(\d{4,})', item_refer).group(1)
        if create_time.date() == datetime.datetime.now().date():
            print("成功获取到今天主线中的日志，所有任务已结束^_^")
            print(igoal_id, owener_name, read_status, title)
            break
    else:
        print("没有获取到今天的主线中的日志，请重新执行本脚本")

if __name__ == "__main__":
    write_week_log()

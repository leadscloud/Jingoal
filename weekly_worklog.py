

from jingoal import JinGoal
from GoogleAanlytics import ga_main
import datetime
import time

PROXY_HOST = '192.168.1.2'
PROXY_PORT = 1690

"""
工作小结
第34周
上周总结：
1. 菲律宾站内搜索网站完成。
2. 沙特阿语优化网站，关键词更新，外链发布。
3. 日常外链发布 ，排名优化工作。

本周计划：
1. 针对菲律宾，沙特开发新的引流平台。
2. 肯尼亚网站流量有下降趋势，安排外链工作。
3. 日常外链发布 ，排名优化工作。
"""

CURRENT_WEEK_LOG = """
上周总结：
=========

1. 每天网站发链接，优化网站
2. 检查网站内容，安排每天内容更新
3. 日常优化工作

"""

def reset_today_log():
    jg = JinGoal("15639270312", "mypassword")
    # jg._login()
    log_id = jg.reset_today_log() # 先删除已经上传的所有文件，在重新写日志时这个需要操作
    jg.reset_today_igoal(log_id) # 重置今天已经添加到主线上的日志


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

def write_week_log2017(log_content=""):
    current_week_num = datetime.datetime.now().date().isocalendar()[1]
    title = "2017年第{week}周工作总结".format(week=current_week_num-1)

    html_content = ""
    excel_filename = ""

    # 填写每周工作总结
    print("login...")
    jg = JinGoal("15639270312", "mypassword")
    print("login success.")
    # jg._login()
    print("")
    print("*" * 40 + "以下为你的工作日志内容" + "*" * 40)
    print(html_content)
    print("*" * 40 + "*工作日志内容显示结束*" + "*" * 40)
    print("")
    # 先上传文件，没法同时写日志并上传
    log_id = jg.reset_today_log()  # 先删除已经上传的所有文件，在重新写日志时这个需要操作
    jg.reset_today_igoal(log_id)  # 重置今天已经添加到主线上的日志
    print(log_id)
    result = jg.write_worklog(content=html_content, files=[excel_filename])
    if result:
        print("上传文件成功", result)
        # log_id = result["id"]
    # 再写日志
    result = jg.write_worklog(content=html_content)
    if result:
        print("写日志成功", result, log_id)
    result = jg.add_igoal(log_id, igoalId=3054350, frameId=29175275, title=title)
    print("添加日志到主线成功", result)
    for item in jg.list_igoal_item(frame_id=29175275):
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

import argparse

def get_command_line(only_print_help=False):
    version = '1.0.0'
    parser = argparse.ArgumentParser(prog='JinGoalHelper',
                                     description='A atuo write worklog script for jingoal.com')

    parser.add_argument('-u', '--username', type=str, action='store', default='',
                        help='Ther account username')
    parser.add_argument('-p', '--password', type=str, action='store', default='',
                        help='Ther account password')
    parser.add_argument('-t', '--type', type=str, default='daily',
                        help='Write worklog',
                        choices=('daily', 'weekly'))
    parser.add_argument('-l', '--list-log', action='store_false', default=True, required=False,
                        help='List and view all users worklog.')
    parser.add_argument('-c', '--content', type=str, default='',
                        help='Worklog content')
    parser.add_argument('-r', '--reset', action='store_true', default=False, required=False,
                        help='This will reset all my log content, include attachs')
    parser.add_argument('-V', '--v', '--version', action='store_true', default=False, dest='version',
                        help='Prints the version of GoogleScraper')
    if only_print_help:
        parser.print_help()
    else:
        args = parser.parse_args()

        return vars(args)

def main(parse_cmd_line=True):
    if parse_cmd_line:
        config = get_command_line()
        print(config)
        username = config.get('username')
        password = config.get('password')
        log_type = config.get('type')
        log_content = config.get('content', '')
        is_reset = config.get('reset', False)
        list_log = config.get('list-log', True)

        jg = JinGoal(username, password)
        if is_reset:
            jg.reset_today_log()
        if list_log:
            jg.read_all_user_log()
        if log_type == "daily":
            log_id = jg.reset_today_log()
            jg.reset_today_igoal(log_id)
            print("Now start write daily log...")
            result = jg.write_worklog(content=log_content)
            print("Write complete!", result)
        elif log_type == "weekly":
            print("Now start write weekly log...")
            write_week_log2017()
            print("Write complete!")
        else:
            print('Not support log type')
            return


def main2(jg, parse_cmd_line=True):

    cmd = input(
    """
1. Write my daily worklog

2. Write my weekly worklog

3. View all users worklog

4. qiandao

5. qiantui
    """
    )
    if cmd == '1':
        content = input("pls input your daily worklog content, default is null:")
        jg.write_worklog(content=content)
        print("wirte log complete.")
    elif cmd == '2':
        w_content = input("pls input your weekly worklog content, default is auto:")
        if w_content == '':
            write_week_log()
            print("wirte weekly log complete!")
    elif cmd == '3':
        jg.read_all_user_log()
    elif cmd == '4':
        # 签到
        jg.add_attend(1)
    elif cmd == '5':
        # 签退
        jg.add_attend(2)
    elif cmd == '6':
        write_week_log2017()
    else:
        main(jg)



if __name__ == "__main__":
    print('start run script....')
    start = time.time()
    #write_week_log()
    # reset_today_log()
    # print("First, let me login the system...")
    # jg = JinGoal("15639270312", "mypassword")
    write_week_log2017()
    # main()
    print("Script run spent time is: {0:.1f} Second".format(time.time() - start))

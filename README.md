

为了应对公司每天日志的要求，写的自动发布日志的脚本。可以保证你每天的日志都是按时去写的，内容丰满，领导满意。

程序运行在python3上面


**关键词：**

- http://jingoal.com/index.html
- jingoal
- 今目标
- 协同工具
- 今目标自动化
- 自动填写今目标的工作日记



**需要安装的包：**

- requests

其它可能需要的包：

看出错提示，自己安装。


各个文件说明：
----------------

`weekly_worklog.py`
================

我每周的工作内容： 自动生成GA报表，格式化成table，然后上传并放到任务主线上面，同时可以@其它人。

`GoogleAanlytics.py` 
========

这个文件是因为工作日志内容需要几个团队成员相应网站的流量信息，所以写了这个脚本，可以自动获取团队成员相应网站的流量情况。

`tabulate.py` 
==========

作用同上，仅仅为了格式化生成日志中的table，让它显示的好看，在网页中能对齐。

`download_log.py`
================

下载今目标中所有的你可以看到的信息：包括公司架构、公司所有人的通讯录，你的所有日志或者你团队成员的所有日志

`win32com_excel.py`
=================

处理excel的


添加定时任务
============


可以制作两个文件  `weekly_run.vbs` `daily_run.vbs`


**weekly_run.vbs 每周运行一次**

```
Set ws = WScript.CreateObject("WScript.Shell")
ws.run "cmd.exe /C ""C:\Python35-32\python.exe weekly_worklog.py -l -t weekly -u username -p password"" ", 0, True
```

**`daily_run.vbs` 每日运行一次**

```
Set ws = WScript.CreateObject("WScript.Shell")
ws.run "cmd.exe /C ""C:\Python35-32\python.exe weekly_worklog.py -l -t daily -u username -p password"" ", 0, True
```

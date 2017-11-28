"""Hello Analytics Reporting API V4."""

import argparse

from apiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials

import httplib2
from oauth2client import client
from oauth2client import file
from oauth2client import tools
import socks
# from tabulate import tabulate
import tabulate
import datetime

tabulate.WIDE_CHARS_MODE = True


SCOPES = ['https://www.googleapis.com/auth/analytics.readonly']
DISCOVERY_URI = ('https://analyticsreporting.googleapis.com/$discovery/rest')
KEY_FILE_LOCATION = 'ga-9a7fb73f881d.p12'
SERVICE_ACCOUNT_EMAIL = 'id-report-631@eloquent-hangar-142201.iam.gserviceaccount.com'
VIEW_ID = '42140195'
PROXY_HOST = '127.0.0.1'
PROXY_PORT = 1080


def initialize_analyticsreporting(proxy_host=PROXY_HOST, proxy_port=PROXY_PORT):
    """Initializes an analyticsreporting service object.

    Returns:
      analytics an authorized analyticsreporting service object.
    """

    credentials = ServiceAccountCredentials.from_p12_keyfile(
        SERVICE_ACCOUNT_EMAIL, KEY_FILE_LOCATION, scopes=SCOPES)

    socks.setdefaultproxy(socks.PROXY_TYPE_SOCKS5, proxy_host, proxy_port)
    socks.wrapmodule(httplib2)

    # p = httplib2.proxy_info_from_url("http://192.168.1.2:1690")
    # proxy = httplib2.ProxyInfo(socks.PROXY_TYPE_HTTP, "192.168.1.2", 1690)
    http = credentials.authorize(httplib2.Http(timeout=35))

    # http = credentials.authorize(httplib2.Http())

    # Build the service object.
    analytics = build('analytics', 'v4', http=http,
                      discoveryServiceUrl=DISCOVERY_URI)

    return analytics


def get_report(analytics, view_id=VIEW_ID):
    # Use the Analytics Service Object to query the Analytics Reporting API V4.
    # https://developers.google.com/analytics/devguides/reporting/core/v4/rest/v4/reports/batchGet
    country_filter = {
        "dimensions": [{"name": "ga:country"}],
        "dimensionFilterClauses": [
            {
                "filters": [
                    {
                        "dimensionName": "ga:country",
                        "operator": "EXACT",
                        "expressions": ["China"]
                    }
                ]
            }
        ]
    }
    return analytics.reports().batchGet(
        body={
            'reportRequests': [
                {
                    'viewId': view_id,
                    'dateRanges': [{'startDate': '30daysAgo', 'endDate': 'today'}],
                    'metrics': [
                        {'expression': 'ga:sessions', "alias": "会话数"},
                        {'expression': 'ga:users', "alias": "用户数"},
                        {'expression': 'ga:pageviews', "alias": "网页浏览量"},
                        {'expression': 'ga:pageviews/ga:sessions',
                            "alias": "每次会话浏览页数", "formattingType": "FLOAT"},
                        {'expression': 'ga:avgSessionDuration',
                            "alias": "平均会话时长", "formattingType": "TIME"},
                        {'expression': 'ga:bounceRate',
                            "alias": "跳出率", "formattingType": "FLOAT"},
                        {'expression': 'ga:percentNewSessions',
                            "alias": "新会话百分比", "formattingType": "FLOAT"},
                    ],
                    # "dimensions": [{"name": "ga:country"}],
                    "dimensionFilterClauses": [
                        {
                            "filters": [
                                {
                                    "dimensionName": "ga:country",
                                    "not": True,
                                    "operator": "EXACT",
                                    "expressions": ["China"],
                                    "caseSensitive": True,
                                }
                            ]
                        }
                    ]
                }]
        }
    ).execute()

def parse_response(response):
    header_result = result = []
    for report in response.get('reports', []):
        columnHeader = report.get('columnHeader', {})
        dimensionHeaders = columnHeader.get('dimensions', [])
        metricHeaders = columnHeader.get(
            'metricHeader', {}).get('metricHeaderEntries', [])
        rows = report.get('data', {}).get('rows', [])
        header_result.append()
        result.append(rows)

def print_response(response):
    """Parses and prints the Analytics Reporting API V4 response"""
    result = []
    for report in response.get('reports', []):
        columnHeader = report.get('columnHeader', {})
        dimensionHeaders = columnHeader.get('dimensions', [])
        metricHeaders = columnHeader.get(
            'metricHeader', {}).get('metricHeaderEntries', [])
        rows = report.get('data', {}).get('rows', [])
        print(rows)

        for row in rows:
            dimensions = row.get('dimensions', [])
            dateRangeValues = row.get('metrics', [])

            for header, dimension in zip(dimensionHeaders, dimensions):
                print(header + ': ' + dimension)

            for i, values in enumerate(dateRangeValues):
                # print(metricHeaders, values.get('values'))
                print('Date range (' + str(i) + ')')
                print("\t".join([item["name"] for item in metricHeaders]))
                print("\t".join(values.get('values')))
                # for metricHeader, value in zip(metricHeaders, values.get('values')):
                #     print(metricHeader.get('name') + ': ' + value)
    # return values
from win32com_excel import RemoteExcel
from tabulate import _format, _type
import os

def ga_main(proxy_host=PROXY_HOST, proxy_port=PROXY_PORT):
    ga_list = {
        "名字1": {
            "boliviatrituradora.com": 122634071,
            "ghanagoldcrusher.com": 123755226,
            "sbmcrush.com": 121998307,
            "grindingmill.org": 75864545,
            "whitecrusher.com": 121980632
        },
        "名字2": {
            "sbmmolienda.es": 124922819,
            "arabiacrusher.org": 124910045,
            "ht-rent.net": 119759542,
            "orecrusherplant.org": 119777155,
            "canadacrusher.com": 121975207,
            "safiaalsouhail.com": 124912845,
        },
        "名字3": {
            "etcrusher.com": 125573509,
            "penghancurkon.com": 123796155,
            "qualitygrinder.com": 116835153,
            "grindermarket.com": 119780443,
        },
        "名字4": {
            "gcphome.org": 124821844,
            "crusher-manufacturers.com": 95115399,
            "gauleyriverland.com": 126651806,
            "homeandhoteldepot.com": 127850279,
            "hotelroyalparkkashmir.com": 128404423,
            "swingwestband.com": 121989104
        },
        "名字5": {
            "rahangpenghancur.com": 123807634,
            "myanmarcrusher.com": 124121474,
        },
        "名字6": {
            # "monteverdehotel.net": 126755256,
            "crusherdz.com": 125726328,
            "akpcollegekhurja.org": 122676205,
            "animalhospitallongbeachca.com": 126669954,
            "bellgardenshotel.com": 126667466,
            "bernardweismanfoundation.org": 122614425,
            "brickcapitalcdc.com": 121978120,
            "canbalafi.com": 121999421,
            "canterburyfarmshoa.org": 125029844
        },
        "名字7": {
            "rpgroupbd.com": 124243410,
            "algeriacrusher.org": 125889411,
            "oleinrecovery.net": 119843464,
            "kassiopi.org": 119836171,
            "eacheer.com": 119834166,
            "clemsondsp.com": 119870413,
            "bobopastachef.com": 119836165,
        },
        "名字8": {
            "trituradoradecantera.com": 123001133,
            "maynghienmay.com": 124020031,
            "crushermaintenance.com": 119772033,
        }
    }
    analytics = initialize_analyticsreporting(proxy_host, proxy_port)
    print_header = True
    table = table_headers = []
    for name, data in ga_list.items():
        # print("-"*20 + name + "-"*20)
        for domain, view_id in data.items():
            # print(domain, view_id)
            response = get_report(analytics, str(view_id))
            for report in response.get('reports', []):
                columnHeader = report.get('columnHeader', {})
                # dimensionHeaders = columnHeader.get('dimensions', [])
                metricHeaders = columnHeader.get('metricHeader', {}).get('metricHeaderEntries', [])
                rows = report.get('data', {}).get('rows', [])
                for row in rows:
                    dateRangeValues = row.get('metrics', [])
                    for i, values in enumerate(dateRangeValues):
                        if print_header:
                            table.append(["姓名", "域名", *[item["name"] for item in metricHeaders]])
                            # table_headers = ["姓名", "域名", *[item["name"] for item in metricHeaders]]
                            print("姓名\t域名\t{data}".format(data="\t".join([item["name"] for item in metricHeaders])))
                            print_header = False
                        row =  [_format(v, _type(v), floatfmt=".2f") for v in values.get('values') ]
                        table.append([name, domain, *row])
                        print("{name}\t{domain}\t{data}".format(name=name, domain=domain, data="\t".join(values.get('values'))))

    start_date = (datetime.datetime.now() + datetime.timedelta(-30)).strftime("%Y-%m-%d")
    end_date = (datetime.datetime.now() + datetime.timedelta(-1)).strftime("%Y-%m-%d")
    md_content = "最近30天排除中国流量统计({0} - {1})：\n".format(start_date, end_date)
    md_content += "================================================== \n\n"
    md_content += tabulate.tabulate(table, headers="firstrow", tablefmt="pipe", numalign="right", floatfmt=".2f")
    # print(md_content)
    print("strart create excel file...")
    exce_file = "重点市场优化网站列表.xlsx"
    excel = RemoteExcel(os.path.abspath(exce_file))
    first_sheet = excel.xlBook.Worksheets(1)
    if first_sheet.Name == "流量统计概览":
        ws = first_sheet
    else:
        ws = excel.xlBook.Sheets.Add(Before=excel.xlBook.Worksheets(1))
        ws.Name = "流量统计概览"
    ws.Columns.AutoFit()

    rowCnt = 1
    for row in table:
        ws.Range(ws.Cells(rowCnt,1), ws.Cells(rowCnt, len(row))).Value = row
        # ws.Columns.AutoFit()
        rowCnt += 1
    # 应用样式
    ws.Range(ws.Cells(1,1), ws.Cells(len(table), len(table[0]))).AutoFilter(1)
    ws.Columns.AutoFit()
    ws.Rows(1).Font.Bold = True
    ws.Range(ws.Cells(1,1), ws.Cells(len(table), len(table[0]))).RowHeight = 15
    ws.Rows(1).RowHeight = 18
    ws.Range(ws.Cells(1,1), ws.Cells(len(table), len(table[0]))).Borders.LineStyle = 1
    ws.Range(ws.Cells(1,1), ws.Cells(len(table), len(table[0]))).Borders.Weight = 2
    excel.xlApp.ActiveWindow.SplitRow = 1
    excel.xlApp.ActiveWindow.FreezePanes  = True
    excel.save()
    excel.close()
    print("end create excel file.")
    return md_content, exce_file
    # response = get_report(analytics)
    # print(response)
    # print_response(response)

if __name__ == '__main__':
    ga_main()

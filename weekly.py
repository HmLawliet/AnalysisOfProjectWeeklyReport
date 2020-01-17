
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr, parsedate_to_datetime
import datetime
import pandas as pd
from bs4 import BeautifulSoup
import poplib
import copy
import re
from itertools import zip_longest
import plotly.figure_factory as ff
import plotly.graph_objects as go
from reportlab.platypus import SimpleDocTemplate,Paragraph,Image
from reportlab.lib.pagesizes import A2,A4
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.styles import getSampleStyleSheet
import fitz
from robot import send_work_notice_file
from collections import OrderedDict
from pyecharts.render import make_snapshot
from snapshot_phantomjs import snapshot
from pyecharts.components import Table
from pyecharts.options import ComponentTitleOpts
import imgkit
from os import listdir
from PIL import Image


def guess_charset(msg):
    '''
    邮件的编码解析
    param msg: 解析的邮件内容
    return  邮件的编码格式
    '''
    charset = msg.get_charset()
    if charset is None:
        content_type = msg.get('Content-Type', '').lower()
        pos = content_type.find('charset=')
        if pos >= 0:
            charset = content_type[pos + 8:].strip()
    return charset


def decode_str(s):
    '''
    邮件解码字符串
    param s ： 要解析的字符串
    return  解析后的值
    '''
    value, charset = decode_header(s)[0]
    if charset:
        value = value.decode(charset)
    return value


def select_by_mail_header(msg, receiver, date_start,projectManager):
    '''
    判断是否为要解析的邮件
    param msg:
    param receiver:
    param date_start:
    param projectManager:
    return:
    '''
    mail_date = parsedate_to_datetime(msg.get('Date', "")).date()
    mail_sender = parseaddr(msg.get('From', ""))[1]
    mail_receiver = parseaddr(msg.get('To', ""))[1]
    try: 
        flag = decode_str(msg.get('Subject', "")).startswith('项目周报')
    except Exception:
        flag = False
    return mail_receiver == receiver and mail_sender in projectManager and mail_date >= date_start and  flag , mail_date < date_start


def get_mail_content(msg):
    '''
    解析邮件内容
    param msg： 解码后的邮件
    return 返回html格式的内容
    '''
    htmlcontent = None
    parts = msg.get_payload()
    for n, part in enumerate(parts):
        content_type = part.get_content_type()
        if content_type == 'text/plain' or content_type == 'text/html':
            content = part.get_payload(decode=True)
            charset = guess_charset(part)
            if charset:
                content = content.decode(charset)
                if content_type == 'text/html':
                    htmlcontent = copy.deepcopy(content)
    return htmlcontent


def get_tasks(htmlcontent):
    '''
    从html中获得项目周报内容
    param htmlcontent ： html格式的邮件内容
    return 存储任务的字典格式数据
    '''
    works_dict={}
    if not isinstance(htmlcontent, str):
        return
    soup = BeautifulSoup(htmlcontent, 'html.parser')
    rows = soup.findAll('table')[0].tbody.findAll('tr')
    project = None
    flag = 8
    inner_flag = 0
    for index, row in enumerate(rows):
        if index == 1:
            project = row.findAll('td')[1].text
            if not project:
                return
            if not project in works_dict.keys():
                works_dict[project] = {}
            works_dict[project]['项目经理'] = row.findAll('td')[3].text
            works_dict[project]['项目里程碑'] = OrderedDict()
            works_dict[project]['项目较大变更说明'] = OrderedDict()
            works_dict[project]['项目问题及风险点'] = OrderedDict()
            works_dict[project]['本周主要工作进展'] = OrderedDict()
            works_dict[project]['下周工作计划'] = OrderedDict()
            works_dict[project]['项目文档地址'] = OrderedDict()
        if index == 3:
            tmp_row = row.findAll('td')
            works_dict[project]['当前阶段'] = (tmp_row[0].text,tmp_row[1].text,tmp_row[2].text,tmp_row[3].text,tmp_row[4].text,tmp_row[5].text)
        if index == 5:
            works_dict[project]['项目概要'] = row.findAll("td")[0].text
        if index >= flag:
            tmp_row = row.findAll("td")
            if inner_flag == 0:
                if tmp_row[0].text == '':
                    continue  
                if tmp_row[0].text == '项目较大变更说明':
                    flag = index + 2
                    inner_flag = 1
                    continue
                works_dict[project]['项目里程碑'][tmp_row[0].text] = (tmp_row[1].text,tmp_row[2].text,tmp_row[3].text,tmp_row[4].text,tmp_row[5].text)  
            if inner_flag == 1:
                if tmp_row[0].text == '':
                    continue
                if tmp_row[0].text == '项目问题及风险点':
                    flag = index + 2
                    inner_flag = 2
                    continue
                works_dict[project]['项目较大变更说明'][tmp_row[0].text] = (tmp_row[1].text,tmp_row[2].text,tmp_row[3].text)  
            if inner_flag == 2:
                if tmp_row[0].text == '':
                    continue
                if tmp_row[0].text == '本周主要工作进展':
                    flag = index + 2
                    inner_flag = 3
                    continue
                works_dict[project]['项目问题及风险点'][tmp_row[0].text] = (tmp_row[1].text,tmp_row[2].text,tmp_row[3].text,tmp_row[4].text,tmp_row[5].text)  
            if inner_flag == 3:
                if tmp_row[0].text == '':
                    continue 
                if tmp_row[0].text == '下周工作计划':
                    flag = index + 2
                    inner_flag = 4
                    continue
                works_dict[project]['本周主要工作进展'][tmp_row[0].text] = (tmp_row[1].text,tmp_row[2].text,tmp_row[3].text,tmp_row[4].text,tmp_row[5].text)         
            if inner_flag == 4:
                if tmp_row[0].text == '':
                    continue 
                if tmp_row[0].text == '项目文档地址':
                    flag = index + 2  
                    inner_flag = 5
                    continue 
                works_dict[project]['下周工作计划'][tmp_row[0].text] = (tmp_row[1].text,tmp_row[2].text,tmp_row[3].text,tmp_row[4].text)   
            if inner_flag == 5:
                if tmp_row[0].text == '':
                    continue
                if tmp_row[0].text =='项目组成员':
                    works_dict[project]['项目组成员'] = tmp_row[1].text  
                    break
                works_dict[project]['项目文档地址'][tmp_row[0].text] = tmp_row[1].text  
    return works_dict


def parse_project_milestones(stages):
    '''
    解析当前项目阶段
    param stages： 当前项目阶段   type orderdict
    return     返回换行后的字符串格式数据
    '''
    tmp_list = []
    for key,values in stages.items():
        if values[0] == '进行中':
            tmp_list.append(f'{key} <— ')
            continue 
        tmp_list.append(key)
    return '\n'.join(tmp_list)

def parse_project_risks(risks):
    '''
    解析当前项目风险和问题
    param risks: 当前项目风险  type orderdict
    return  返回换行后的字符串格式数据
    '''
    tmp_list = []
    for _ ,values in risks.items():
        tmp_list.append('\n'.join(values))
    return '\n'.join(tmp_list)
    
def parse_project_processes(processes):
    '''
    解析当前项目本周工作情况
    param processes : 本周的工作情况 type orderdict
    return 返回换行后的字符串格式数据
    '''
    tmp_processes = []
    for key, _ in processes.items():
        tmp_processes.append(key)
    works = '\n'.join(tmp_processes)
    return works
    
def parse_project_outlines(outlines):
    '''
    解析当前项目的项目概要
    param outlines: 项目概要   type orderdict
    return 换行后的项目概要  
    '''
    try:
        tmp_str_0,tmp_str_1 = outlines.strip().split('详述')
    except Exception:
        return outlines
    tmp_list = []
    if tmp_str_1:
        for index, item in enumerate(tmp_str_1):
            if index % 10 == 1:
                tmp_list.append('\n')
            tmp_list.append(item)
    tmp_str_1 = ''.join(tmp_list)
    return f'{tmp_str_0}\n详述\t{tmp_str_1}'

def pares_project_members(members):
    '''
    解析项目成员
    param members: 项目成员  type orderdict
    return 换行后的项目成员
    '''
    members = members.strip().replace('产品：', '\n产品').replace("开发", "\n开发").replace("测试", "\n测试").replace("设计", "\n设计")
    member_lists = members.strip().split('\n')
    result_list = []
    for members in member_lists:
        member_list = [item for item in members]
        if '、' in member_list:
            index_list = [index for index, item in enumerate(member_list) if item == '、']
            for index, l_index in enumerate(index_list):
                if index % 2 == 0:
                    continue
                member_list[l_index] = '\n\t'
        if '，' in member_list:
            index_list = [index for index, item in enumerate(member_list) if item == '，']
            for index, l_index in enumerate(index_list):
                if index % 2 == 0:
                    continue
                member_list[l_index] = '\n'
        if ',' in member_list:
            index_list = [index for index, item in enumerate(member_list) if item == ',']
            for index, l_index in enumerate(index_list):
                if index % 2 == 0:
                    continue
                member_list[l_index] = '\n'
        result_list.append(''.join(member_list))
    return '\n'.join(result_list)


def parse_time_degrees(time_degrees):
    '''
    解析项目进度与项目时间
    param time_degrees  项目时间与项目进度情况  type tuple or list
    return 格式化后的项目时间与项目项目进度
    '''
    degrees, start_time, end_time = time_degrees[2],time_degrees[3],time_degrees[4]
    start_time = re.findall(r"\d+", start_time)
    end_time = re.findall(r"\d+", end_time)
    time_str = f'{start_time[0]}.{start_time[1]}.{start_time[2]}-{end_time[0]}.{end_time[1]}.{end_time[2]}'
    return time_str,degrees

def parse_tasks(tasks):
    '''
    解析项目周报任务
    param tasks : 项目周报任务  type dict
    return 元组  归类后的项目数据
    '''
    p_names = []  # 项目名字
    p_outlines = []  # 项目概要
    p_risks = []  # 项目风险
    p_charges = []  # 项目经理
    p_stages = []  # 项目阶段
    p_members = []  # 项目成员
    p_processes = []  # 本周工作任务
    p_times = []  # 项目起止时间
    p_degrees = []  # 项目进度
    for k1, v1 in tasks.items():
        p_names.append(k1)
        for k2, v2 in v1.items():
            if k2 == '项目经理':
                p_charges.append(v2)
            if k2 == '项目里程碑':
                p_stages.append(parse_project_milestones(v2))
            if k2 == '项目问题及风险点':
                p_risks.append(parse_project_risks(v2))
            if k2 == '本周主要工作进展':
                p_processes.append(parse_project_processes(v2))
            if k2 == '当前阶段':
                times,degrees = parse_time_degrees(v2)
                p_times.append(times)
                p_degrees.append(degrees)
            if k2 == '项目概要':
                p_outlines.append(parse_project_outlines(v2))
            if k2 == '项目组成员':
                p_members.append(pares_project_members(v2))
    return p_names,p_times,p_degrees,p_outlines,p_risks,p_stages,p_charges,p_members,p_processes


TYPEFACE = 'SimSun'
TYPEFACEFILE = 'SimSun.ttf'

# 注册字体
pdfmetrics.registerFont(TTFont(TYPEFACE, TYPEFACEFILE))


def drawTitle(content, text=''):
    '''
    绘画标题
    param content  画布  type list
    param text   标题文本
    '''
    style = getSampleStyleSheet()
    ct = style['Normal']
    ct.fontName = TYPEFACE
    ct.fontSize = 18
    # 设置行距
    ct.leading = 50
    # 颜色
    ct.textColor = colors.black  
    # 居中 
    ct.alignment = 1
    # 添加标题并居中 
    content.append(Paragraph(text, ct))

def drawText(content, text=''):
    '''
    绘画文本内容
    param content  画布  type list
    param text   标题文本
    '''
    style = getSampleStyleSheet()
    # 常规字体(非粗体或斜体) 
    ct = style['Normal']
    # 使用的字体s 
    ct.fontName = TYPEFACE
    ct.fontSize = 14
    # 设置自动换行 
    ct.wordWrap = 'CJK'
    # 居左对齐 
    ct.alignment = 0
    # 第一行开头空格 
    ct.firstLineIndent = 32
    # 设置行距 
    ct.leading = 30
    content.append(Paragraph(text, ct)) 


def drawGentt(args):
    '''
    绘画甘特图
    param content  画布  type list
    param args  显示的数据   type list or tuple 
    '''
    p_names,p_times,p_degrees = args
    df = []
    now = datetime.datetime.now()
    for pn, pt,pd in zip_longest(p_names, p_times, p_degrees):
        y0, m0, d0 = pt.strip().split('-')[0].strip().split('.')
        y1, m1, d1 = pt.strip().split('-')[1].strip().split('.')
        mid_time = None
        percent = int(pd[:-1])
        if percent == 0:
            df.append(dict(Task = pn, Start = f'{y0}-{m0}-{d0}', Finish = f'{y1}-{m1}-{d1}', Resource='未完成'))
        if percent == 100:
            df.append(dict(Task = pn, Start = f'{y0}-{m0}-{d0}', Finish = f'{y1}-{m1}-{d1}', Resource='已完成'))
        if 0 < percent < 100:
            start_time = datetime.datetime(year=int(y0), month=int(m0), day=int(d0))
            end_time = datetime.datetime(year=int(y1), month=int(m1), day=int(d1))
            if now <= end_time: 
                datedelta = int((end_time - start_time).days * percent / 100)
            else:
                datedelta = int((end_time - start_time).days * percent / 100)
            mid_time = (start_time + datetime.timedelta(days=int(datedelta))).strftime("%Y-%m-%d")
            df.append(dict(Task=pn, Start=f'{y0}-{m0}-{d0}', Finish=f'{mid_time}', Resource='已完成'))
            df.append(dict(Task = pn, Start = f'{mid_time}', Finish = f'{y1}-{m1}-{d1}', Resource='未完成'))
    colors = {
            '已完成':'rgb(30,144,255)',
           '未完成': 'rgb(192,192,192)'}
 
    fig_gantt = ff.create_gantt(df, colors=colors, index_col='Resource', group_tasks=True, show_colorbar=True,
            showgrid_x=True, showgrid_y=True,title='项目进度图')
    gantt_jpg = 'static/fig_gantt.png'
    # 存储成png格式
    fig_gantt.write_image(gantt_jpg,width=1024)
    # 存储成pdf格式
    # fig_gantt.write_image("fig_gantt.pdf")


def drawTable(args):
    table = Table()
    headers = ["项目名称", "项目概要", "项目风险", "项目阶段", '项目经理', '项目成员', '工作进展']
    p_names, _, _, p_outlines, p_risks, p_stages, p_charges, p_members, p_processes = args
    rows = []
    lastWeek = deltaWeek()
    for item in zip_longest(p_names, p_outlines, p_risks, p_stages, p_charges, p_members, p_processes):
        rows.append(item)
    table.add(headers, rows).set_global_opts(
        title_opts=ComponentTitleOpts(title="\t项目周报详情",subtitle=f'\t周报日期:{lastWeek}',title_style={'text-indent':'2em'},subtitle_style={'text-indent':'2em'}) # subtitle="副主标题",
    )
    table_jpg = 'static/fig_table.png'
    # http://wkhtmltopdf.org 下载wkhtmltoimage
    config = imgkit.config(wkhtmltoimage='D:/Install-CustomSoftware/Install-wkhtmltopdf/wkhtmltopdf/bin/wkhtmltoimage.exe')
    imgkit.from_file(table.render(), table_jpg,config=config)   


def deltaWeek(weekdelta=-1):
    '''
    周的起止时间
    param weekdelta 偏移量    0本周, -1上周, 1下周
    return tuple : （周一,周日）
    '''
    now = datetime.datetime.now()
    week = now.weekday()
    _from = (now - datetime.timedelta(days=week - 7 * weekdelta)).date().strftime("%Y.%m.%d")
    _to = (now + datetime.timedelta(days=6 - week + 7 * weekdelta)).date().strftime("%Y.%m.%d")
    return f'{_from} - {_to}' 


def genLongImage(args):
    '''
    生成长图
    '''
    try:
        drawGentt(args[:3])
        drawTable(args)
        imgs = [Image.open(f'static/{fn}') for fn in listdir(path='static/') if fn.endswith('.png')]
        width, height = imgs[0].size
        result = Image.new(imgs[0].mode, (width, height * len(imgs)))
        for index, img in enumerate(imgs):
            result.paste(img, box=(0, index * height))
        result.save('pws.png')
    except Exception:
        return False
    return True


def parseweekmail(mails, server, projectManager):
    '''
    解析周报
    param mails 邮件内容
    param server 邮件服务
    param projectManager 项目经理列表或者元组
    '''
    # 获取最新一封邮件, 注意索引号从1开始:
    # print('共有{}封邮件'.format(len(mails)))
    for index in range(len(mails), 0, -1):
        resp, lines, octets = server.retr(index)
        # lines存储了邮件的原始文本的每一行,
        # 可以获得整个邮件的原始文本:
        msg_content = b'\r\n'.join(lines).decode('utf-8')
        # 稍后解析出邮件:
        msg = Parser().parsestr(msg_content)
        parseflag, overflag = select_by_mail_header(msg=msg, receiver='report_hz@wopuwulian.com',
                                 date_start=(datetime.datetime.now() - datetime.timedelta(weeks=2)).date(),projectManager=projectManager)
        if overflag:
            return 
        if parseflag:
            try:
                htmlcontent = get_mail_content(msg)
                if htmlcontent:
                    tasks = get_tasks(htmlcontent)
                    if tasks:
                        results = parse_tasks(tasks)
                        genLongImage(results)
            except Exception as e:
                print(e)
                continue  


def getMail():
    '''获取邮件'''
    # 输入邮件地址, 口令和POP3服务器地址:
    email = '*********'
    password = '*********'
    pop3_server = 'pop3.wopuwulian.com'

    # 项目经理列表
    projectManager = ('*********',)
    
    # 连接到POP3服务器:
    server = poplib.POP3(pop3_server)
    # 可以打开或关闭调试信息:
    # server.set_debuglevel(1)
    # 可选:打印POP3服务器的欢迎文字:
    # print(server.getwelcome().decode('utf-8'))

    # 身份认证:
    server.user(email)
    server.pass_(password)

    # stat()返回邮件数量和占用空间:
    # print('Messages: %s. Size: %s' % server.stat())
    # list()返回所有邮件的编号:
    resp, mails, octets = server.list()
    # 可以查看返回的列表类似[b'1 82923', b'2 2184', ...]
    # print(mails)

    parseweekmail(mails=mails, server=server,projectManager=projectManager)
    # 关闭连接:
    server.quit()


def run():
    '''运行'''
    getMail()
    # send_work_notice_file(file='pws.png')

if __name__ == '__main__':
    pass

    tasks ={
        '项目周报模板测试': {
            '项目经理': '骆年松',
            '项目里程碑': OrderedDict([
                ('启动', ('完成', '2020.01.04', '2020.01.04', '2020.01.07', '')),
                ('需求', ('完成', '2020.01.10', '2020.01.10', '2020.01.10', '')),
                ('UI设计', ('完成', '2020.01.13', '2020.01.13', '', '')),
                ('方案设计', ('进行中', '', '', '', 'UI评审后计划2020.01.14发布项目详细排期')),
                ('开发', ('', '待定', '', '', '')), ('测试', ('', '', '', '', '')),
                ('发布', ('', '', '', '', ''))
            ]),
            '项目较大变更说明': OrderedDict(),
            '项目问题及风险点': OrderedDict(),
            '本周主要工作进展': OrderedDict([
                ('项目需求评审', ('需求确认', '100%', '2020.01.04', '2020.01.07', '项目需求评审')),
                ('项目设计评审', ('开发设计', '100%', '2020.01.10', '2020.01.10', '项目设计评审'))
            ]),
            '下周工作计划': OrderedDict([
                ('项目UI评审', ('UI设计评审', '2020.01.13', '全体', ''))
            ]),
            '项目文档地址': OrderedDict([
                ('需求文档', ''),
                ('项目文档', 'http://doc.wpwl-inc.com/pages/viewpage.action?pageId=23888199'),
                ('接口文档', '')
            ]),
            '当前阶段': ('项目详细设计', '已完成/已评审', '50%', '2020.01.06', '2020.01.10', '2020.01.10完成设计评审'),
            '项目概要': '需求方代表：内部详述：项目周报的测试',
            '项目组成员': '产品：骆年松开发：柴纯洁、伊国锐、冯建齐、周罗朋测试：周园芳、柯玲飞设计：李取豪'
            }
    }


    # 生成 png 图片
    results = parse_tasks(tasks)
    # print(results)
    # drawTable(None,results)
    flag = genLongImage(results)
    # if flag: 
    #     pdf2Image()

    

    # # 发送钉钉png图片
    # send_work_notice_file(file='static/pws.png')
    # run()
    



#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time : 2024/2/18 16:17
"""
    目的：采集小红书作者主页的所有笔记内容
    分析思路：
    1、打开小红书主页与登录
    2、打开小红书作者主页
    3、提取页面笔记数据
    4、循环下滑页面刷新数据，循环获取笔记数据
    5、处理获取到的数据（去重，排序）
    6、保存到本地Excle文件
"""
import math
import os
import random
import time
import openpyxl
import pandas as pd
from DataRecorder import Recorder
from DrissionPage import ChromiumPage
from tqdm import tqdm


# 登录
def countdown(n):
    for i in range(n,0,-1):
        # \r让光标回到行首，end=''--结束符为空即不换行
        print(f'\r倒计时{i}秒',end='')
        time.sleep(1)
    else:
        print('\r倒计时结束')


# 只有第一次登录需要这个函数，以后因为免登录了这个函数无需执行
def sign_in():
    sign_in_page=ChromiumPage()
    sign_in_page.get('http://www.xiaohongshu.com')
    # 第一次运行需要扫码登录
    print('请扫码登录')
    # 倒计时30秒
    countdown(30)


#第二步
def open_url(url):
    global  page,user_name
    page=ChromiumPage()
    page.get(f'{url}')
    # 页面最大化
    page.set.window.max()
    # 定位作者信息
    user=page.ele('.info')
    # 作者名字
    user_name=user.ele('.user-name',timeout=0).text


# 第三步
def get_page_content():
    # 定位包含笔记信息的sections
    container=page.ele('.feeds-container')
    sections=container.eles('.note-item')
    # 笔记类型
    if sections.ele('.play-icon',timeout=0):
        note_type='视频'
    else:
        note_type='图文'
    # 文章链接
    note_link=sections.ele('tag:a',timeout=0).link
    # 标题
    title=sections.ele('.title',timeout=0).text
    # 作者
    author=sections.ele('.author-wrapper')
    # 点赞
    like=author.ele('.count').text
    # notes列表存放当前页面的笔记
    notes=[]
    note={
        '作者':author,
        '笔记类型':note_type,
        '标题':title,
        '点赞数':like,
        '笔记链接':note_link
    }
    notes.append(note)
    # 使用DataRecorder库保证程序异常情况下，已经获取的数据也不会丢失
    r.add_data(notes)


# 第四步
def page_scroll_down():
    print(f'********下滑页面********')
    page.scroll.to_bottom()
    # 生成一个1~2秒随机时间
    random_time=random.uniform(1,2)
    time.sleep(random_time)


# 设置向下翻页爬取次数
def crawler(times):
    global i
    for i in tqdm(range(1,times+1)):
        get_page_content()
        page_scroll_down()

# 第五步
def re_save_excle(filepath):
    df=pd.read_excel(filepath)
    print(f'总计向下翻页{times}次，获取{df.shape[0]}条笔记（包含重复获取）。')
    df['点赞数']=df['点赞数'].astype(int)
    # 删除重复行
    df=df.drop_duplicates()
    # 按照点赞数降序
    df=df.sort_values(by='点赞数',ascending=False)
    # 文件路径
    final_filepath=f'小红书作者主页所有笔记-{author}--{df.shape[0]}条.xlsx'
    df.to_excel(final_filepath,index=False)
    print(f'总计向下翻页{times}次，笔记去重后剩余{df.shape[0]}条，保存到文件：{final_filepath}。')
    print(f'数据已保存到：{final_filepath}')


# 自动调整excle列宽
def auto_resize_column(excle_path):
    wb=openpyxl.load_workbook(excle_path)
    worksheet=wb.active
    for col in worksheet.iter_cols(min_col=1,max_col=5):
        max_length=0
        column=col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length=len(str(cell.value))
            except:
                pass
        # 计算调整后的列宽
        adjusted_width=max_length+5
        # 使用worksheet.column_dimensions属性设置列宽度
        worksheet.column_dimensions[column].width = adjusted_width
    wb.save(excle_path)
    wb.close()


# 删除初始excle文件
def delete_file(filepath):
    # 检查文件是否存在
    if os.path.exists(filepath):
        # 删除文件
        os.remove(filepath)
        print(f'已删除初始化excle文件：{filepath}')
    else:
        print(f'文件不存在：{filepath}')

if __name__ == '__main__':
    # 1.第一次运行需要登陆，需要执行sign_in函数。第二次之后不用登陆，可以注释掉此步骤
    sign_in()
    # 2.设置主页地址url
    author_url='http://www.xiaohongshu.com/user/profile/60bc91d400000000100497f'
    # 3.设置向下翻页爬取次数
    # 根据小红书作者主页'当前发布笔记数'计算浏览器下滑次数
    # 计算向下滑动页面次数的方法如下
    """
    :param:note_num是笔记数量
    """
    note_num = 62
    times = math.ceil(note_num / 20 * 1.1)
    print(f'需要执行翻页次数为：{times}')
    # 4.设置要保存的文件名filepath
    current_time = time.localtime()
    formatted_time = time.strftime("%Y-%m-%d %H%M%S", current_time)
    init_file_path = f'小红书作者主页所有笔记-{formatted_time}.xlsx'
    r = Recorder(path=init_file_path, cache_size=100)

    # 打开主页
    open_url(author_url)
    # 根据设置的次数开始爬取数据
    crawler(times)
    # 避免数据丢失，爬虫结束强制保存数据
    r.record()
    # 结束
    re_save_excle(init_file_path)









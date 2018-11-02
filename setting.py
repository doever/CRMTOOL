#!/usr/bin/python3
# -*- coding:utf-8 -*-

import os
import sys


BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

TITLE = 'CRM TOOLS'

PAGE_HEIGHT = 585
PAGE_WIDTH = 580

COLOR = {
        'theme_color': '#5C87d9',
        'font_color': '#212121',
        'line_color': '#eee',
        'button_color': '#5bc0de',
        'link_color': '#878787',
        'warning': '#f0ad4e',
        'danger': '#d9534f'
    }

FONT_SIZE = {
        'title_lg': 24,
        'title_sm': 18,
        'title_xs': 16,
        'text': 14,
        'link': 12,
        'assistant': 10,
    }

MENU_NAME = {
        'page_one': '首页',
        'page_two': '浩泽撤单',
        'page_three': '浩优单据',
        'page_four': '定时邮件',
        'page_five': '替换工具',
        'page_six': '数据导出',
        'page_seven': '数据监控',
        }

INDEX_CONFIG = {
    'logo': 'OZNER',
    'title': '浩泽CRM工具',
    'background': os.path.join(BASE_DIR, 'image/background.png'),
    'default_background': os.path.join(BASE_DIR, 'image/default.png')
}

# -*- coding: utf-8 -*-

from __future__ import (print_function, unicode_literals, with_statement,
                        absolute_import, division)

"""
PlantSpecimenInfoInput
======================

Introduction
------------

Input plant specimen informations from xlsx file and Internet
and write to xlsx files automatically.

Dependencies
------------

- requests
- BeautifulSoup4
- openpyxl

Usage
-----
1. Quick use for people who are not familiar with commands:

    Change the names of your files to:

    1. query.xlsx   (query_file)
    2. data.xlxs    (data_file)

    Then, type this in console or Windows cmd:

        $ python specimen_input

2. Common use:

        $ python specimen_input.py \
            -i query.xlsx \
            -d data.xlsx \
            -o outfile.xslx
"""

import re
import os
try:
    import bs4
except ImportError:
    bs4 = None
import sys
import time
import json
import logging
try:
    import openpyxl
except ImportError:
    openpyxl = None
try:
    import requests
except ImportError:
    requests = None
import argparse
import threading
from collections import namedtuple
from multiprocessing.dummy import Pool

if sys.version[0] == '2':
    import Tkinter as tk
    import ttk
    import ScrolledText as st
    import Queue
elif sys.version[0] == '3':
    import tkinter as tk
    from tkinter import ttk
    import tkinter.scrolledtext as st
    import queue as Queue
else:
    raise ImportError('Cannot identify your Python version.')

__version__ = "v1.3.0"

__all__ = ['Query', 'write_to_xlsx_file', 'gui_main']

# ==================================================
# You can change settings here if needed
# ==================================================
POOL_NUM = 30
MAX_ERROR_NUM = 50

LIBRARY_CODE = "FUS"
COLLECTION_COUNTRY = "中国"

# query 文件的列标题
QUERY_FILE_HEADER_TUPLE = {
    "物种编号",
    "流水号",
    "条形码",
    "物种名（二名法）",
    "同一物种编号"
}

# data 文件的列标题
DATA_FILE_HEADER_TUPLE = (
    "物种编号",
    "中文名",
    "种名（拉丁）",
    "科名",
    "科名（拉丁）",
    "省",
    "市",
    "具体小地名",
    "纬",
    "东经",
    "海拔",
    "日期",
    "份数",
    "草灌",
    "采集人",
    "鉴定人",
    "鉴定日期",
    "录入员",
    "录入日期"
)

# output 文件的列标题
HEADER_TUPLE = (
    "馆代码", "流水号", "条形码", "模式类型", "库存", "标本状态",
    "采集人", "采集号", "采集日期", "国家", "省市", "区县", "海拔",
    "负海拔", "科", "属", "种", "定名人", "种下等级", "中文名",
    "鉴定人", "鉴定日期", "备注", "地名", "生境", "经度", "纬度",
    "备注2", "录入员", "录入日期", "习性", "体高", "胸径", "茎",
    "叶", "花", "果实", "寄主")

# ==================================================
# Be careful if you want to change values below
# ==================================================

# If no content, this is the number of blank tuple
TOTAL_LINES = 38

# Local JSON cache file name for web search
LOCAL_JSON_CACHE_FILE = 'cache.json'

# Dictionaries used for cache
# Web data cache
_web_data_cache_dict = {}
# xlsx data cache
_xlsx_data_cache_dict = {}

# For fancy display
BAR = '\n' + '=' * 60 + '\n'
THIN_BAR = '\n' + '-' * 60 + '\n'
LINE_SPLITER = '-' * 50
THIN_BAR_NO_NEWLINE = '-' * 60

HELP = """
PlantSpecimenInfoInput

Introduction

    Input plant specimen informations from xlsx file and Internet
    and write to xlsx files automatically.

Dependencies

    - requests
    - BeautifulSoup4
    - openpyxl

Usage

    1. Quick usage for people who are not familiar with commands:

        a) Install Python

        b) Change the names of your files to:

            1. query.xlsx   (query_file)
            2. data.xlxs    (data_file)

        c) Then, type this in console or Windows cmd:

            $ python specimen_input

    2. Advanced usage:

        a) After installing Python, type this command in console
           or Windows cmd:

            $ python specimen_input.py -i query.xlsx -d data.xlsx -o outfile.xslx
"""

DEFAULT_LATIN_NAME_FILE = os.path.join('.', 'data', 'latin_names.txt')
DEFAULT_LATIN_NAME_FILE_2 = os.path.join('.', 'data',
                                         'latin_names_only_head_and_tail.txt')

# logging
file_handler_format = ('%(message)s')
logging.basicConfig(level=logging.DEBUG,
                    format=file_handler_format,
                    datefmt="%Y-%m-%d %H:%M",
                    filename="log.txt",
                    filemode="w")

# logging handler for displaying output to screen
console = logging.StreamHandler()
console.setLevel(logging.INFO)
formatter = logging.Formatter("%(message)s")
console.setFormatter(formatter)

logging.getLogger("").addHandler(console)

# Seppress logging info from urllib3 which was called by requests
logging.getLogger("requests").setLevel(logging.WARNING)
logging.getLogger("urllib3").setLevel(logging.WARNING)


def check_unicode(unknown):
    """Check if unknown type is unicode."""
    return isinstance(unknown, unicode)


class XlsxFile(object):
    """
    Handel xlsx files and return a matrix of content.
    """

    def __init__(self, excel_file):
        try:
            self.wb = openpyxl.load_workbook(excel_file)
        # Invalid xlsx format
        except openpyxl.utils.exceptions.InvalidFileException as e:
            logging.error("[ ERROR ] Invalid xlsx format.\n%s" % e)
            raise ValueError('[ ERROR ] Invalid xlsx format: {}'.format(e))
        except IOError as e:
            logging.error("[ ERROR ] No such xlsx file: %s. (%s)" % (excel_file, e))
            raise ValueError('[ ERROR ] IOError: {}. {}'.format(excel_file, e))
        except BaseException as e:
            logging.error(e)
            raise ValueError('[ Error ] {}'.format(e))

        self.ws = self.wb.active
        if not self.ws:
            logging.error("[ ERROR ] 无法获取 xlsx 文件中的 active sheet：{}".format(excel_file))
            raise ValueError("[ ERROR ] 无法获取 xlsx 文件中的 active sheet：{}".format(excel_file))
        self.ws_title = self.ws.title
        self.xlsx_matrix = []
        self.species_info_dict = {}
        self._get_matrix()

    @property
    def all_sheet_names(self):
        """Get a list of sheet names from that xlsx file."""
        return self.wb.get_sheet_names()

    def load_specific_sheet(self, sheet_name):
        """Specify the sheet name you want to open."""
        if sheet_name not in self.all_sheet_names:
            logging.error("No such sheet in xlsx file: %s" % sheet_name)
            sys.exit(1)
        else:
            logging.info("[ Load Sheet by Name  ]:  Openning sheet...")
            self.ws = self.wb.get_sheet_by_name(sheet_name)
            self.ws_title = self.ws.title

    def load_sheet_by_index(self, index_num=1):
        """Load sheet by the index of sheet in xlsx file."""
        try:
            index_num = int(index_num)
        except:
            error_msg = ("Illegal index_num value:  %s, must be "
                         "0 < index_num < sheet_num." % index_num)
            logging.error(error_msg)
            raise ValueError(error_msg)
        logging.info("[ Load Sheet by Index ]:  Choose sheet No. %d "
                     % index_num)
        if index_num < len(self.all_sheet_names) and index_num >= 0:
            self.load_specific_sheet(self.all_sheet_names[index_num])
        else:
            error_msg = "Invalid sheet index number: %d" % index_num
            logging.error(error_msg)
            raise ValueError(error_msg)

    def _get_matrix(self):
        """Get a two dimensional matrix from the xlsx file."""
        self.xlsx_matrix = []
        for i, row in enumerate(self.ws.rows):
            row_container = []
            for i, cell in enumerate(row):
                row_container.append(cell.value)
            self.xlsx_matrix.append(tuple(row_container))

    def get_xlsx_data_dict(self, key_column_index):
        """Return a dictionary with data from xlsx matrix.

        Key:     The Nth elements (namely: key_column_index).
        Value:   a list of all elements

        if matrix = [('1', 'a', 'x'), ('2', 'b', 'z'), ('3', 'c', 'y')]
        set key_column_index=1,
        return: {
            'a': ('1', 'a', 'x'),
            'b': ('2', 'b', 'z'),
            'c': ('3', 'c', 'y')
        }
        """
        xlsx_data_dict = {}
        if not self.xlsx_matrix:
            self._get_matrix()
        for i, row_tuple in enumerate(self.xlsx_matrix):
            if not row_tuple[key_column_index]:
                continue
            elements = [_.strip() if type(_) == str else _
                        for _ in row_tuple]
            # Add key=species_name : value=info_list to dictionary
            # use " ".join(species_name.split()) to avoid search failure by
            # format error (If there are more than one blanks or tabs)
            species_name = " ".join(elements[key_column_index].split())
            xlsx_data_dict[species_name] = tuple(elements)

        return xlsx_data_dict


class QueryParser(object):
    """Parse query file and return a list of query tuples.

    >>> query = QueryParser(query_file)
    >>> query_tuple = query.query_tuple
    """

    def __init__(self, query_file):
        if not query_file:
            error_msg = "No such query file: %s" % query_file
            logging.error(error_msg)
            raise IOError(error_msg)
        self._query_xlsx_file = XlsxFile(query_file)

    @property
    def query_tuple(self):
        return self._query_xlsx_file.xlsx_matrix


class WebInfo(object):
    """Web crawler class. Get info from Internet.

    >>> w = WebInfo("Eupatorium coelestinum")
    >>> web_info_tuple = w.pretty_info_tuple
    """

    def __init__(self, species_name):
        self.species_name = species_name
        self.response = None
        self._cook_soup()

    def _cook_soup(self):
        """Prepare requests response and BeautifulSoup soup."""
        logging.info("    [ Species ]  %s" % self.species_name)
        if len(self.species_name.split()) == 2:
            genus, species = self.species_name.split()
        else:
            logging.warning("    [ WARNING ]  Is this llegal species name?"
                            " -->  %s" % self.species_name)
            genus, blank, species = [
                _.strip() for _ in self.species_name.partition(' ')]

        requests_url = ('http://frps.eflora.cn/frps/'
                        + genus
                        + '%20'
                        + species)
        try:
            self.response = requests.get(requests_url).text
            self.soup = bs4.BeautifulSoup(self.response, "html.parser")
        except bs4.FeatureNotFound as e:
            logging.error(" *  Cannot find parser: html.parser.")
            logging.error(" *  You may need to use lxml or html5lib")
            logging.error("        pip install lxml")
            logging.error("        pip install html5lib")
            raise bs4.FeatureNotFound()
        except requests.ConnectionError as e:
            logging.error(
                ' *  Internet Connection Failed.'
                '    Output will only get data form date file.\n    %s' % e)
            self.soup = None
        except BaseException as e:
            logging.error(" *  %s" % e)
            sys.exit(1)

    @property
    def all_paragraph_tuple(self):
        """All paragraphes in the website with <p> tags."""
        if not self.soup:
            return None
        paragraphe_tuple_list = [p.find(text=True)
                                 for p in self.soup.select('p')]
        return paragraphe_tuple_list

    @staticmethod
    def _find_keyword_info(one_paragraph_content):
        """From <p> taged paragraphes, try to extact informations that has
        relevant keywords."""
        re_1 = re.compile('[^，。]*高[^，。]*')
        height_list = ' | '.join(re_1.findall(one_paragraph_content))  # 体高

        re_2 = re.compile('[^，。]*胸径[^，。]*')
        DBH_list = ' | '.join(re_2.findall(one_paragraph_content))  # 胸径, DBH

        re_3 = re.compile('[^。]*茎[^。]*')
        stem_list = '。 | '.join(re_3.findall(one_paragraph_content))  # 茎

        re_4 = re.compile('[^。]*叶[^。]*')
        leaf_list = '。 | '.join(re_4.findall(one_paragraph_content))  # 叶

        re_5 = re.compile('[^。]*花[^。]*')
        flower_list = '。 | '.join(re_5.findall(one_paragraph_content))  # 花

        re_6 = re.compile('[^。]*果[^。]*')
        fruit_list = '。 | '.join(re_6.findall(one_paragraph_content))  # 果实

        re_7 = re.compile('[^。]*寄主[^。]*')
        host_list = '。 | '.join(re_7.findall(one_paragraph_content))  # 寄主

        # Return a tuple
        # 0. 高
        # 1. 胸径, DBH
        # 2. 茎
        # 3. 叶
        # 4. 花
        # 5. 果
        # 6. 寄主
        return (height_list, DBH_list, stem_list, leaf_list,
                flower_list, fruit_list, host_list)

    def _get_target_info(self):
        """Search infos with specific keywords."""
        if not self.all_paragraph_tuple:
            return tuple(["" for _ in xrange(7)])
        paragraphe_tuple_list = self.all_paragraph_tuple

        strict_word_tuple = ['高', '茎', '叶', '花', '果']
        moderate_word_tuple = ['叶', '花']
        # For example: Gymnospermae (裸子植物)
        relaxed_word_tuple = ['茎', '叶']

        detailed_paragraph = ''
        for each_paragraph in paragraphe_tuple_list:
            # Check if this paragraph is the main description graph.
            if all(word in each_paragraph for word in strict_word_tuple) \
                    or all(word in each_paragraph
                           for word in moderate_word_tuple) \
                    or all(word in each_paragraph
                           for word in relaxed_word_tuple):
                detailed_paragraph = each_paragraph
                break

        if not detailed_paragraph:
            (height_list, DBH_list, stem_list, leaf_list,
             flower_list, fruit_list, host_list) = \
                tuple(["" for _ in xrange(7)])
        else:
            # try:
            (height_list, DBH_list, stem_list, leaf_list,
             flower_list, fruit_list, host_list) = \
                self._find_keyword_info(detailed_paragraph)
            # except UnicodeEncodeError as e:
            #     height = DBH = stem = leaf = flower = fruit = host = ''

        return (height_list, DBH_list, stem_list, leaf_list,
                flower_list, fruit_list, host_list)

    @property
    def pretty_info_tuple(self):
        """Format infos from web."""

        # Get genus, species, namer
        if not self.species_name:
            return ['' for x in range(11)]
        if len(self.species_name.split()) >= 2:
            genus, species = self.species_name.split()[0], \
                             self.species_name.split()[1]
        else:
            genus, species = self.species_name, ''
        re_namer = re.compile('(?<=<b>%s</b> <b>%s</b>)[^><]*(?=<span)'
                              % (genus, species))
        try:
            namer = re_namer.findall(self.response)[0].strip()
        except IndexError as e:
            logging.error("  * [ WARNING ]  无法从网络获取命名人：{}".format(self.species_name))
            namer = ""

        # Get habitat (TODO.)
        habitat = ''

        # Get height, DBH, stem, leaf, flower, fruit, host
        try:
            (height_list, DBH_list, stem_list, leaf_list,
             flower_list, fruit_list, host_list) = \
                self._get_target_info()
        except Exception as e:
            logging.error(
                'Cannot get height, DBH, stem, ... for %s. (%s)' %
                (self.species_name, e))
            (height_list, DBH_list, stem_list, leaf_list,
             flower_list, fruit_list, host_list) = ['' for x in range(7)]

        web_info_tuple = (
            genus, species, namer,
            habitat,
            height_list, DBH_list, stem_list, leaf_list,
            flower_list, fruit_list, host_list)

        return web_info_tuple


class WebInfoCacheMultithreading(object):
    def __init__(self, query_file):
        self.query_file = query_file
        self.non_repeatitive_species_name_list = \
            self._get_non_repeatitive_species_name_list()

    def _get_non_repeatitive_species_name_list(self):
        query_tuple_list = QueryParser(self.query_file).query_tuple
        non_repeatitive_species_name_list = list(
            set([_[3] for _ in query_tuple_list]))

        return non_repeatitive_species_name_list

    def _single_query(self, one_species_name):
        global _web_data_cache_dict

        try:
            pretty_info_tuple = WebInfo(one_species_name).pretty_info_tuple
            _web_data_cache_dict[one_species_name] = pretty_info_tuple
        except Exception as e:
            logging.error('Cannot get info from web: %s (%s)' %
                          (one_species_name, e))

    def get_web_dict_multithreading(self):
        if POOL_NUM > 1 and POOL_NUM < 50:
            pool = Pool(POOL_NUM)
            logging.info("You are using multiple threads to get info from web:"
                         "  [ %d ]\n" % POOL_NUM)
        else:
            pool = Pool()
        if os.path.isfile(LOCAL_JSON_CACHE_FILE):
            with open(LOCAL_JSON_CACHE_FILE, 'rb') as f:
                local_web_cache_dict = json.loads(f.read())
            species_in_local_json_cache = local_web_cache_dict.keys()
            logging.info(
                '[ CACHE ] Get cache from local JSON file:\n  |- %s' %
                '\n  |- '.join(species_in_local_json_cache))
        else:
            species_in_local_json_cache = []
            local_web_cache_dict = {}
        species_not_in_cache = list(
            set(self.non_repeatitive_species_name_list).difference(
                set(species_in_local_json_cache)))
        pool.map(self._single_query, species_not_in_cache)
        _web_data_cache_dict.update(local_web_cache_dict)
        with open(LOCAL_JSON_CACHE_FILE, 'wb') as f:
            json.dump(_web_data_cache_dict, f,
                      indent=4, separators=(',', ': '))
            logging.info(
                '[ CACHE ] Write all cache to local JSON file:\n  |- %s' %
                '\n  |- '.join(_web_data_cache_dict.keys()))
        pool.close()
        pool.join()


class OfflineDataCache(object):
    def __init__(self, offline_data_file):
        self.offline_data_file = offline_data_file

    def get_xlsx_data_dict(self):
        global _xlsx_data_cache_dict
        _xlsx_data_cache_dict = XlsxFile(
            self.offline_data_file).get_xlsx_data_dict(key_column_index=0)


def get_cache(query_file, offline_data_file):
    """Generate cache dictionary for web info and offline info."""
    # Web Cache
    if not _web_data_cache_dict:
        WebInfoCacheMultithreading(query_file).get_web_dict_multithreading()

    # Offline Cache
    if not _xlsx_data_cache_dict:
        OfflineDataCache(offline_data_file).get_xlsx_data_dict()


class Query(object):
    """Do query for one query line and return orderd infos.

    >>> q = Query(
    ...    ("113678", "00098484", u"Stellaria media", "1"),
    ...    xlsx_data_dict)
    >>> out_tuple = q._formatted_single_output()
    """

    def __init__(self, query_file, offline_data_file):
        self.query_file = query_file
        self.offline_data_file = offline_data_file
        self.query_tuple_list = QueryParser(query_file).query_tuple

    def _do_single_raw_query(self, one_query_tuple):
        """Do query for one species and get raw results."""
        global _web_data_cache_dict
        global _xlsx_data_cache_dict

        collection_id_prefix, serial_number, barcode, species_name, same_species_num = one_query_tuple
        if not species_name:
            return ['' for x in range(11)], None
        species_name = " ".join(species_name.split())

        # ===============================================================
        # Web Crawler Cache
        # ===============================================================
        if species_name in _web_data_cache_dict:
            web_info_tuple = _web_data_cache_dict[species_name]
        else:
            if len(one_query_tuple[3].split()) >= 2:
                web_info_tuple = tuple([
                                           one_query_tuple[3].split()[0],
                                           ' '.join(one_query_tuple[3].split()[1:])]
                                       + ['' for x in range(9)])
            else:
                web_info_tuple = tuple([one_query_tuple[3]]
                                       + ['' for _ in range(10)])

        # ===============================================================
        # Offline Data Cache
        # ===============================================================
        if collection_id_prefix in _xlsx_data_cache_dict:
            offline_info_tuple = _xlsx_data_cache_dict[collection_id_prefix]
        else:
            offline_info_tuple = None

        return web_info_tuple, offline_info_tuple

    def _formatted_single_output(self, one_query_tuple):
        """Format raw results for single query."""
        web_info_tuple, offline_info_tuple = \
            self._do_single_raw_query(one_query_tuple)
        FinalInfo = namedtuple(
            "FinalInfo",
            [
                "library_code",  # 0.  馆代码
                "serial_number",  # 1.  流水号
                "barcode",  # 2.  条形码
                "pattern_type",  # 3.  模式类型
                "inventory",  # 4.  库存
                "specimen_condition",  # 5.  标本状态
                "collectors",  # 6.  采集人
                "collection_id",  # 7.  采集号
                "collection_date",  # 8.  采集日期
                "collection_country",  # 9.  国家
                "province_and_city",  # 10. 省市
                "county",  # 11. 区县
                "altitude",  # 12. 海拔
                "negative_altitude",  # 13. 负海拔
                "family",  # 14. 科
                "genus",  # 15. 属
                "species",  # 16. 种
                "namer",  # 17. 定名人
                "level",  # 18. 种下等级
                "chinese_name",  # 19. 中文名
                "identifier",  # 20. 鉴定人
                "identify_date",  # 21. 鉴定日期
                "remarks",  # 22. 备注
                "place_name",  # 23. 地名
                "habitat",  # 24. 生境
                "longitude",  # 25. 经度
                "latitude",  # 26. 纬度
                "remarks_2",  # 27. 备注2
                "inputer",  # 28. 录入员
                "input_date",  # 29. 录入日期
                "habit",  # 30. 习性
                "body_height",  # 31. 体高
                "DBH",  # 32. 胸径
                "stem",  # 33. 茎
                "leaf",  # 34. 叶
                "flower",  # 35. 花
                "fruit",  # 36. 果实
                "host"  # 37. 寄主
            ])

        # =======================================================
        # Offline info tuple
        #
        # 0.  物种编号
        # 1.  种名
        # 2.  种名（拉丁）
        # 3.  科名
        # 4.  科名（拉丁）
        # 5.  省
        # 6.  市
        # 7.  具体小地名
        # 8.  纬度
        # 9.  东经
        # 10. 海拔
        # 11. 采集日期
        # 12. 份数
        # 13. 草灌
        # 14. 采集人
        # 15. 鉴定人
        # 16. 鉴定日期
        # 17. 录入员
        # 18. 录入日期
        #
        if not offline_info_tuple:
            offline_info_tuple = tuple(['' for _ in range(TOTAL_LINES)])
        # =======================================================

        # =======================================================
        # Web info tuple
        #
        # 0.  genus
        # 1.  species
        # 2.  namer
        # 3.  habitat   # 生境
        # 4.  height
        # 5.  DBH       # 胸径
        # 6.  stem
        # 7.  leaf
        # 8.  flower
        # 9.  fruit
        # 10. host
        #
        # if not web_info_tuple:
        #     web_info_tuple = tuple(['' for _ in xrange(TOTAL_LINES)])
        # =======================================================

        # Get values for each entry
        library_code = LIBRARY_CODE
        pattern_type = ''
        inventory = ''
        specimen_condition = ''
        collection_country = COLLECTION_COUNTRY
        county = ''
        negative_altitude = ''
        level = ''
        remarks = ''
        remarks_2 = ''

        # Infos from qeury file
        try:
            serial_number = one_query_tuple[0]
            barcode = str(one_query_tuple[1]).zfill(8)
        except IndexError as e:
            error_msg = "Illegal query file format.\n%s" % e
            logging.error(error_msg)
            raise IndexError(error_msg)

        # Infos from offline data file
        try:
            collection_id = "%s-%s" % (
                offline_info_tuple[0], one_query_tuple[3])
            chinese_name = offline_info_tuple[1]
            family = offline_info_tuple[4]
            province_and_city = "%s,%s" % offline_info_tuple[5:7]
            place_name = offline_info_tuple[7]
            longitude = offline_info_tuple[9]
            latitude = offline_info_tuple[8]
            altitude = offline_info_tuple[10]
            collection_date = offline_info_tuple[11]
            habit = offline_info_tuple[13]
            collectors = offline_info_tuple[14]
            identifier = offline_info_tuple[15]
            identify_date = offline_info_tuple[16]
            inputer = offline_info_tuple[17]
            input_date = offline_info_tuple[18]
        except IndexError as e:
            error_msg = "Illegal offline data format.\n%s" % e
            logging.error(error_msg)
            raise IndexError(error_msg)

        # Infos from web
        try:
            genus = web_info_tuple[0] if web_info_tuple[0] \
                else one_query_tuple[2].split()[0]
            species = web_info_tuple[1] if web_info_tuple[1] \
                else ' '.join(one_query_tuple[2].split()[1:])
            namer = web_info_tuple[2]
            habitat = web_info_tuple[3]
            body_height = web_info_tuple[4]
            DBH = web_info_tuple[5]
            stem = web_info_tuple[6]
            leaf = web_info_tuple[7]
            flower = web_info_tuple[8]
            fruit = web_info_tuple[9]
            host = web_info_tuple[10]
        except Exception as e:
            # logging.warning("Skip... Cannot get info from web for:  %s. %s" %
            #                 (one_query_tuple[2], e))
            name = one_query_tuple[3]
            genus = name.split()[0] if name else ''
            species, namer, habitat, body_height, DBH, stem, leaf, \
            flower, fruit, host = ['' for x in range(10)]

        f = FinalInfo(
            library_code=library_code,
            serial_number=serial_number,
            barcode=barcode,
            pattern_type=pattern_type,
            inventory=inventory,
            specimen_condition=specimen_condition,
            collectors=collectors,
            collection_id=collection_id,
            collection_date=collection_date,
            collection_country=collection_country,
            province_and_city=province_and_city,
            county=county,
            altitude=altitude,
            negative_altitude=negative_altitude,
            family=family,
            genus=genus,
            species=species,
            namer=namer,
            level=level,
            chinese_name=chinese_name,
            identifier=identifier,
            identify_date=identify_date,
            remarks=remarks,
            place_name=place_name,
            habitat=habitat,
            longitude=longitude,
            latitude=latitude,
            remarks_2=remarks_2,
            inputer=inputer,
            input_date=input_date,
            habit=habit,
            body_height=body_height,
            DBH=DBH,
            stem=stem,
            leaf=leaf,
            flower=flower,
            fruit=fruit,
            host=host
        )

        return f

    def do_multi_query(self):
        """Do multiple query."""
        out_tuple_list = []

        logging.info("{}程序需要先从互联网查询所有物种的详细信息，这可能需要一些时间，请耐心等待...{}".format(THIN_BAR, THIN_BAR))

        # Generate global cache dict for web and offline data
        get_cache(self.query_file, self.offline_data_file)

        logging.info("\n{}开始处理每一个物种 ...{}".format(THIN_BAR, THIN_BAR))

        # Do query for each entry
        log_info = list()
        for i, each_query_tuple in enumerate(self.query_tuple_list):
            if (i + 1) % 10 == 0:
                logging.info('  {} - {} Done!'.format(i - 9, i + 1))
            log_info.append(
                "[ {} ] {}\n\n".format(i + 1, each_query_tuple[2]) +
                "       采集号（{}）\n".format(each_query_tuple[0]) +
                "       流水号（{}）\n".format(each_query_tuple[1]) +
                "       条形码（{}）\n".format(str(each_query_tuple[2]).zfill(8)) +
                "       物种名（{}）\n\n".format(each_query_tuple[3])
            )
            out_tuple = self._formatted_single_output(each_query_tuple)
            out_tuple_list.append(out_tuple)

        return out_tuple_list, '\n'.join(log_info)


def write_to_xlsx_file(out_tuple_list, xlsx_outfile_name="out.xlsx"):
    """Write tuple list to xlsx file.

    >>> write_to_xlsx_file([('a', 'b', 'c'), ('e', 'f', 'g')])

    +-----+-----+-----+
    |  a  |  b  |  c  |
    +-----+-----+-----+
    |  e  |  f  |  g  |
    +-----+-----+-----+
    """
    out_wb = openpyxl.Workbook()

    ws1 = out_wb.active
    ws1.title = "Specimen"

    # Header
    ws1.append(HEADER_TUPLE)

    # Content
    for i, tuple_row in enumerate(out_tuple_list):
        ws1.append(tuple_row)
    try:
        out_wb.save(filename=xlsx_outfile_name)
        logging.info("{}[ xlsx File ]  结果已写入文件：{}{}".format(
            THIN_BAR, xlsx_outfile_name, THIN_BAR))
    except IOError as e:
        basename, dot, ext = xlsx_outfile_name.rpartition(".")
        alt_xlsx_outfile = "%s.alt.%s" % (basename, ext)
        logging.error(" *  无权限写入文件：{}，当前是否处于打开状态而被占用？{}".format(xlsx_outfile_name, e))
        out_wb.save(filename=alt_xlsx_outfile)
        logging.info("\n{}[ xlsx File ]  结果已写入临时文件：{}{}".format(THIN_BAR, alt_xlsx_outfile, THIN_BAR))


def data_validation(data_file, query_file):
    """Validate data and query files before program run.

    This will save time. If there is error in data file, program may crush
    after long time run. So it's better to validate data file before running.
    """
    error_list = list()

    # Get file tuple list
    data_file_tuple_list = XlsxFile(data_file).xlsx_matrix
    query_file_tuple_list = XlsxFile(query_file).xlsx_matrix

    logging.info("{}数据校验开始 ...".format(THIN_BAR))

    critical_error = False

    # Check if latin name is missing in data file
    error_list.append(THIN_BAR_NO_NEWLINE)
    error_list.append('[ Start ] 检查 data 文件（{}）中的 latin 名是否有缺失 ... '.format(data_file))
    latin_names_in_data_file = []
    for i, each_tuple in enumerate(data_file_tuple_list[1:]):
        # If line is blank line, skip
        try:
            if not each_tuple[0] and not each_tuple[2]:
                continue
        except Exception as e:
            error_list.append('Line %s: %s' % (i + 1, e))
        if each_tuple[2]:
            latin_name = each_tuple[2]
            if len(latin_name.split()) < 2:
                error_list.append('  （{} 行）latin 名需要至少包含: genus + species：{}'.format(i + 2, latin_name))
                critical_error = True
            latin_names_in_data_file.append(each_tuple[2].strip())
        else:
            error_list.append('  （{} 行）latin 名缺失'.format(i + 1, data_file))
            critical_error = True

    if len(error_list) > MAX_ERROR_NUM:
        error_list = error_list[:MAX_ERROR_NUM]
        error_list.append('... ...')
        error_list.append('错误过多（大于 {} 个），是否文件错误？'.format(MAX_ERROR_NUM))
        logging.error('\n'.join(error_list))
        return False

    # Check if latin name is missing in query file
    logging.info(THIN_BAR_NO_NEWLINE)
    error_list.append('[ Start ] 检查 query 文件（{}）中的 latin 名是否有缺失 ... '.format(query_file))
    latin_names_in_query_file = []
    for i, each_tuple in enumerate(query_file_tuple_list):
        # If line is blank line, skip
        try:
            if not each_tuple[0] and not each_tuple[2]:
                continue
        except Exception as e:
            error_list.append('Line %s: %s' % (i + 1, e))
        if each_tuple[3]:
            latin_name = each_tuple[3]
            if len(latin_name.split()) < 2:
                error_list.append('  （{} 行）latin 名需要至少包含: genus + species：{}'.format(i + 1, latin_name))
                critical_error = True
            latin_names_in_query_file.append(latin_name.strip())
        else:
            error_list.append('  （{} 行）latin 名缺失'.format(i + 1, query_file))
            critical_error = True

    if len(error_list) > MAX_ERROR_NUM:
        error_list = error_list[:MAX_ERROR_NUM]
        error_list.append('... ...')
        error_list.append('错误过多（大于 {} 个），是否文件错误？'.format(MAX_ERROR_NUM))
        logging.error('\n'.join(error_list))
        return False

    if critical_error:
        error_list.append('[ ERROR ] 请确保 data 文件及 query 文件中均无 latin 名缺失！')
        logging.info('\n'.join(error_list))
        return False

    # Check if number of lines of data file is correct
    error_list.append(THIN_BAR_NO_NEWLINE)
    error_list.append('[ Start ] 开始校验 data 文件的列数目是否正确 ...')
    if len(data_file_tuple_list[0]) != len(DATA_FILE_HEADER_TUPLE):
        error_list.append(
            '[ ERROR ] data 文件的列数目有误：当前为 {}，应该为 {}（{}）'.format(
                len(data_file_tuple_list[0]),
                len(QUERY_FILE_HEADER_TUPLE),
                QUERY_FILE_HEADER_TUPLE))
        logging.info('\n'.join(error_list))
        return False

    # Check if number of lines of query file is correct
    error_list.append(THIN_BAR_NO_NEWLINE)
    error_list.append('[ Start ] 开始校验 query 文件的列数目是否正确 ...')
    if len(query_file_tuple_list[0]) != len(QUERY_FILE_HEADER_TUPLE):
        error_list.append(
            '[ ERROR ] data 文件的列数目有误：当前为 {}，应该为 {}（{}) '.format(
                len(query_file_tuple_list[0]),
                len(DATA_FILE_HEADER_TUPLE),
                DATA_FILE_HEADER_TUPLE))
        logging.info('\n'.join(error_list))
        return False

    # Check if is there any missing cell in data file
    error_list.append(THIN_BAR_NO_NEWLINE)
    error_list.append('[ Start ] 检查 data 文件中是否有缺失的单元格 ...')
    for i, row in enumerate(data_file_tuple_list):
        for j, cell in enumerate(row):
            if not cell:
                error_list.append(
                    '    -> data 文件中存在缺失的单元格: [%s:  行: %s, 列: %s]' %
                    (data_file, i + 1, j + 1))

    if len(error_list) > MAX_ERROR_NUM:
        error_list = error_list[:MAX_ERROR_NUM]
        error_list.append('... ...')
        error_list.append('错误过多（大于 {} 个），是否文件错误？'.format(MAX_ERROR_NUM))
        logging.error('\n'.join(error_list))
        return False

    # Check if is there any missing cell in query file
    error_list.append(THIN_BAR_NO_NEWLINE)
    error_list.append('[ Start ] 检查 query 文件中是否有缺失的单元格 ...')
    for i, row in enumerate(query_file_tuple_list):
        for j, cell in enumerate(row):
            if not cell:
                error_list.append(
                    '    -> query 文件中存在缺失的单元格: [%s:  行: %s, 列: %s]' %
                    (query_file, i + 1, j + 1))

    if len(error_list) > MAX_ERROR_NUM:
        error_list = error_list[:MAX_ERROR_NUM]
        error_list.append('... ...')
        error_list.append('错误过多（大于 {} 个），是否文件错误？'.format(MAX_ERROR_NUM))
        logging.error('\n'.join(error_list))
        return False

    # Check if latin names in query file in data file
    error_list.append(THIN_BAR_NO_NEWLINE)
    error_list.append('[ Start ] 检查 query 文件中的 latin 名是否在 data 文件中也存在 ...')
    tmp_latin_name_set = set([])
    latin_names_set_from_data_file = set(latin_names_in_data_file)
    for i, latin_name in enumerate(latin_names_in_query_file):
        if latin_name not in latin_names_set_from_data_file:
            if latin_name not in tmp_latin_name_set:
                tmp_latin_name_set.add(latin_name)
                error_list.append(
                    '    -> query 文件（%s）中的 latin 名（%s）不在 data 文件中 [行 %s]' %
                    (query_file, latin_name, i + 1))
    error_list.append(THIN_BAR_NO_NEWLINE)

    if len(error_list) > MAX_ERROR_NUM:
        error_list = error_list[:MAX_ERROR_NUM]
        error_list.append('... ...')
        error_list.append('错误过多（大于 {} 个），是否文件错误？'.format(MAX_ERROR_NUM))
        logging.error('\n'.join(error_list))
        return False

    # # Check if Latin names in built-in Latin name list
    # try:
    #     with open(DEFAULT_LATIN_NAME_FILE, 'r') as f:
    #         default_latin_names_list_1 = [x.strip() for x in f.readlines()
    #                                       if x.strip()]
    #     with open(DEFAULT_LATIN_NAME_FILE_2, 'r') as f:
    #         default_latin_names_list_2 = [x.strip() for x in f.readlines()
    #                                       if x.strip()]
    #     default_latin_names = set(default_latin_names_list_1
    #                               + default_latin_names_list_2)
    # except IOError as e:
    #     logging.warning(
    #         'Pass latin name check because no built-in latin name file '
    #         'was found: %s. (%s)' % (DEFAULT_LATIN_NAME_FILE, e))
    # else:
    #     logging.info('[ Start ] Validating if latin names from data file in '
    #                  'built-in latin name list...')
    #     tmp_warning_set = set([])
    #     for i, latin_name in enumerate(latin_names_in_data_file):
    #         if latin_name not in default_latin_names:
    #             if latin_name not in tmp_warning_set:
    #                 tmp_warning_set.add(latin_name)
    #                 logging.warning(
    #                     '[ WARNING ] [%s:  Line %s]  %s  ' %
    #                     (data_file, i+1, latin_name))
    #     logging.info(THIN_BAR_NO_NEWLINE)
    #     logging.info('[ Start ] Validating if latin names from query file in'
    #                  ' built-in latin name list...')
    #     tmp_warning_set = set([])
    #     for i, latin_name in enumerate(latin_names_in_query_file):
    #         if latin_name not in default_latin_names:
    #             if latin_name not in tmp_warning_set:
    #                 tmp_warning_set.add(latin_name)
    #                 logging.warning(
    #                     '[ WARNING ] [%s:  Line %s]  %s  ' %
    #                     (query_file, i+1, latin_name))
    error_list.append(BAR)
    logging.info('\n'.join(error_list))
    return True


def arg_parse():
    """Parse arguments and return filenames."""
    parser = argparse.ArgumentParser()

    parser.add_argument('-i', '--input', dest='query_file',
                        default='query.xlsx', help="Query file, xlsx format")
    parser.add_argument('-d', '--data', dest='data_file', default='data.xlsx',
                        help="Data file, xlsx format")
    parser.add_argument('-o', '--output', dest='output_file',
                        default='output.xlsx', help="Output file, xlsx format")

    args = parser.parse_args()
    logging.info("Plant Speciem Info Input Program:%s" % BAR)
    if any([args.query_file == "query.xlsx", args.data_file == 'data.xlsx',
            args.output_file == 'output.xslx']):
        logging.warning("You are using one or more default name(s).")
        logging.warning("You can use other names by:")
        logging.warning("    -i [--input]   query_file")
        logging.warning("    -d [--data]    data_file")
        logging.warning("    -o [--output]  output_file")
        logging.warning("Type --help to see full help message.\n")

    logging.info("%s    [  Query file ]  %s" % (THIN_BAR, args.query_file))
    logging.info("    [   Date file ]  %s" % args.data_file)
    logging.info("    [ Output file ]  %s%s" % (args.output_file, THIN_BAR))

    if not os.path.isfile(args.query_file):
        logging.error(" *  Query file does not exist:  %s"
                      % args.query_file)
        logging.warning(" [ Possible Solution]")
        logging.warning("       1. Please use default name:  query.xlsx.")
        logging.warning("       2. Specify query file by [-i query_file].")

    if not os.path.isfile(args.data_file):
        logging.error(" *  Data file does not exist:  %s"
                      % args.data_file)
        logging.warning(" [ Possible Solution]")
        logging.warning("       1. Please use default name:  data.xlsx.")
        logging.warning("       2. Specify data file by [-i data_file].")

    return args


# GUI implementation starts from here


class TextEmit(object):
    """Redirect standard output to Tkinter widget."""

    def __init__(self, widget, tag='stdout'):
        self.widget = widget
        self.tag = tag

    def write(self, out_str):
        self.widget.insert('end', out_str, (self.tag,))
        self.widget.tag_configure('stderr', foreground='red',
                                  background='yellow')
        self.widget.see('end')


class TextHandler(logging.Handler):
    """This class allows you to log to a Tkinter Text or ScrolledText widget"""

    def __init__(self, text):
        # run the regular Handler __init__
        logging.Handler.__init__(self)
        # Store a reference to the Text it will log to
        self.text = text

    def emit(self, record):
        msg = self.format(record)

        def append():
            self.text.configure(state='normal')
            self.text.insert(tk.END, msg + '\n')
            self.text.configure(state='disabled')
            # Autoscroll to the bottom
            self.text.yview(tk.END)

        # This is necessary because we can't modify the Text from other threads
        self.text.after(0, append)


class Application(tk.Frame):
    """GUI for main program."""

    def __init__(self, master=None):
        tk.Frame.__init__(self, master)
        self.grid()
        self.master.title('Specimen Input Graphical User Interface')
        self.master.geometry('1100x600')
        self.set_style()
        self.create_widgets()
        self.configure_layout()
        self.bind_command()
        self.queue = Queue.Queue()
        self.data_file = ''
        self.query_file = ''

    def set_style(self):
        """Set style for widgets."""
        s = ttk.Style()

        s.configure('TButton', padding='10 5')
        s.configure('exe.TButton', foreground='red')
        s.configure('TCombobox', padding='7')
        s.configure('TLabel', padding='3 7')
        s.configure('log.TLabel', foreground='blue')
        s.configure('TEntry', padding='5 7')

    def create_widgets(self):
        """Create widgets in main window."""
        self.content = ttk.Frame(self.master, padding='8')

        # Left side
        self.query_file_label_value = tk.StringVar()
        self.query_file_label = ttk.Label(
            self.content,
            textvariable=self.query_file_label_value)
        self.query_file_label_value.set('Query File:')

        self.query_file_combobox = ttk.Combobox(
            self.content,
        )

        self.data_file_label_value = tk.StringVar()
        self.data_file_label = ttk.Label(
            self.content,
            textvariable=self.data_file_label_value)
        self.data_file_label_value.set('Data File:')

        self.data_file_combobox_value = tk.StringVar()
        self.data_file_combobox = ttk.Combobox(
            self.content,
        )

        self.query_content_area = st.ScrolledText(
            self.content)

        self.input_status_label_value = tk.StringVar()
        self.input_status_label = ttk.Label(
            self.content,
            textvariable=self.input_status_label_value,
            style='log.TLabel')

        # Right side
        self.execute_button = ttk.Button(
            self.content,
            text='Start Query',
            style='exe.TButton')

        self.out_file_label = ttk.Label(
            self.content,
            text='Output file:')

        self.out_file_entry = ttk.Entry(
            self.content, )
        self.out_file_entry.insert('0', 'specimen.xlsx')

        self.log_area = st.ScrolledText(
            self.content)

        self.log_label_value = tk.StringVar()
        self.log_label = ttk.Label(
            self.content,
            textvariable=self.log_label_value,
            style='log.TLabel')

    def configure_layout(self):
        """Configure layout of widgets."""
        # grid
        self.content.grid(row=0, column=0, sticky='wens')

        self.query_file_label.grid(row=0, column=0)
        self.query_file_combobox.grid(row=0, column=1, sticky='we')
        self.data_file_label.grid(row=0, column=2)
        self.data_file_combobox.grid(row=0, column=3, sticky='we')
        self.query_content_area.grid(
            row=1, column=0, columnspan=4, sticky='wens')
        self.input_status_label.grid(
            row=2, column=0, columnspan=4, sticky='w')

        self.execute_button.grid(row=0, column=4, sticky='w')
        self.out_file_label.grid(row=0, column=5, )
        self.out_file_entry.grid(row=0, column=6, sticky='we')
        self.log_area.grid(row=1, column=4, columnspan=4, sticky='wens')
        self.log_label.grid(
            row=2, column=4, columnspan=4, sticky='w')

        # rowconfigure and columnconfigure
        self.master.rowconfigure(0, weight=1)
        self.master.columnconfigure(0, weight=1)

        self.content.rowconfigure(0, weight=0)
        self.content.rowconfigure(1, weight=1)
        self.content.rowconfigure(2, weight=0)
        self.content.columnconfigure(0, weight=1)
        self.content.columnconfigure(1, weight=1)
        self.content.columnconfigure(2, weight=1)
        self.content.columnconfigure(3, weight=1)
        self.content.columnconfigure(4, weight=1)
        self.content.columnconfigure(5, weight=1)
        self.content.columnconfigure(6, weight=1)
        self.content.columnconfigure(7, weight=1)

    def bind_command(self):
        """Bind command to widgets."""
        text_handler = TextHandler(self.log_area)
        logger = logging.getLogger()
        logger.addHandler(text_handler)

        self.query_file_combobox['values'] = self._candidate_query_files
        self.query_file_combobox.bind(
            '<<ComboboxSelected>>',
            lambda event: self._choose_query_file())

        self.data_file_combobox['values'] = self._candidate_data_files
        self.data_file_combobox.bind(
            '<<ComboboxSelected>>',
            lambda event: self._choose_data_file())

        self.query_content_area.insert('end', HELP.encode('utf-8'))

        self.execute_button['command'] = self._do_query

    @property
    def _candidate_query_files(self):
        """Values for combobox: All xlsx files ended with .xlsx."""
        _xlsx_file_list = [_ for _ in os.listdir('.') if _.endswith('.xlsx')]
        return _xlsx_file_list

    @property
    def _candidate_data_files(self):
        """values for combobox: All xlsx files ended with .xlsx."""
        _xlsx_file_list = [_ for _ in os.listdir('.') if _.endswith('.xlsx')]
        return _xlsx_file_list

    def _choose_query_file(self):
        """Command for button: Choose query file from combobox."""
        if not Application.check_dependencies():
            return

        self.query_file = self.query_file_combobox.get()
        if not self.query_file or \
                not os.path.isfile(self.query_file):
            sys.stderr.write('query 文件无效：{}. 请重新选择\n'.format(self.query_file))
            self.log_label_value.set('Error:  No query content!')
        xlsx_matrix = XlsxFile(self.query_file).xlsx_matrix
        query_content = "{}预览前 50 行内容:{}\n\n".format(BAR, BAR)
        for i, each_tuple in enumerate(xlsx_matrix[:50]):
            try:
                line_content = ' | '.join([unicode(_) for _ in each_tuple])
                query_content += ("[%d]  %s\n%s\n"
                                  % (i + 1, line_content, LINE_SPLITER))
            except UnicodeEncodeError as e:
                query_content = "query 文件无效，无法预览\n{}".format(e)
                break
            except BaseException as e:
                query_content = "query 文件无效，无法预览\n{}".format(e)
                break
        self.query_content_area.delete("0.1", "end-1c")
        self.query_content_area.insert("end", query_content)
        self.input_status_label_value.set('query 文件已选择:  {}'.format(self.query_file))

    def _choose_data_file(self):
        """Command for button: choose data file from combobox."""
        if not Application.check_dependencies():
            return

        data_file_path = self.data_file_combobox.get()
        if not data_file_path or not os.path.isfile(data_file_path):
            sys.stderr.write('data 文件无效：{}. 请重新选择\n'.format(data_file_path))
            self.log_label_value.set('Error:  No data file content!')
        xlsx_matrix = XlsxFile(data_file_path).xlsx_matrix
        query_content = "{}预览前 50 行内容:{}\n\n".format(BAR, BAR)
        for i, each_tuple in enumerate(xlsx_matrix[:50]):
            try:
                line_content = ' | '.join([unicode(_) for _ in each_tuple])
                query_content += ("[%d]  %s\n%s\n"
                                  % (i + 1, line_content, LINE_SPLITER))
            except UnicodeEncodeError as e:
                query_content = "data 文件无效，无法预览！\n{}".format(e)
                break
            except BaseException as e:
                query_content = "data 文件无效，无法预览！\n{}".format(e)
                break
        self.query_content_area.delete("0.1", "end-1c")
        self.query_content_area.insert("end", query_content)
        self.input_status_label_value.set('data 文件已选择：{}'.format(data_file_path))
        self.data_file = data_file_path

    def _do_query(self):
        """Main GUI program and core function."""
        if not Application.check_dependencies():
            return

        self.log_label_value.set('任务开始 ...')

        out_xlsx_file = self.out_file_entry.get().strip()

        logging.info('{}植物标本信息处理软件{}'.format(BAR, BAR))

        if any([self.query_file == "query.xlsx",
                self.data_file == 'data.xlsx',
                out_xlsx_file == 'output.xslx']):
            logging.warning("您使用了一个或多个默认参数名称.")

        logging.info("%s    [  Query file  ]  %s" % (THIN_BAR, self.query_file))
        logging.info("    [   Data file  ]  %s" % self.data_file)
        logging.info("    [ xlsx Outfile ]  %s" % out_xlsx_file)

        if not self.query_file or not self.data_file or not out_xlsx_file:
            logging.error("缺少参数！")
            self.log_label_value.set('缺少参数')
            return

        ThreadedTask(self.data_file, self.query_file, out_xlsx_file, self.log_label_value, self.queue).start()
        self.master.after(1000, self.process_queue)

    def process_queue(self):
        try:
            log_info = self.queue.get(0)
            logging.info(log_info)
        except Queue.Empty:
            print("blank")
            self.master.after(1000, self.process_queue)

    @staticmethod
    def check_dependencies():
        dependency_error_message = '请安装依赖包：pip install'
        dependency_ok = True
        if not bs4:
            dependency_error_message += ' bs4'
            dependency_ok = False
        if not openpyxl:
            dependency_error_message += ' openpyxl'
            dependency_ok = False
        if not requests:
            dependency_error_message += ' requests'
            dependency_ok = False

        if not dependency_ok:
            logging.error(dependency_error_message)

        return dependency_ok


class ThreadedTask(threading.Thread):
    def __init__(self, data_file, query_file, output_file, log_label_widget, queue):
        threading.Thread.__init__(self)
        self.data_file = data_file
        self.query_file = query_file
        self.out_xlsx_file = output_file
        self.log_label_value = log_label_widget
        self.queue = queue

    def run(self):
        time_start = time.time()

        try:
            # Data validation
            self.log_label_value.set('开始进行数据校验 ...')
            ok = data_validation(data_file=self.data_file, query_file=self.query_file)
            if not ok:
                self.queue.put('数据校验失败')
                self.log_label_value.set("数据校验失败！")
                return
            self.log_label_value.set('开始进行预处理，请耐心等待 ... ')
            query = Query(self.query_file, self.data_file)

            self.log_label_value.set('开始进行多进程处理，请耐心等待 ... ')
            out_tuple_list, log_info = query.do_multi_query()

            write_to_xlsx_file(out_tuple_list, xlsx_outfile_name=self.out_xlsx_file)
            write_result = '已将 {} 条记录写入到输出文件中；{}！'.format(
                len(out_tuple_list), self.out_xlsx_file)
            self.log_label_value.set(write_result)
            log_info += '\n{}'.format(write_result)

            time_end = time.time()
            self.log_label_value.set('任务完成，共花费时间: %.2f 秒' % (time_end - time_start))
            self.queue.put(log_info)
        except Exception as e:
            logging.error(e.message)
            self.queue.put('ERROR')
            self.log_label_value.set("失败！")


def main():
    """Main function."""
    args = arg_parse()
    query_file, offline_data_file, output_file = (
        args.query_file,
        args.data_file,
        args.output_file)
    # Data validation before program run
    try:
        data_validation(offline_data_file, query_file)
    except Exception as e:
        logging.error('无法进行数据校验，跳过校验 ... （原因：%s）' % e)

    try:
        q = Query(query_file, offline_data_file)
        out_tuple_list, log_info = q.do_multi_query()
        write_to_xlsx_file(out_tuple_list, xlsx_outfile_name=output_file)
    except KeyboardInterrupt as e:
        raise


def gui_main():
    """Main program for GUI."""
    app = Application()
    app.mainloop()


if __name__ == '__main__':
    gui_main()
    # main()

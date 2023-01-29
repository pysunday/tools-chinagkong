#coding: utf-8
import re
import xlsxwriter
import json
import tqdm
from os import path
from io import BytesIO
from urllib.parse import urlencode
from sunday.core import Logger, getParser, Fetch, printTable, Auth, MultiThread, printTable, getException
from sunday.tools.chinagkong.params import CMDINFO
from bs4 import BeautifulSoup
from pydash import find, get, chunk

CkError = getException({
    10001: '类型错误请核查',
    10002: '无法获取数据页数与条数'
    })

logger = Logger(CMDINFO['description']).getLogger()

class Chinagkong():
    def __init__(self, *args, **kwargs):
        urlBase = 'http://www.chinagkong.com'
        self.urlBase = urlBase
        self.urls = {
                'list': urlBase + '/company/',
                }
        self.fetch = Fetch()
        self.typename = None
        self.getUrlByType = None
        self.isShowlist = False
        self.companys = []
        self.thread_num = None
        self._thread_num_page = 20
        self.pbar = None
        self.tableTitleList = [{
                'key': 'name',
                'title': '公司名称',
                }, {
                'key': 'contactor',
                'title': '联系人',
                }, {
                'key': 'phone',
                'title': '手机',
                }, {
                'key': 'mobile',
                'title': '电话',
                }, {
                'key': 'fax',
                'title': '传真',
                }, {
                'key': 'good',
                'title': '货品描述',
                }, {
                'key': 'introduce',
                'title': '公司描述',
                }, {
                'key': 'url',
                'title': '数据来源',
                }]

    def printList(self):
        printTable(['编号', '类型', '代码', '链接'])(self.getAllTypes())

    def getAllTypes(self):
        res = self.fetch.get(self.urls['list'])
        soup = BeautifulSoup(res.text, 'lxml')
        links = soup.select('.jd_con_ul dl dd a')
        datas = [[
            idx + 1,
            it.text,
            it.attrs.get('href').split('/').pop().replace('.html', ''),
            self.urlBase + it.attrs.get('href'),
            ] for idx, it in enumerate(links)]
        def getUrlByType(typename):
            for it in datas:
                if it[1] == typename:
                    return it[2], it[3]
            raise CkError(10001, other=typename)
        self.getUrlByType = getUrlByType
        return datas

    def getPageInfo(self, url):
        res = self.fetch.get(url)
        soup = BeautifulSoup(res.text, 'lxml')
        pageEle = soup.select_one('#lblpage')
        pageMat = re.match(r'.*共有(\d+?)条记录.*共(\d+?)页.*', pageEle.text.replace('\xa0', ''))
        if pageMat:
            return pageMat.groups()
        raise CkError(10002)

    def getPageCompany(self, url):
        res = self.fetch.get(url)
        soup = BeautifulSoup(res.text, 'lxml')
        lis = soup.select('.jd_con_ul.jdqy li')
        for li in lis:
            nameEle = li.select_one('.gy_list_info_title a')
            goodEle = li.select_one('.gy_list_info_zy.jdqyzy')
            self.companys.append({
                'name': nameEle.text.strip(),
                'good': goodEle.text.strip(),
                'url': self.urlBase + nameEle.attrs.get('href'),
                'code': nameEle.attrs.get('href').split('/').pop().replace('.html', ''),
                })

    def wrapper(self, func, items):
        for item in items: func(item)

    def getData(self):
        self.pbar = tqdm.tqdm(total=len(self.companys))
        if self.thread_num:
            multiData = [[item for item in self.companys[i::self.thread_num]] for i in range(self.thread_num)]
            MultiThread(multiData, lambda item, _: [self.wrapper, (self.getDataByCompany, item)]).start()
        else:
            self.wrapper(self.getDataByCompany, self.companys)
        self.pbar.close()

    def getDataByCompany(self, company):
        try:
            res = self.fetch.get(company.get('url'), timeout_time=10)
            soup = BeautifulSoup(res.text, 'lxml')
            introduceEle = soup.select_one('.gsjj')
            infoEle = soup.select_one('.gsmes')
            contactEle = infoEle.select('.mes-top')[1].select('.mes-list span')
            if len(contactEle) != 4: __import__('ipdb').set_trace()
            company.update({
                'contactor': get(contactEle, '0.text', '').strip(),
                'mobile': get(contactEle, '1.text', '').strip(),
                'phone': get(contactEle, '2.text', '').strip(),
                'fax': get(contactEle, '3.text', '').strip(),
                'introduce': get(introduceEle, 'text', '').strip(),
                'success': True,
                })
        except Exception as e:
            logger.exception(e)
            logger.error(f'获取数据失败：{company}')
            company.update({
                'success': False
                })
        self.pbar.update(1)

    def getDataByPage(self, typename=None):
        url = self.urls['list']
        if typename: code, url = self.getUrlByType(typename)
        (count, pages) = self.getPageInfo(url)
        urls = []
        for idx in range(int(pages)):
            page = str(idx + 1)
            flag = f'yp_vlist_{page}' if typename else 'index'
            urls.append(f'{self.urlBase}/company/{flag}_{page}.html')
        multiData = [[item for item in urls[i::self._thread_num_page]] for i in range(self._thread_num_page)]
        MultiThread(multiData, lambda item, _: [self.wrapper, (self.getPageCompany, item)]).start()
        self.getData()

    def saveExcel(self, filename='工控信息网'):
        companys = [it for it in self.companys if it.get('success') == True]
        filepath = path.abspath(f'./{filename}.xlsx')
        print(f'全部数据{len(self.companys)}条，成功抓取{len(companys)}条, 保存文件：{filepath}')
        workbook = xlsxwriter.Workbook(filepath)
        bold = workbook.add_format({'bold': True})
        cell_format = workbook.add_format()
        cell_format.set_text_wrap()
        cell_format.set_align('center')
        cell_format.set_align('vcenter')
        worksheet = workbook.add_worksheet()
        worksheet.set_default_row(80)
        worksheet.set_row(0, 30)
        worksheet.set_column('A:A', 35)
        worksheet.set_column('B:E', 15)
        worksheet.set_column('F:F', 35)
        worksheet.set_column('G:G', 50)
        worksheet.set_column('H:H', 35)
        for idx, item in enumerate(self.tableTitleList):
            worksheet.write(0, idx, item.get('title'), cell_format)
            for didx, data in enumerate(companys):
                worksheet.write(didx + 1, idx, data.get(item.get('key')), cell_format)
        workbook.close()

    def run(self):
        self.getAllTypes()
        if self.isShowlist:
            self.printList()
        elif self.typename:
            self.getDataByPage(self.typename)
            self.saveExcel(self.typename)
        else:
            self.getDataByPage()
            self.saveExcel()


def runcmd():
    parser = getParser(**CMDINFO)
    handle = parser.parse_args(namespace=Chinagkong())
    handle.run()


if __name__ == "__main__":
    runcmd()

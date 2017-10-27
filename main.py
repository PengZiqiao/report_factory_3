import numpy as np
import pandas as pd
from winsun.func import percent
from winsunDB.query import ZHUZHAI, SHANGYE, BANGONG, BIESHU
from utils import Month

from office import Excel, PPT
from query import ReporterQuery

path = r'E:\dell\Documents\Python_Projects\report_factory_3\templates'


def update_data():
    """从数据库查询数据，写入excel文件
    :return shuoli 说理字典
    """

    def rank_plate(usg, df, in_):
        col = {'Sale': '上市面积(万㎡)', 'Sold': '成交面积(万㎡)'}[in_]
        df_plate = df.sort_values(col, ascending=False)[[col]][:3]
        df_plate.columns = ['plate_space']
        df_rank = pd.DataFrame()
        for plate in df_plate.index:
            plate = '仙西' if plate == '仙林' else plate
            df_ = q.rank(usg, plate, in_, num=3)[['板块', 'space']]
            df_rank = df_rank.append(df_)

        # 调整
        df_plate['板块'] = df_plate.index
        df_rank['space'] = round(df_rank['space'] / 1e4, 2)
        df_rank['name'] = df_rank.index
        df_rank = pd.merge(df_rank, df_plate, on='板块')
        df_rank = df_rank[['板块', 'plate_space', 'name', 'space']]
        df_rank = df_rank.set_index('板块')
        # 写入excel
        sheet = f'{key}rank_splate' if in_ == 'Sale' else f'{key}rank_plate'
        print(sheet)
        excel.df2sheet(df_rank, sheet)

    excel = Excel()
    q = ReporterQuery()
    shuoli = dict()
    # 遍历4种物业类型
    usage = {'住宅': ZHUZHAI, '办公': BANGONG, '商业': SHANGYE, '别墅': BIESHU}
    for key, usg in usage.items():
        # 遍历两张表
        for item in ['range', 'plate']:
            if item == 'range':
                period = 21 if key == '住宅' else 13
            else:
                period = 1
            df = q.gxj_(usg, period, item)
            sheet = f'{key}{item}'
            print(sheet)
            # 写入excel文件
            excel.df2sheet(df, sheet)

            # 说理
            if item == 'range':
                sl = ShuoLi(df)
                shuoli[f'{key}'] = sl.all()
            # 板块排行
            if item == 'plate':
                rank_plate(usg, df, 'Sold')
                rank_plate(usg, df, 'Sale')
        # 排行
        for by in ['space', 'money']:
            num = 10 if key == '住宅' else 5
            df = q.rank(usg, by=by, num=num)
            sheet = f'{key}rank_{by}'
            excel.df2sheet(df, sheet)

    excel.save()
    return shuoli


class ShuoLi():
    """
    df 传入一个供销走势的DataFrame
    ss, cj, jj 分别代表上市、成交、均价
    ss_h, cj_h, jj_h 对应环比
    ss_t, cj_t, jj_t 对应同比
    """

    def __init__(self, df, degree=0):
        # 计算出环、同比
        df_h = df.pct_change()
        df_t = df.pct_change(12)
        # 调整
        self.df = pd.concat([df.iloc[-1:], df_h.iloc[-1:], df_t.iloc[-1:]])
        self.df.index = ['v', 'h', 't']
        self.df.columns = ['ss', 'cj', 'jj']

    def thb_text(self, h, t):
        """同环比文字"""
        hb = '' if np.isnan(h) else f'，环比{percent(h,0)}'
        tb = '' if np.isnan(t) else f'，同比{percent(t,0)}'
        return f'{hb}{tb}'

    def text(self, by):
        v = self.df.at['v', by]
        h = self.df.at['h', by]
        t = self.df.at['t', by]

        v_text = {
            'ss': ('本月无上市。', f'本月上市{v:.2f}万㎡'),
            'cj': ('本月无成交。', f'本月成交{v:.2f}万㎡'),
            'jj': ('', f'成交均价{v:.0f}元/㎡')
        }

        if v == 0 or np.isnan(v):
            return v_text[by][0]
        else:
            return f'{v_text[by][1]}{self.thb_text(h, t)}。'

    def all(self):
        ss = self.text('ss')
        cj = self.text('cj')
        jj = self.text('jj')
        return f'{ss}\r{cj}\r{jj}'


class Report(PPT):
    def __init__(self, shuoli):
        super(Report, self).__init__()
        date = Month().date()
        self.month = date.month
        self.year = date.year
        self.date_text = f'{self.year}年{self.month}月'
        self.shuoli = shuoli

    def range(self, wuye, shuoli, page_idx):
        print(f'>>> 生成{wuye}月度量价页...')
        # 说理
        self.text(shuoli[wuye], page_idx, 2)

    def plate(self, wuye, page_idx):
        print(f'>>> 生成{wuye}板块表现页...')
        i = int(wuye == '住宅')
        # 量价标题
        txt = f'{self.date_text}南京{wuye}市场供销价'
        self.text(txt, page_idx, 0 + i)
        # 左下标题
        txt = f'{self.date_text}上市量（前三板块的前三项目）'
        self.text(txt, page_idx, 3 + i)
        # 右下标题
        txt = f'{self.date_text}成交量（前三板块的前三项目）'
        self.text(txt, page_idx, 4 + i)

    def rank(self, page_idx):
        # 左标题
        txt = f'{self.date_text}成交面积排行榜'
        self.text(txt, page_idx, 3)
        # 右标题
        txt = f'{self.date_text}成交金额排行榜'
        self.text(txt, page_idx, 4)

    def three_pages(self, usg, page_idx):
        self.range(usg, self.shuoli, page_idx)
        self.plate(usg, page_idx + 1)
        self.rank(page_idx + 2 + int(usg == '住宅'))


if __name__ == '__main__':
    shuoli = update_data()
    rpt = Report(shuoli)

    rpt.three_pages('住宅', 0)
    rpt.three_pages('办公', 4)
    rpt.three_pages('商业', 7)
    rpt.three_pages('别墅', 10)
    rpt.save(f'e:/月报测试/{rpt.date_text}月报.pptx')

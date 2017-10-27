from winsunDB.query import Query, ZHUZHAI
from utils import Month
from winsunDB.model import MonthSale, MonthSold
from sqlalchemy.sql import func, label
import pandas as pd


class ReporterQuery(Query):
    m = Month()

    def month_index(self, period):
        dates = list(self.m.before(i) for i in range(period))
        dates.reverse()
        return list(f'{each[0]%100}{each[1]:02d}' for each in dates)

    def gxj_(self, usage, period, output, plate=None):
        end = self.m.date()
        start = self.m.date_before(period - 1)
        df = self.gxj(by='Month', date_range=(start, end), usage=usage, output_by=output, plate=plate)

        df.columns = ['上市面积(万㎡)', '成交面积(万㎡)', '成交均价(元/㎡)']
        if output == 'range':
            df.index = self.month_index(period)

        return df

    def rank(self, usage=ZHUZHAI, plate=None, in_='Sold', by='space', num=10):
        """ 排行榜

        :param usage: 功能，填ZHUZHAI\BANGONG\SHANGYE\BIESHU...
        :param plate: 板块
        :param in_: Sale上市\Sold成交
        :param by: space面积\set套数\money金额
        :param num: 前n名， 默认
        :return: df
        """
        # 查询
        table = eval(f'Month{in_}')
        keys = [
            table.pop_name,
            table.plate,
        ]
        res = self.query(
            *keys,
            func.sum(table.space).label('space'),
            func.sum(table.set).label('set'),
            func.sum(table.money).label('money')
        ) \
            .group_by(
            *keys
        ) \
            .filter(
            table.date == self.m.date(),
            table.usage.in_(usage)
        )
        if plate:
            res = res.filter(table.plate == plate)

        # dateframe处理
        df = self.dateframe(res, 'popularizename')
        df = df.sort_values(by, ascending=False).iloc[:num]
        df['price'] = round(df['money'] / df['space'], 0).astype('int')
        df['money'] = round(df['money'] / 1e4, 2)
        return df


if __name__ == '__main__':
    rq = ReporterQuery()


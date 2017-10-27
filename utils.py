from datetime import date


class Month:
    def __init__(self):
        self.month = date.today().month
        self.year = date.today().year

    def before(self, i):
        """i个月之前"""
        if self.month - i > 0:
            return self.year, self.month - i
        elif self.month - (i - 12) > 0:
            return self.year - 1, self.month - (i - 12)
        else:
            return self.year - 2, self.month - (i - 24)

    def date(self):
        past = True if 1 <= date.today().day < 25 else False
        return date(*self.before(1), 1) if past else date(self.year, self.month, 1)

    def date_before(self, i):
        return date(*self.before(i), 1)

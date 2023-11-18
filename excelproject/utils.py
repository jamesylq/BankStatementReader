from typing import *

SPACE = 11.3375
SHEETTOP = [
    ("Index", 5),
    ("Transaction Date", 14),
    ("Value Date", 14),
    ("Description", 90),
    ("Withdrawal", 10),
    ("Deposit", 10),
    ("Balance", 10)
]
OCBCIGNORE = ["BALANCE B/F", "BALANCE C/F"]
DBSIGNORE = ["Balance Brought Forward", "Balance Carried Forward"]


def find(l, y):
    for x in l:
        if x.y == y:
            return x


def rws(s: str, char: Iterable = None):
    if char is None:
        char = '\n\t ,'
    for c in char:
        s = s.replace(c, '')
    return s


class Entry:
    def __init__(self, y):
        self.transDate = None
        self.valueDate = None
        self.description = None
        self.withdrawal = None
        self.deposit = None
        self.balance = None
        self.y = y
        self.l = None

    def countNone(self):
        c = 0
        for x in [self.transDate, self.valueDate, self.description, self.balance, self.l]:
            if x is None:
                c += 1
        return c

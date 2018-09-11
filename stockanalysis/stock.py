from xlrd import xldate_as_tuple
import xlrd
import xlwt
from datetime import date,datetime

class Record(object):
    def __init__(self, time, price, count, earning):
        self.time = time
        self.price = price
        self.count = count
        self.earning = earning

    def __str__(self):
        return "Record: %s price: %.2f count: %d earnings: %.2f " % (self.time, self.price, self.count, self.earning)

class GridStock(object):
    Rate = 0.9
    tradingunit = 0.1
    def __init__(self, allMoney, openPrice, gridrate, dealtimes, maxPrices, minPrices, closePrices):
        self.gridrate = gridrate
        self.dealtimes = dealtimes
        self.maxPrices = maxPrices
        self.minPrices = minPrices
        self.closePrices = closePrices
        self.allMoney = allMoney
        self.currMoney = allMoney
        self.currTime = ""
        self.records = []
        self.jetton = []
        self.openPrice = openPrice
        self.earnings = 0.0
        # save result into excel
        self.savefd = xlwt.Workbook()
        self.recordsfd = None
        self.selloutfd = None
        self.currRecordRow = 0
        self.currSelloutRow = 0

    def init_save_file(self):
        srow0 = [u'日期', u'当前资金', u'收益', u'收益率']
        self.selloutfd = self.savefd.add_sheet(u'空仓记录',cell_overwrite_ok=True)
        for i in range(0, len(srow0)):
            self.selloutfd.write(0, i, srow0[i])
        self.currSelloutRow += 1

        self.recordsfd = self.savefd.add_sheet(u'收益汇总',cell_overwrite_ok=True)
        row0 = [u'日期', u'价格', u'交易数量', u'收益', u'收益率']
        for i in range(0, len(row0)):
            self.recordsfd.write(0, i, row0[i])
        self.currRecordRow += 1

    def save_record(self, item):
        # self.time, self.price, self.count, self.earning
        self.recordsfd.write(self.currRecordRow, 0, item.time)
        self.recordsfd.write(self.currRecordRow, 1, item.price)
        self.recordsfd.write(self.currRecordRow, 2, item.count)
        self.recordsfd.write(self.currRecordRow, 3, item.earning)
        self.recordsfd.write(self.currRecordRow, 4, (item.earning / self.allMoney))
        self.currRecordRow += 1
    
    def save_sellout_record(self):
        self.selloutfd.write(self.currSelloutRow, 0, self.currTime)
        self.selloutfd.write(self.currSelloutRow, 1, self.currMoney)
        self.selloutfd.write(self.currSelloutRow, 2, self.earnings)
        self.selloutfd.write(self.currSelloutRow, 3, (self.earnings / self.allMoney))
        self.currSelloutRow += 1

    def show(self):
        for rec in self.records:
            self.save_record(rec)

    def buy(self, price):
        if price == 0:
            return
        lowmoney = GridStock.tradingunit * self.allMoney
        workmoney = lowmoney
        if self.currMoney < lowmoney:
            return

        count = int(lowmoney / price)
        count -= (count % 100)
        if count == 0:
            '''china must buy int count'''
            return
        
        self.currMoney -= (count * price)
        rec = Record(self.currTime, price, count, self.earnings)
        self.records.append(rec)
        self.jetton.append(rec)

    def sell(self, price):
        dep = len(self.jetton)
        if price == 0 or dep == 0:
            return
        trec = self.jetton.pop()
        count = -1 * trec.count
        self.currMoney += trec.count * price
        tearnings = trec.count * (price - trec.price)
        self.earnings += tearnings
        rec = Record(self.currTime, price, count, self.earnings)
        self.records.append(rec)
            
    def caculate(self):
        self.init_save_file()
        fsellout = False
        i = 0
        price = self.openPrice
        while i < len(self.closePrices):
            if len(self.jetton) == 0 and fsellout is False:
                self.save_sellout_record()
                fsellout = True
            self.currTime = self.dealtimes[i]
            work1 = (self.maxPrices[i] - self.closePrices[i]) * GridStock.Rate + self.closePrices[i]
            work2 = (self.closePrices[i] - self.minPrices[i]) * (1 - GridStock.Rate) + self.minPrices[i]
            if work1 < work2:
                work1 = self.maxPrices[i]
                work2 = self.minPrices[i]
            i += 1
            maxThreshold = price * (self.gridrate + 1)
            if  work1 >= maxThreshold:
                self.sell(work1)
                if len(self.jetton) == 0:
                    price = work1
                else:
                    price = self.jetton[-1].price
                continue
            minThreshold = price * (1 - self.gridrate)
            if work2 <= minThreshold:
                fsellout = False
                self.buy(work2)
                price = work2
                continue
        
        self.show()
        # save into excel
        nowTime=datetime.now().strftime('%H%M%S')
        fname = r"E:\stockdatas\result_%s.xls" % nowTime
        self.savefd.save(fname)


def read_excel(fname):
    workbook = xlrd.open_workbook(fname)
    print(workbook.sheet_names())

    sheet2 = workbook.sheet_by_index(0)

    print(sheet2.name,sheet2.nrows,sheet2.ncols)

    fdealTimes = sheet2.col_values(5)[1:]
    strhighestPrices = sheet2.col_values(9)[1:]
    strlowestPrices = sheet2.col_values(10)[1:]
    strclosePrices = sheet2.col_values(11)[1:]
    
    dealTimes = []
    highestPrices = []
    lowestPrices = []
    closePrices = []
    i = 0
    while i < len(strclosePrices):
        date = datetime(*xldate_as_tuple(fdealTimes[i], 0))
        cell = date.strftime('%Y/%m/%d')
        dealTimes.append(cell)
        highestPrices.append(float(strhighestPrices[i]))
        lowestPrices.append(float(strlowestPrices[i]))
        closePrices.append(float(strclosePrices[i]))
        i += 1

    openPrice = sheet2.cell(1,8).value

    grid = GridStock(100000, openPrice, 0.1, dealTimes, highestPrices, lowestPrices, closePrices)
    grid.caculate()

  
if __name__ == '__main__':
  read_excel(r"E:\stockdatas\jiansheprice.xlsx")
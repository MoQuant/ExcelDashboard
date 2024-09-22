import xlwings as xw
import websocket
import json
import threading
import numpy as np
import time

'''
sheet = xw.Book("Sheet.xlsm").sheet[0]
sheet.range("").value
'''

class OrderBook(threading.Thread):
    
    def __init__(self):
        threading.Thread.__init__(self)
        self.url = 'wss://ws-feed.exchange.coinbase.com'
        self.bids = {}
        self.asks = {}

    def run(self):
        conn = websocket.create_connection(self.url)
        msg = {'type':'subscribe','product_ids':['BTC-USD'],'channels':['level2_batch']}
        conn.send(json.dumps(msg))
        while True:
            resp = json.loads(conn.recv())
            if 'type' in resp.keys():
                if resp['type'] == 'snapshot':
                    self.bids = {float(price):float(volume) for price, volume in resp['bids']}
                    self.asks = {float(price):float(volume) for price, volume in resp['asks']}
                if resp['type'] == 'l2update':
                    for (side, price, volume) in resp['changes']:
                        price, volume = float(price), float(volume)
                        if side == 'buy':
                            if volume == 0:
                                if price in self.bids.keys():
                                    del self.bids[price]
                            else:
                                self.bids[price] = volume
                        else:
                            if volume == 0:
                                if price in self.asks.keys():
                                    del self.asks[price]
                            else:
                                self.asks[price] = volume

def ExcelBids(bids, depth=25):
    bids = list(sorted(bids.items(), reverse=True))[:depth]
    bids = np.array(bids)
    bids[:, 1] = np.cumsum(bids[:, 1])
    return bids.tolist()

def ExcelAsks(bids, depth=25):
    bids = list(sorted(bids.items()))[:depth]
    bids = np.array(bids)
    bids[:, 1] = np.cumsum(bids[:, 1])
    return bids.tolist()

book = OrderBook()
sheet = xw.Book("CS.xlsx").sheets[0]

book.start()

while True:
    if len(book.bids) > 0:
        print(len(book.bids), len(book.asks))
        exbids = ExcelBids(book.bids)
        exasks = ExcelAsks(book.asks)
        sheet.range("B5:C30").value = exbids
        sheet.range("E5:F30").value = exasks



book.join()

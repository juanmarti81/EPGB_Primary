import time
import pandas as pd
import pyRofex
import xlwings as xw
import numpy as np
import config

'''-----------
Excel sheets
--------------'''
wb = xw.Book('EPGB V3.3.1 FE_Primary.xlsb')
shtTickers = wb.sheets('Tickers')
shtData = wb.sheets('Primary')
shtOperaciones = wb.sheets('Posiciones')
shtOpciones = wb.sheets('Opciones')

'''-----------
Connect to rofex
--------------'''
RFX_Environment = pyRofex.Environment.LIVE
pyRofex.initialize(user=config.COMITENTE, password=config.PASSWORD,
                   account=config.COMITENTE, environment=RFX_Environment)

'''-----------
Get all Tickers
--------------'''
rng = shtTickers.range('A2:C500').expand()
tickers = pd.DataFrame(rng.value, columns=['ticker', 'symbol', 'strike'])
opciones = tickers.copy()
opciones = opciones.set_index('ticker')
opciones['bidSize'] = 0
opciones['bid'] = 0
opciones['ask'] = 0
opciones['asksize'] = 0
calls = pd.DataFrame()
puts = pd.DataFrame()

rng = shtTickers.range('E2:F500').expand()
tickers = tickers.append(pd.DataFrame(rng.value, columns=['ticker', 'symbol']))
rng = shtTickers.range('H2:I500').expand()
tickers = tickers.append(pd.DataFrame(rng.value, columns=['ticker', 'symbol']))
rng = shtTickers.range('K2:L500').expand()
tickers = tickers.append(pd.DataFrame(rng.value, columns=['ticker', 'symbol']))
rng = shtTickers.range('N2:O500').expand()
tickers = tickers.append(pd.DataFrame(rng.value, columns=['ticker', 'symbol']))
rng = shtTickers.range('Q2:R500').expand()
tickers = tickers.append(pd.DataFrame(rng.value, columns=['ticker', 'symbol']))

# Get the list of active tickers from Primary and filter our list
instruments_2 = pyRofex.get_detailed_instruments()
data = pd.DataFrame(instruments_2['instruments'])
df = pd.DataFrame.from_dict(dict(data['instrumentId']), orient='index')
df = df['symbol'].to_list()
tickers['remove'] = tickers['ticker'].isin(df).astype(int)
tickers = tickers[tickers['remove'] !=0]
instruments = tickers['ticker'].to_list()


'''-----------
Set global dataframes
--------------'''
df_datos = pd.DataFrame({'ticker': tickers['ticker'].to_list(),'symbol': tickers['symbol'].to_list()}, columns=[
    'ticker', 'symbol', 'bidsize', 'bid', 'ask', 'asksize', 'last', 'close','open', 'high', 'low', 'volume','lastupdate'])
df_datos = df_datos.set_index('ticker')

thisData = pd.DataFrame(columns=['ticker','symbol', 'bidsize', 'bid', 'ask', 'asksize', 'last', 'close','open', 'high', 'low', 'volume', 'lastupdate'])

def addToOptions(symbol, bidSize, bid, ask, askSize):
    global opciones
    thisData = pd.DataFrame([{'ticker': symbol, 'bidsize': bidSize, 'bid': bid, 'ask': ask, 'asksize': askSize}])
    thisData = thisData.set_index('ticker')
    opciones.update(thisData)
    calcular_opciones()

def calcular_opciones():
    global opciones, calls, puts, df_datos
    calls =  opciones.copy()
    calls = calls.filter(like="GFGC", axis=0)
    calls.sort_values(by=['strike'], ascending=True)
    calls['strike_1'] = calls['strike'].shift(-1)
    calls['bid_1'] = calls['bid'].shift(-1)
    calls['ratio_1'] = np.where(calls['ask'] == 0, 0, np.where(calls['bid_1'] == 0, 0,((calls['ask'] - calls['bid_1']) / (calls['strike_1'] - calls['strike']) )))
    calls['strike_dif'] = calls['strike_1'] - calls['strike']
    calls['dif_1'] = np.where(calls['ask'] == 0, 0, np.where(calls['bid_1'] == 0, 0, calls['ask'] - calls['bid_1']))
    calls['strike_2'] = calls['strike'].shift(-2)
    calls['bid_2'] = calls['bid'].shift(-2)
    calls['ratio_2'] = np.where(calls['ask'] == 0, 0, np.where(calls['bid_2'] == 0, 0,((calls['ask'] - calls['bid_2']) / (calls['strike_2'] - calls['strike']) )))
    calls['strike_dif_2'] = calls['strike_2'] - calls['strike']
    calls['dif_2'] = np.where(calls['ask'] == 0, 0, np.where(calls['bid_2'] == 0, 0, calls['ask'] - calls['bid_2']))
    calls.rename(columns={"strike_1": "Base + 1", "strike_2": "Base + 2","bid_1": "Bid + 1","bid_2": "Bid + 2", "ratio_1": "Ratio + 1", "ratio_2": "Ratio + 2", "dif_1": "Dif + 1", "dif_2": "Dif + 2"})

    puts = opciones.copy()
    puts = puts.filter(like="GFGV", axis=0)
    puts.sort_values(by=['strike'], ascending=True)
    puts['strike_1'] = puts['strike'].shift(1)
    puts['bid_1'] = puts['bid'].shift(1)
    puts['ratio_1'] = np.where(puts['ask'] == 0, 0, np.where(puts['bid_1'] == 0, 0, (
                (puts['ask'] - puts['bid_1']) / (puts['strike'] - puts['strike_1']))))
    puts['strike_dif'] = puts['strike'] - puts['strike_1']
    puts['dif_1'] = np.where(puts['ask'] == 0, 0, np.where(puts['bid_1'] == 0, 0, puts['ask'] - puts['bid_1']))
    puts['strike_2'] = puts['strike'].shift(2)
    puts['bid_2'] = puts['bid'].shift(2)
    puts['ratio_2'] = np.where(puts['ask'] == 0, 0, np.where(puts['bid_2'] == 0, 0, (
        (puts['ask'] - puts['bid_2']) / (puts['strike'] - puts['strike_2']))))
    puts['strike_dif_2'] = puts['strike'] - puts['strike_2']
    puts['dif_2'] = np.where(puts['ask'] == 0, 0, np.where(puts['bid_2'] == 0, 0, puts['ask'] - puts['bid_2']))
    puts.rename(columns={"strike_1": "Base + 1", "strike_2": "Base + 2", "bid_1": "Bid + 1", "bid_2": "Bid + 2",
                          "ratio_1": "Ratio + 1", "ratio_2": "Ratio + 2", "dif_1": "Dif + 1", "dif_2": "Dif + 2"})

def addTick(symbol, bidSize, bid, ask, askSize, last, close, open, high, low, volume, lastUpdate):
    global thisData, bonos, opciones
    thisData = pd.DataFrame([{'ticker': symbol, 'bidsize': bidSize, 'bid': bid, 'ask': ask, 'asksize': askSize, 'last': last, 'close':close, 'open': open, 'high': high, 'low': low, 'volume': volume, 'lastupdate': time.strftime('%m/%d/%Y %H:%M:%S', time.gmtime(lastUpdate / 1000.))}])
    thisData = thisData.set_index('ticker')
    df_datos.update(thisData)
    if len(opciones.filter(like = symbol, axis=0)) >0:
        addToOptions(symbol, bidSize, bid, ask, askSize)

def market_data_handler(message):
    symbol = message['instrumentId']['symbol']
    lastUpdate = None if not message['marketData']['LA'] else message['marketData']['LA']['date']
    last = None if not message['marketData']['LA'] else message['marketData']['LA']['price']
    bid = None if not message['marketData']['BI'] else message['marketData']['BI'][0]['price']
    bidSize = None if not message['marketData']['BI'] else message['marketData']['BI'][0]['size']
    ask = None if not message['marketData']['OF'] else message['marketData']['OF'][0]['price']
    askSize = None if not message['marketData']['OF'] else message['marketData']['OF'][0]['size']
    close = None if not message['marketData']['CL'] else message['marketData']['CL']['price']
    open = None if not message['marketData']['OP'] else message['marketData']['OP']
    high = None if not message['marketData']['HI'] else message['marketData']['HI']
    low = None if not message['marketData']['LO'] else message['marketData']['LO']
    volume = None if not message['marketData']['LO'] else message['marketData']['EV']
    print(symbol, bidSize, bid, ask, askSize, last, close, open, high, low, volume, lastUpdate)
    addTick(symbol, bidSize, bid, ask, askSize, last, close, open, high, low, volume, lastUpdate)

def error_handler(message):
    print("Error Message Received: {0}".format(message))

def exception_handler(e):
    print("Exception Occurred: {0}".format(e.message))

'''---------------------------
Order Report Subscription 
----------------------------'''
df_order = pd.DataFrame()
def order_report_handler(message):
    pass
def order_error_handler(message):
    print("Error Message Received: {0}".format(message))

def order_exception_handler(e):
    print("Exception Occurred: {0}".format(e.message))

pyRofex.init_websocket_connection(market_data_handler=market_data_handler,
                                  error_handler=error_handler,
                                  exception_handler=exception_handler,
                                  order_report_handler=order_report_handler,)

entries = [pyRofex.MarketDataEntry.BIDS,
           pyRofex.MarketDataEntry.OFFERS,
           pyRofex.MarketDataEntry.LAST,
           pyRofex.MarketDataEntry.OPENING_PRICE,
           pyRofex.MarketDataEntry.CLOSING_PRICE,
           pyRofex.MarketDataEntry.HIGH_PRICE,
           pyRofex.MarketDataEntry.LOW_PRICE,
           pyRofex.MarketDataEntry.TRADE_VOLUME,
           pyRofex.MarketDataEntry.NOMINAL_VOLUME,
           pyRofex.MarketDataEntry.TRADE_EFFECTIVE_VOLUME]

pyRofex.market_data_subscription(tickers=instruments, entries=entries, depth=5)
# pyRofex.order_report_subscription()

shtOperaciones.range('C4:K500').value = ""
shtData.range('A1:L1000').value = ""

while True:
    try:
        # loop = asyncio.get_event_loop()
        # update = updateSheet()
        # loop.run_until_complete(update)
        shtData.range('A1').options(index=False, headers=True).value = df_datos
        shtOpciones.range('C3').options(index=False, headers=True).value = calls
        shtOpciones.range('X3').options(index=False, headers=True).value = puts
        time.sleep(1)
    except AssertionError as error:
        print(error)
    except Exception as inst:
        print(type(inst))  # the exception instance
        print(inst.args)  # arguments stored in .args
        print(inst)  # __str__ allows args to be printed directly,
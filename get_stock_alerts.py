import time,sys
from yahoo_fin import stock_info as si
from numpy import around
from win10toast import ToastNotifier
import pprint as pp

import argparse

parser = argparse.ArgumentParser()
parser.add_argument(
    '--symbol',
    '-s',
    default='YESBANK.NS',
    help='Symbol name from yahoo finance website Ex: for Yes Bank, YESBANK.NS')

parser.add_argument(
    '--percent',
    '-p',
    default='0.01',
    help='Change in percent, on which price has to be notified')

parser.add_argument(
    '--information',
    '-i',
    default='',
    help='Information about the stock. Ex: data, quote, stats, holders, analysis, income, balance, cashflow')

parser.add_argument(
    '--repeat',
    '-r',
    default='yes',
    help='Get stock prices repeatedly?')

parser.add_argument(
    '--price',
    '-pr',
    default='0',
    help='Bought/Sold price')

parser.add_argument(
    '--quantity',
    '-q',
    default='1',
    help='Quantity of shares bought/sold')

parser.add_argument(
    '--order_type',
    '-o',
    default='buy',
    help='Order type. buy or sell')

input_arguments = parser.parse_args()



def ShowNotification(title, description, delay =2):
    toaster = ToastNotifier()
    toaster.show_toast(title,
               description,
               icon_path=None,
               duration=delay,
               threaded=True)
    while toaster.notification_active(): time.sleep(0.1)
    


if input_arguments.information != '':
    if input_arguments.information == 'data':
        print(si.get_data(input_arguments.symbol))
        
    elif input_arguments.information == 'quote':
        pp.pprint(si.get_quote_table(input_arguments.symbol))
        
    elif input_arguments.information == 'stats':
        pp.pprint(si.get_stats(input_arguments.symbol))

    elif input_arguments.information == 'holders':
        pp.pprint(si.get_holders(input_arguments.symbol))
        
    elif input_arguments.information == 'analysis':
        pp.pprint(si.get_analysts_info(input_arguments.symbol))

    elif input_arguments.information == 'income':
        pp.pprint(si.get_income_statement(input_arguments.symbol))

    elif input_arguments.information == 'balance':
        pp.pprint(si.get_balance_sheet(input_arguments.symbol))
        
    elif input_arguments.information == 'cashflow':
        pp.pprint(si.get_cash_flow(input_arguments.symbol))
        
    sys.exit()
    
    
    
previous_price = 0
profit = 0
current_price = 0

while True:  
    
    while (True):
        try:
            current_price  = around(si.get_live_price(input_arguments.symbol),2) 
            break
        except:
            print('Unable to get share price from yahoo. Retrying !!!')
        
    
    if input_arguments.order_type.lower() == 'buy':
        profit =   around((float(current_price) - float(input_arguments.price))*float(input_arguments.quantity),2)
    else:        
        profit =   around((float(input_arguments.price) - float(current_price))*float(input_arguments.quantity),2)
           
    print (current_price)
    
    if previous_price == 0:
        previous_price = current_price   
        ShowNotification(  str(current_price) ,'Welcome')
    else:
        print('Profit : ',profit)
    
    #        YESBANK.NS
    change = abs(100*(float(previous_price)-float(current_price))/float(previous_price) )
    
    if  change >= float(input_arguments.percent):
        previous_price = current_price   
        title = str(current_price)
        description =  'Profit : ' + str(profit) + '\nChange : ' + str(around(change,2))    
        ShowNotification( title ,description)   
    
    if input_arguments.repeat.lower() == 'no':
        break
         
    time.sleep(0.1)
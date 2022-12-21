import datetime
import win32com.client
import os
import time
import random

queue_info = win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
computer_name = os.getenv('COMPUTERNAME')

queues = 'queue_test_1'

def generate_body(ReferenceNo=None,Time=None,CheckSum=None,BuySell=None,Exchange=None,Code=None,Symbol=None,Series=None,StrikePrice=None,OptionType=None,ProtectionPercent=None,Qty=None,Price=None,ProClient=None,ClientID=None,BookType=None,DisclosedQty=None,TriggerPrice=None,Product=None):
    '''generate message body for saral money maker software and return body of the message as dict''' 
    body = {}
    if ReferenceNo:
        body["ReferenceNo"]=ReferenceNo
    if Time:
        body["Time"]=Time
    if CheckSum:
        body["CheckSum"]=CheckSum
    if BuySell:
        body["BuySell"]=BuySell
    if Exchange:
        body["Exchange"]=Exchange
    if Code:
        body["Code"]=Code
    if Symbol:
        body["Symbol"]=Symbol
    if Series:
        body["Series"]=Series
    if StrikePrice:
        body["StrikePrice"]=StrikePrice
    if OptionType:
        body["OptionType"]=OptionType
    if ProtectionPercent:
        body["ProtectionPercent"]=ProtectionPercent
    if Qty:
        body["Qty"]=Qty
    if Price:
        body["Price"]=Price
    if ProClient:
        body["ProClient"]=ProClient
    if ClientID:
        body["ClientID"]=ClientID
    if BookType:
        body["BookType"]=BookType
    if DisclosedQty:
        body["DisclosedQty"]=DisclosedQty
    if TriggerPrice:
        body["TriggerPrice"]=TriggerPrice
    if Product:
        body["Product"]=Product

    return body

def send_message(queue_name: str, label: str, message: str):
    '''Send message to MSMQ queue'''
    queue_info.FormatName = f'direct=os:{computer_name}\\PRIVATE$\\{queue_name}'
    queue = None

    try:
        queue = queue_info.Open(2, 0)

        msg = win32com.client.Dispatch("MSMQ.MSMQMessage")
        msg.Label = label
        msg.Body = message
        msg.Send(queue)

    except Exception as e:
        print(f'Error! {e}')

    finally:
        queue.Close()


def main():
    i = 0
    while True:
        i += 1
        body=generate_body()
        send_message(queues, 'test label', body)
        print(f'{i}: Message sent!')
        time.sleep(0.5)


if __name__ == '__main__':
    main()

# -*- coding: utf-8 -*-
#
#  get_data.py
#  Project: "market"
#
#  Created by Gennady Matveev on 12-05-2022.
#  Last modified on: 13-05-2022.
#  Copyright 2022. All rights reserved.
#
#  gm@og.ly
# %%
import telegram
from telegram.ext import Updater, CommandHandler, MessageHandler, Filters

from datetime import datetime, timedelta
import investpy as inv
import pandas as pd
import requests
from tabulate import tabulate

import smtplib
import ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
from email.header import Header

import credentials


# Helper function
def check_date(date):

    return (
        int(date[0:2]) > 0
        and int(date[0:2]) <= 31
        and int(date[2:4]) > 0
        and int(date[2:4]) <= 12
        and int(date[4:6]) > 0
        and int(date[4:6]) <= 25
    )


# Path to file with chat IDs

path = "./data/chats.txt"

# Bot commands


def start(update, context):
    update.message.reply_text("Bot started, usage: /send ddmmyy")
    chat_id = update.message.chat_id

    with open(path, "a+") as f:
        chats = f.readlines()
        if str(chat_id) + "\n" not in chats:
            f.write(str(chat_id) + "\n")


def help(update, context):
    update.message.reply_text("Usage: /send ddmmyy")


def send(update, context):
    tg_date = context.args[0]

    if check_date(tg_date):
        update.message.reply_text(
            "Thank you. Your request is being processed.")
        get_market_close(tg_date)
    else:
        update.message.reply_text(
            "Oops! Wrong date format - please use ddmmyy.")

# ------------------------- Get_data part -------------------------------------


def get_market_close(tg_date):

    # Define download dates

    close_date = tg_date[:2] + "/" + tg_date[2:4] + "/20" + tg_date[4:6]
    prev_date_ts = datetime.strptime(
        close_date, "%d/%m/%Y") - timedelta(days=1)
    prev_date = prev_date_ts.strftime("%d/%m/%Y")

    # Get data

    tickers = {
        "bonds": ["us1y", "us2y", "us3y", "us5y", "us7y", "us10y", "us20y", "us30y"],
        "indices": ["S&P 500", "Nasdaq", "Shanghai Comp. ", "MOEX Russia", "DXY"],
        "commodities": ["Gold", "Brent"],
    }
    assets = {}
    ind = {}
    df_assets = {}  # pd.DataFrame()
    for asset_class in tickers.keys():
        ind[asset_class] = []
        df_assets[asset_class] = pd.DataFrame()
        # Get tickers quotes
        for ticker in tickers.get(asset_class):
            # print(ticker, type(ticker))
            assets[ticker] = inv.search_quotes(
                text=ticker, products=[asset_class], n_results=1
            )
            try:
                data = (
                    assets[ticker]
                    .retrieve_historical_data(from_date=prev_date, to_date=close_date)
                    .iloc[-1:]
                )
                # print(ticker, data)
                df_assets[asset_class] = pd.concat(
                    [df_assets[asset_class], data], axis=0, sort=False
                )
                ind[asset_class].append(ticker)
            except:
                pass
        df_assets[asset_class].set_index(
            pd.Series(ind.get(asset_class)), inplace=True)

    # Add bonds spreads

    df_assets["bonds"].loc["2-10", "Close"] = (
        df_assets["bonds"].loc["us10y", "Close"] -
        df_assets["bonds"].loc["us2y", "Close"]
    )
    df_assets["bonds"].loc["5-20", "Close"] = (
        df_assets["bonds"].loc["us20y", "Close"] -
        df_assets["bonds"].loc["us5y", "Close"]
    )

    # Print results to txt file

    with open("./data/market_close.txt", "w") as f:
        print(f"Date: {close_date}\n", file=f)

        print(f"bonds\n", file=f)

        try:
            print(
                tabulate(
                    pd.DataFrame(df_assets["bonds"]["Close"]),
                    headers=["Years/spreads", "Close"],
                    tablefmt="grid",
                    floatfmt=".2f",
                ),
                file=f,
            )
        except:
            pass

        for a in ["indices", "commodities"]:
            print(f"\n{a}\n", file=f)

            try:
                print(
                    tabulate(
                        pd.DataFrame(df_assets[a]["Close"]),
                        headers=["Commodity    " if a ==
                                "commodities" else "Index", "Close"],
                        tablefmt="grid",
                        floatfmt=".2f",
                    ),
                    file=f,
                )
            except:
                pass

    # Make summary dataframe for xlsx export

    df_final = pd.DataFrame(
        pd.concat(
            [
                df_assets["bonds"]["Close"],
                df_assets["indices"]["Close"],
                df_assets["commodities"]["Close"],
            ],
            axis=0,
            sort=False,
        )
    )

    df_final.reset_index(inplace=True)
    df_final.rename(columns={"index": "Name"}, inplace=True)
    b = "Bonds," * (len(ind["bonds"]) + 2)  # add 2 for 210 and 520 spreads
    i = "Indices," * len(ind["indices"])
    c = "Commodities," * len(ind["commodities"])
    groups = list(filter(None, b.split(",") + i.split(",") + c.split(",")))

    df_final["Group"] = pd.Series(groups)
    df_final.set_index(["Group", "Name"], inplace=True)

    # Write results to xlsx file

    # Create a Pandas Excel writer using XlsxWriter as the engine

    writer = pd.ExcelWriter(
        "./data/market_close.xlsx", engine="xlsxwriter"
    )

    # Convert the dataframe to an XlsxWriter Excel object

    df_final.to_excel(writer, sheet_name=f"{tg_date}")

    # Get the xlsxwriter workbook and worksheet objects

    workbook = writer.book
    worksheet = writer.sheets[f"{tg_date}"]

    # Set format of data

    format1 = workbook.add_format({"align": "right"})
    format2 = workbook.add_format({"num_format": "#,##0.00", "border": 1})

    worksheet.set_column(0, 1, 18, format1)
    worksheet.set_column(2, 2, 10, format2)  # Width of cell 10
    worksheet.set_column("D:XFD", None, None, {
                         "hidden": True})  # Hide empty columns

    # Close the Pandas Excel writer and output the Excel file

    writer.save()

    # Send mail

    # Set operation mode: production ('p') or test ('t')

    switch = "t"

    send_email(credentials.login["email"],
               credentials.login["password"], switch)

    # Telegram bot part

    # Read results from txt file

    # with open("./data/market_close.txt", "r") as f:
    #     txt = f.read()  # .replace('\n', '')

    # Send results to Telegram

    path = "./data/chats.txt"
    token = credentials.t_bot_token
    end_msg = "Tutto opossum - please check your mail"  # txt
    with open(path, "r") as chats:
        for chat in chats.readlines():
            results = requests.get(
                f"https://api.telegram.org/bot{token}/sendMessage?chat_id={chat}&text={end_msg}"
            )

# --------------------------------- Email part --------------------------------


def send_email(sender_mail, sender_pass, switch):

    # Create email
    subject = "market close"
    body = "--\nС уважением,\n\nМ.Г."
    sender_email = sender_mail
    from_sender = formataddr(
        (str(Header('Меркурий Гусевич О.', 'utf-8')), sender_email))
    # Production/test switch
    receiver_email = credentials.addressee[switch][0]
    cc_email = credentials.addressee[switch][1]

    password = sender_pass

    # Create a multipart message and set headers
    message = MIMEMultipart()
    message["From"] = from_sender
    message["To"] = receiver_email
    message["Subject"] = subject
    message["Cc"] = cc_email

    # Add body to email
    message.attach(MIMEText(body, "plain"))

    filename = "./data/market_close.xlsx"

    # Open xlsx file in binary mode
    with open(filename, "rb") as attachment:
        # Add file as application/octet-stream
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    # Encode file in ASCII characters to send by email
    encoders.encode_base64(part)

    # Add header as key/value pair to attachment part
    file_name = filename.rsplit('/', 1)[1]
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {file_name}",
    )

    # Add attachment to message and convert message to string
    message.attach(part)
    text = message.as_string()
    recipients = [receiver_email] + [cc_email]

    # Log in to server using secure context and send email
    try:
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
            server.login(sender_email, password)
            server.sendmail(sender_email, recipients, text)
    except:
        print("Something went wrong - can't send email...")
        # pass


def main():
    token = credentials.t_bot_token
    updater = telegram.ext.Updater(token)
    dispatcher = updater.dispatcher

    dispatcher.add_handler(CommandHandler("start", start))
    dispatcher.add_handler(CommandHandler("help", help))
    dispatcher.add_handler(CommandHandler("send", send, pass_args=True))

    # print("Ready")
    updater.start_polling(poll_interval=10.0)
    updater.idle()


if __name__ == "__main__":
    main()

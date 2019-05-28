# Morningstar-Financials
## Introduction
This is my first Windows Powershell Project. This tool helps to extract the financial information from www.morningstar.com, which aims to help investors automate the process of copying financials out for their own analysis and chart plotting. The financial information is saved in a .csv file located at the same directory of where the tool is located.

The following financial information are extracted

**From Income Statement:**
1) Revenue
2) Cost of Goods Sold (COGS)
3) Gross Profit
4) Operating Income
5) Total number of shares

**From Balance Sheet:**
1) Total Cash
2) Total Debt
3) Total Equity
4) Inventory

**From Cash Flow Statement:**
1) Operating Cash Flow (OCF)
2) Capital Expenditure (CAPEX)
3) Free Cash Flow (FCF)

**Stock Exchanges Supported:**
1) HKEX
2) NASDAQ
3) NYSE AMERICAN
4) NYSE

*NOTE: There's a difference between NYSE American and NYSE.*

## Usage

1) Download and save the script "ExtractFinancials.ps1" in your desired directory
2) From the directory, "Right-click" the file -> "Run with PowerShell"
3) Select the stock exchange<br />
    [1] HKEX<br />
    [2] NASDAQ<br />
    [3] NYSE AMERICAN<br />
    [4] NYSE<br />
4) Key in the stock's ticker symbol. If the stock is a HKEX stock, key in the 4 digit stock code.
5) Wait...
6) The .csv file will be saved in the directory as "Financials_(ticker symbol).csv"
*7) If you encounter any errors, please exit the program and retry from step 2.*

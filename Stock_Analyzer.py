import yfinance as yf
import numpy as np
import pandas as pd

aapl = yf.Ticker("AAPL")
print(aapl.quarterly_balancesheet)
print(aapl.info)

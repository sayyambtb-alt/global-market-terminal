# Global Market Terminal

A Streamlit workstation for researching equities, valuations, technical charts and macro data.

- **Fundamental Deep Dive** covers company financials and related context.
- **Valuation Modeler** covers assumptions and valuation scenarios.
- **Technical Analysis** shows price history, 50- and 200-close moving averages, Bollinger Bands, RSI, MACD and volume. Choose the listed security and then a chart period.
- **Macro & CAPM** covers macro comparisons and the cost of equity.

## Read a market chart

Open **Technical Analysis** and expand **How to read these charts**. The price chart labels the currency and latest bar returned by Yahoo Finance. The daily averages need 50 or 200 observations, so they may be blank on shorter periods or recently listed securities. Check the bar date before using a chart as a current market reference. The app caches price history for up to one hour.

## Run locally

```bash
python -m pip install -r requirements.txt
streamlit run app.py
```

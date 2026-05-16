# Persona 7 — Day traders + quant hobbyists

Status: **done** · Last revised: 2026-05-16

## TL;DR (5 lines)

- Two adjacent subsegments: active retail day traders (~600k+ US, post-2020 boom durable) and quant hobbyists / algo-trading enthusiasts (~50–100k US in r/algotrading orbit). **Both already pay for ticker tape on a monitor — wallpaper-mode is selling them what they already buy.**
- Hardware culture: **3–6 monitors is the brag**. The unused real estate behind charting tools is enormous. QuickSheet wallpaper is the highest-density use of that real estate they've ever seen.
- Highest-leverage cells: live ticker row with %change colours, P/L cell (today / month / YTD), watchlist with alerts, options chain mini-table, crypto + FX side-row, news headline strip.
- Where to seed: r/Daytrading, r/wallstreetbets (cautiously — needs *trader* angle not meme), r/algotrading, r/options, r/CryptoCurrency, r/quant, Elite Trader forum, TradingView's Pine Script community, FinTwit on X/Bluesky.
- Direct competition: Bloomberg ($24k/yr — not a competitor), TradingView (browser, $$$ for live), thinkorswim (broker-locked), TradingView desktop (paid). **Free + offline + your-data-here is genuinely novel for this audience.**

## 1. Profile

### 1a. Retail day traders

- Hardware: gaming-rig-grade desktops with 3–6 monitors. Religion-level commitment to monitor count. Average new-trader buys hardware *before* learning to trade.
- Symbols watched: 5–50 tickers, ~10–30 options contracts, often + crypto + FX overlay.
- Income shape: highly variable. ~10% of active traders profitable consistently. Most pay subscription costs (data feeds, screeners).
- Buying brain: "Will this make me a better trader, or is it noise?" Sophisticated enough to see through marketing; allergic to "AI signals" hype.

### 1b. Quant hobbyists + algo-trading enthusiasts

- Overlap with developers-sre (often the same people). Backtesting in Python (pandas, vectorbt, backtrader), running personal algos on broker APIs (IBKR, Alpaca, Tradier).
- Hardware: same multi-monitor setup, plus a "research" monitor for Jupyter/QuantConnect.
- Cultural moment: Robinhood-era retail boom durable; LLMs revived interest in algo-trading 2024-2026; "AI trading bot" hype + skepticism both high.

## 2. Why their desktop is wasted

- Already chock-full of TradingView charts. But:
  - Charts cover charts. There's no *meta* layer showing P/L, win rate, today's trades summary.
  - News feeds live in a separate window or get missed.
  - Watchlists are inside TradingView (locked).
  - Crypto + FX + equities are siloed across different windows/sites.
- Maximised charts on monitors 2, 3, 4. The wallpaper *behind* the charts is dead space.
- Worse: on the *primary* monitor (where the IDE/spreadsheet lives during research), the wallpaper is pure static image.

What QuickSheet replaces: nothing chart-shaped. **It is the orientation layer on top of the chart layer.** Live P/L total, watchlist with sparklines, news strip, alerts.

For quant hobbyists specifically: a wallpaper grid of *strategy backtest results* — live "today's PnL for strategy A vs B vs C" while charts are doing their thing on other monitors.

## 3. Glanceable data they actually want behind windows

### Day traders

1. **Today's P/L cell** — single big number, red/green, biggest font available. The day's score.
2. **Open positions row** — ticker, qty, avg cost, current, %P/L. Coloured.
3. **Watchlist with sparklines** — N tickers, last price + intraday sparkline + %change.
4. **VIX / DXY / SPY snapshot** — market-context row.
5. **Crypto + FX side-row** — BTC, ETH, EUR/USD, USD/JPY. (Existing `price:` + `fx:` extensions cover this.)
6. **News strip** — headlines from a chosen RSS/CNBC feed.
7. **Alerts row** — "AAPL above 200" / "RSI<30 on XLF" — fired alerts in red, armed ones in dim.
8. **Account equity curve** — sparkline of last 30 days.

### Quant hobbyists

1. **Per-strategy live P/L** — N strategies, each with a P/L cell + sparkline of equity curve.
2. **Latest backtest result** — cell showing sharpe / max drawdown / CAGR for the last `python backtest.py` run.
3. **Open orders** from broker API.
4. **API rate-limit counters** — "IBKR 30/50 req/sec." Same staleness/health vibe as SRE.
5. **Algo health** — green/red dot per running strategy.

## 4. Candidate extensions / desktop-only features (ranked)

| Rank | Item                                       | Type | Cost | Hit prob | Why                                                                              |
|------|--------------------------------------------|------|------|----------|-----------------------------------------------------------------------------------|
| 1    | `quote:` Yahoo Finance / Stooq quote        | ext  | low  | high     | `stock:` already exists for daily close; need *intraday* quote. Free Yahoo JSON. |
| 2    | **Sparkline cell tied to live ticker**     | feat | low  | high     | `s: A1::A30` already exists; add `L:` loop pattern for live tickers.            |
| 3    | **Value-driven cell colour** (sign-aware)  | feat | low  | high     | Negative P/L red, positive green. Same shared feature, 7th persona to want it.  |
| 4    | `vix:` / `dxy:` / `spy:` market-context     | ext  | low  | high     | Bundle inside `quote:` with friendly aliases.                                    |
| 5    | `news:` extension *(already shipped)*       | ext  | —    | —        | Re-pitch with FinTwit RSS feed example.                                          |
| 6    | `alert:` rule-fires-when-price-crosses     | ext  | med  | high     | `alert: AAPL>200` → cell turns red and persists. Genuinely new product surface.  |
| 7    | `iv:` options implied-vol lookup            | ext  | high | med-high | CBOE / Yahoo options chains. Bigger lift; hits options traders specifically.    |
| 8    | `port:` portfolio CSV → live P/L            | ext  | low  | high     | User provides a `positions.csv`; ext multiplies qty × current. The P/L cell.    |
| 9    | `ibkr:` / `alpaca:` broker connector        | ext  | high | high     | Auth-heavy but the dream extension for algo crowd. Defer until ext-3-5 ship.    |
| 10   | `backtest:` read pandas-backtest JSON       | ext  | low  | med      | Reads a sidecar JSON your Python script writes. No web call.                    |

Note: `stock:`, `fx:`, `price:`, `news:` already shipped. The audience needs to be *told these exist together as a finance dashboard.*

## 5. Where they hang out (multi-monitor culture first)

- **r/Daytrading (1M), r/options (1.6M), r/wallstreetbets (16M, but post quality matters — pick `r/Daytrading` for tool talk)**.
- **r/algotrading (3M)** — closest match for the quant hobbyist subsegment; tool-curious.
- **r/cryptocurrency (8M), r/CryptoMarkets, r/btc, r/ethfinance**.
- **r/quant (350k), r/quantfinance** — degenerate signal but receptive to free tools.
- **Elite Trader forum** — long-form, screenshots of multi-monitor setups frequently posted.
- **TradingView community + Pine Script forums** — they share custom indicators eagerly; an "open-source dashboard companion" pitch lands.
- **FinTwit on X / Bluesky** — @ FinTwit accounts share interesting tools; tradertools-friendly bluesky scientist crowd growing.
- **Discords:** Trader Republic, Bull Trader chat, AlgoTrading101.
- **r/battlestations + r/desktops** — multi-monitor screenshot culture. Wallpaper-tracker-of-PnL would be a *huge* post here.
- **People to be visible to:** @QuantStratTrader, @PythonForFinance accounts, Adam Butler (ReSolve Asset Mgmt), QuantStart-era educators. **Plus**: TradingView's product team (they share third-party tools).

## 6. Discoverability hooks

Hero image = **a 3-monitor setup**. Center monitor: TradingView with SPY chart. Left monitor: thinkorswim. Right monitor: VSCode + Jupyter. **Wallpaper across all three (or just the right one)**: QuickSheet grid with today's big-red `-$1,247` P/L cell, 5 ticker rows with sparklines, news strip, alerts.

Headlines that land:

- "Built a free open-source dashboard for day traders — runs as your wallpaper"
- "I track P/L, watchlist, and alerts on my desktop wallpaper. Bloomberg can keep their $24k/yr."
- "Pine Script taught me to draw on charts. Now my charts have a wallpaper dashboard."
- "[r/battlestations] My 3-monitor day trading setup with a wallpaper PnL grid"

Avoid: "AI signals," "guaranteed wins," anything that pattern-matches to scam tools. **Truth-first, monitor-rice-first, code-on-github.**

Don't lead with TUI. The audience is graphical-first.

## 7. Implications (queue these)

1. **Build `quote:` extension** (Bucket F, post-research). Free Yahoo JSON; intraday refresh; hits the entire persona. Pairs with already-shipped `stock:` (daily close), `fx:`, `price:` (crypto), `news:` to form a finance bundle.
2. **Build `alert:` extension or feature** (rank 6 — genuinely novel product surface). A cell that watches another cell and flips state. Could be a feature on the main repo (`a: A3>200` → cell turns red when triggered) — lower cost than a per-data-source extension.
3. **Bucket A: `examples/daytrader-dashboard.csv` + `examples/algotrader.csv` + `docs/for-traders.md`.** Two complete starter sheets. The image of these sheets is *the entire post*.
4. **Bucket C: r/battlestations / r/desktops multi-monitor screenshot post.** Likely the second-most-viral persona after gamedev-ttrpg DM-screen. Save in `drafts/traders-launch.md` once `quote:` and `alert:` ship.
5. **Re-frame the existing shipped extensions** (`stock:`, `fx:`, `price:`, `news:`) as a "finance bundle" in `docs/for-traders.md`. Zero new code; just packaging.

Cross-link:
- [developers-sre.md](developers-sre.md) — same audience overlap for quant subsegment; same staleness/colour features.
- [accounting-extensions.md](../accounting-extensions.md) — `1099:`, `qtr:` overlap for self-employed traders.

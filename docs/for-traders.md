# QuickSheet for traders

If you trade equities, options, FX, or crypto and you already own multi-monitor real
estate — **your wallpaper can be the orientation layer** on top of your chart layer.
P/L, watchlist, FX side-row, news strip, all painted *behind* your TradingView /
thinkorswim windows. No browser tab, no SaaS, your positions stay on your machine.

> Not a replacement for charts. A complement: the layer that tells you "where am I,
> what's happening, what should I notice" while charts are doing their thing.

## Five-minute setup

```bash
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet
dotnet run -c Release --project ExcelConsole.csproj -- --desktop examples/trader-dashboard.csv
```

Edit cells, autosaves to CSV every 5s. Add to startup applications and the grid is
there every market open — *behind* every window, never stealing focus.

## What goes on the wallpaper

Starter sheet at `examples/trader-dashboard.csv`. Columns:

| Column           | What it shows                                                  | Extension                                                                                |
|------------------|----------------------------------------------------------------|-------------------------------------------------------------------------------------------|
| **Equities**     | Daily close + intra-day change for each ticker                 | [`stock`](https://github.com/Deskworks/quicksheet-stock-ext) (Stooq, no API key)          |
| **Crypto**       | Last trade + 24h change                                        | [`price`](https://github.com/Deskworks/quicksheet-price-ext) (CoinGecko, no API key)      |
| **FX**           | Live currency conversion, 200+ pairs                           | [`fx`](https://github.com/Deskworks/quicksheet-fx) (ECB rates, no API key)                |
| **News strip**   | Headlines from any RSS/Atom feed (FinTwit aggregators, BBC, …) | [`news`](https://github.com/Deskworks/quicksheet-news)                                    |
| **Tax pacing**   | Days to next IRS estimated tax deadline                        | [`qtr`](https://github.com/Deskworks/quicksheet-qtr)                                      |
| **SE tax**       | Self-employment tax estimate for full-time traders             | [`1099`](https://github.com/Deskworks/quicksheet-1099-ext)                                 |
| **Sparklines**   | Mini intraday-equity chart from a range of cells               | built-in `s: A1::A30`                                                                     |

All six extensions install with a single `ext: github:Deskworks/quicksheet-<name>` cell.

## What the wallpaper IS NOT for

- **Order entry.** QuickSheet has no broker integration. Use your broker's app for trades.
- **Real-time level-2 quotes.** `stock` and `price` are end-of-day / last-trade APIs;
  if you need tick data, your broker's terminal owns that screen.
- **Charting.** Use TradingView / thinkorswim / Sierra Chart. The wallpaper sits
  beside or behind those, not in front of them.

It's the *meta* layer: P/L state, watchlist sparklines, today's tax pacing, news
headlines. The boring stuff that should be glanceable, not the active surface.

## Recipe: a watchlist with intra-day sparkline

1. One row per ticker.
2. Column A: ticker symbol.
3. Column B: `stock: <ticker>` for the day's close + change.
4. Columns C–L: collect 10 intra-day samples (manually or via an `L:` loop and
   `i: curl ...`).
5. Column M: `s: C1::L1` — sparkline of the last 10 samples.
6. Wrap critical P/L cells in `c:red: ...` / `c:green: ...` when you spot a level.

## Pair with the freelance / 1099 cluster

If you trade full-time as your own LLC / sole proprietorship:

- [`qtr`](https://github.com/Deskworks/quicksheet-qtr) counts down to your next
  estimated-tax deadline.
- [`1099`](https://github.com/Deskworks/quicksheet-1099-ext) estimates SE tax from
  your YTD gross.
- [`rate`](https://github.com/Deskworks/quicksheet-rate) calculates the minimum
  viable hourly rate against your target income (useful for anyone consulting on
  the side of trading).
- Existing built-in CSV columns track your YTD realised P/L; the wallpaper makes
  sure the number is visible every time you tab out of your broker.

## Make it pretty

Cycle themes with **Ctrl+T**: Dark, Light, Nord, Solarized, Matrix, Dracula,
Synthwave, Gruvbox, Monokai, HotdogStand. Synthwave or Matrix fit the
trader-desk aesthetic; Light + a clean monospaced font is what you want during
prep.

## What's missing (be honest)

- No alert / threshold notifications. If you want "AAPL crossed 200 → desktop
  notification," that's not on the wallpaper today.
- No intraday quote extension yet; `stock` is daily-close. PRs welcome.
- No broker connector (IBKR / Alpaca / Tradier). Same answer — PRs welcome.

These are intentional gaps; the wallpaper-as-spreadsheet philosophy is that you
own a row of CSV cells, not a notification stack.

## More to mix in

- [`fx`](https://github.com/Deskworks/quicksheet-fx) for daily FX conversion side-row.
- [`pomo`](https://github.com/Deskworks/quicksheet-pomodoro) if you trade in
  defined sessions and want a session-boundary marker.
- [`news`](https://github.com/Deskworks/quicksheet-news) subscribed to a FinTwit
  aggregator RSS for headline scroll.
- [Full extension directory](extensions.md).

## Build your own (50-line broker integrations)

Brokers all have a quote endpoint. The QuickSheet extension protocol is two
JSON-lines messages ([extension-protocol.md](extension-protocol.md)) — write a
50-line extension in Python / Go / .NET that wraps your broker's quote API and
publish it as `quicksheet-<broker>-ext`. The wallpaper picks it up via
`ext: github:user/quicksheet-<broker>-ext`.

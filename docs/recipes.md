# Wallpaper dashboard recipes

Concrete layouts you can drop into QuickSheet (`--desktop` mode) to make the wallpaper genuinely useful. Each recipe is a small CSV you can paste into a fresh sheet and adapt.

> All recipes assume the relevant extensions are installed. Install once with `ext: github:cemheren/<name>` and the prefix is registered for every future cell.

---

## 1. Ops on-call dashboard

Hosts you care about, their TLS expiry, their HTTP status, their MX records. One glance at the wallpaper tells you whether anything is on fire.

```
Service,Endpoint,HTTP,TLS,MX
api,api.example.com,"ping: https://api.example.com, 1, 3","tls: api.example.com, 1, 4","mxck: example.com, 1, 5"
auth,auth.example.com,"ping: https://auth.example.com, 1, 3","tls: auth.example.com, 1, 4",
docs,docs.example.com,"ping: https://docs.example.com, 1, 3","tls: docs.example.com, 1, 4",
```

Then wrap the `ping:` cells in a loop:

```
L: C2, 1m
L: C3, 1m
L: C4, 1m
```

Now the status column refreshes every minute. TLS only changes daily — point a slower loop at the TLS column:

```
L: D2, 720m
```

Extensions used: [ping](https://github.com/Deskworks/quicksheet-ping-ext), [tls](https://github.com/cemheren/quicksheet-tls-ext), [mxck](https://github.com/Deskworks/quicksheet-mxck-ext).

---

## 2. Quiet portfolio glance

Crypto holdings, no app to open, no browser tab.

```
Holding,Position,Quote,Change
BTC,0.5,"price: btc, 1, 2",
ETH,4.2,"price: eth, 1, 2",
SOL,80,"price: sol, 1, 2",
```

`L: C2, 5m` to refresh once every five minutes.

The price extension auto-caches for 60 seconds, so the loop never hammers the API even at faster cadence.

Extensions used: [price](https://github.com/Deskworks/quicksheet-price-ext).

---

## 2b. Stock watchlist

Same idea, but for equities and ETFs.

```
Ticker,Quote,Notes
AAPL,"stock: AAPL, 1, 3",
MSFT,"stock: MSFT, 1, 3",
SPY,"stock: SPY, 1, 3",
VOO,"stock: VOO, 1, 3",
```

`L: B2, 30m` is plenty — Stooq's data is end-of-day for most exchanges, and intraday updates lag 15+ minutes anyway.

Extensions used: [stock](https://github.com/Deskworks/quicksheet-stock-ext).

---

## 3. Personal command center

Launchers + bookmarks + small numeric tracker in one sheet.

```
Open today,,,Weight,Steps,Sparkline
r: code .,r: firefox github.com,,72.1,5421,"s: D2::D8"
r: slack,r: zoom,,72.4,8210,
r: notes,r: terminal,,71.8,6100,
```

The sparkline cell uses the range form (`s: A1::A10`) — it pulls live values from a column of your weights, so just typing today's number in `D2..` updates the chart automatically.

---

## 4. Writer's reference panel

```
Word,Definition,Notes
laconic,"def: laconic, 1, 2",
prolix,"def: prolix, 1, 2",
eponymous,"def: eponymous, 1, 2",
```

Wrap the column in a `L:` if you want them to refresh, though definitions don't change.

Extensions used: [define](https://github.com/Deskworks/quicksheet-define-ext).

---

## 5. Pomodoro + tasks side-by-side

```
Timer,Tasks
"pomo: 25, 3, 2",1. ship the refactor
,2. review PR #142
,3. lunch
```

When the timer runs, the rest of the row stays where it is — the timer cell expands within its allotted span.

Extensions used: [pomodoro](https://github.com/Deskworks/quicksheet-pomodoro).

---

## 5b. Freelancer dashboard

Track invoices, set-aside tax, mortgage payment, and a running total of monthly net — the whole "am I going to make rent and pay quarterlies" calculation in one row.

```
Invoice,Gross,SE tax,Mortgage,Notes
Acme Co,5200,"1099: 62400, 1, 5","mort: 380000, 6.5, 30, 1, 4",last month
Globex,3800,"1099: 45600, 1, 5","mort: 380000, 6.5, 30, 1, 4",this month
```

Σ of column B in the status bar gives your YTD gross. The `1099:` cell projects what SE tax you owe at that annualized rate. The `mort:` cell stays constant — that's your fixed cost.

Pair with a sparkline of monthly gross in column F (`s: B2::B13`) for an at-a-glance income curve.

Extensions used: [1099](https://github.com/Deskworks/quicksheet-1099-ext), [mortgage](https://github.com/Deskworks/quicksheet-mortgage-ext).

---

## 6. Academic writing reference

Definitions and live citation lookups in one sheet — drop a DOI in column B and the formatted citation fills column C, type a word in column D for an inline definition.

```
DOI,Citation,Term,Definition
10.1145/3623476.3623525,"cite: 10.1145/3623476.3623525, 1, 4",laconic,"def: laconic, 1, 2"
10.1109/MS.2021.3070752,"cite: 10.1109/MS.2021.3070752, 1, 4",heuristic,"def: heuristic, 1, 2"
```

Citations are cached forever in the extension subprocess (DOIs don't change), so loading is fast on repeat views.

Extensions used: [cite](https://github.com/Deskworks/quicksheet-cite-ext), [define](https://github.com/Deskworks/quicksheet-define-ext).

---

## 6b. HN reading list

Hacker News tabs, but on the wallpaper. Pin the URLs you actually come back to, watch latency to make sure they're up, and use `i:` for a live count of stories you might want to read.

```
What,Link,Status,Comments
news.ycombinator.com,https://news.ycombinator.com,"ping: https://news.ycombinator.com, 1, 3",
Show HN,https://news.ycombinator.com/show,"ping: https://news.ycombinator.com/show, 1, 3",
Ask HN,https://news.ycombinator.com/ask,"ping: https://news.ycombinator.com/ask, 1, 3",
top story,,"i: curl -s 'https://hacker-news.firebaseio.com/v0/topstories.json' | head -c 200 | wc -c",first 200 chars
```

Multi-select column B, hit Enter to open all four in your browser. Add `L: <cell>, 30m` on the status column for a passive uptime indicator.

Extensions used: [ping](https://github.com/Deskworks/quicksheet-ping-ext). The `i:` cell is built-in.

---

## 7. AI scratchpad

Drop in a cell that asks Copilot to summarize whatever else is in the sheet:

```
copilot: summarize the data in {A1::E20}, 5, 1
```

This makes the wallpaper a live "what am I looking at?" surface.

Extensions used: [copilot](https://github.com/Deskworks/quicksheet-copilot-ext).

---

## Tips

- **Multi-select + Enter** triggers many cells at once. Useful for "rebuild the whole dashboard" on demand.
- **`L:` loops are per-cell**, so set faster cadences only where you need them. Most cells don't need refresh.
- **Σ and Π** in the status bar show the column sum and row product live — useful even on cells that look textual, since extension outputs that happen to be numeric also count.
- **CSV is the file**. You can keep multiple recipes as separate `.csv` files and point QuickSheet at whichever one you want.

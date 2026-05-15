# Copilot extension — use cases

Most QuickSheet extensions are deterministic (TLS checker, mortgage calc, MX lookup). The [Copilot extension](https://github.com/cemheren/quicksheet-copilot-ext) is the open-ended one: a cell that runs an LLM prompt with optional access to other cells in the sheet via `{A1::C10}` range references.

The result is that a single sheet can be both a static layout *and* a live AI surface — drop in a prompt, point it at the data you already have, get a row of cells filled.

This page collects use cases that map naturally onto a wallpaper-pinned sheet. Each one is one or two `copilot:` cells you can paste into a fresh `.csv` and adapt.

## Install

```
ext: github:cemheren/quicksheet-copilot-ext
```

After install, `copilot: <prompt>, <cols>, <rows>` activates the prompt and writes the result into a `cols × rows` block.

---

## 1. Summarize a column

```
A,B
todo: write release notes,
todo: replicate the bug from #42,
todo: book flights,
"copilot: summarize as 3 bullets {A1::A3}, 1, 3",
```

Activate cell `B1`. The bullets fill `B1..B3`. Use this when you have a freeform-text column and want a structured digest right next to it.

## 2. Classify a row

```
expense,amount,category
electricity bill,82.40,"copilot: classify as utilities/food/transport/personal/other given description {A2::B2}, 1, 1"
amazon prime,14.99,"copilot: classify as utilities/food/transport/personal/other given description {A3::B3}, 1, 1"
```

Drop the same prompt down a column to label every row. Combined with `Σ` in the status bar this is a one-screen expense tracker.

## 3. Generate test data

```
copilot: generate 10 plausible customer rows with name email signup_date, 3, 10
```

Single-cell prompt that fills a 3×10 block. Useful when prototyping — replace the contents in seconds when the schema changes.

## 4. Extract fields from a freeform note

```
A
"call with sarah re: q3 budget, action items: send the revised forecast by friday, follow up on the hiring freeze"
"copilot: extract action items as a checklist {A1::A1}, 1, 5"
```

Take a wall-of-text meeting note in `A1`, fill 5 cells beneath with extracted action items. Loop the cell with `L: A2, 1h` and the action items refresh as you edit the source.

## 5. Compare two cell ranges

```
copilot: what changed between {A1::C10} (before) and {E1::G10} (after) — list as bullets, 1, 5
```

Two snapshots of the same data, before and after. Lives well as a dashboard cell showing "what did I change in this sheet today."

## 6. Translate / rephrase

```
A,B
hej hur mår du,"copilot: translate to english {A1::A1}, 1, 1"
buenos días,"copilot: translate to english {A2::A2}, 1, 1"
```

Pair with [`def:`](https://github.com/cemheren/quicksheet-define-ext) for a "language scratchpad" cell layout.

## 7. Full dashboard composition

A useful pattern: one row of `copilot:` cells, each pointing at a different range, producing a multi-cell summary header.

```
A,B,C,D
todo: ...,"copilot: top priority from {A1::A20}, 1, 1","copilot: count overdue items in {A1::A20}, 1, 1","copilot: estimate total effort in {A1::A20}, 1, 1"
...
```

`Ctrl+L` (rerun cell) on any of B/C/D refreshes that column's summary; the source list keeps editing freely.

---

## Tips

- **Range refs run live.** When Copilot reads `{A1::C10}` it gets the current cell contents at activation time, so editing source cells and re-activating the prompt gives a fresh answer.
- **Output shape matters.** The prompt's `<cols>, <rows>` tells Copilot how many cells to fill. Asking for "5 bullets" with `1, 5` gives one per cell; asking with `1, 1` gives a single concatenated cell.
- **Loops on `copilot:` cells are expensive.** They cost an API call every fire. Prefer `L: <cell>, 1h` or longer for AI cells. `r:` and `i:` cells are cheap to loop.
- **Pair with deterministic exts.** `def:` for vocabulary, `mort:` for math, `stock:` for prices, `copilot:` for the language layer — the same row can mix all of them.

## See also

- [`quicksheet-copilot-ext`](https://github.com/cemheren/quicksheet-copilot-ext) — the extension repo.
- [docs/extensions.md](extensions.md) — full extension directory.
- [docs/recipes.md](recipes.md) — broader dashboard recipes that build on these ideas.

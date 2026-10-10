# Daily Gold Reconciliation System

A daily stock, rate and margin reconciliation model for a retail gold jeweller, built on a
fine-gold basis so that metal held at different purities is actually comparable.

**1,675 formulas across 9 sheets. All data in this repository is synthetic.**

---

## The problem

A retail jeweller holds the same metal in four places at once:

| Where | At what purity |
|---|---|
| On the shelf as finished stock | mixed, mostly 22K |
| Out with karigars (goldsmiths) being worked | as issued, usually 22K |
| In scrap awaiting melt | mixed |
| In old gold taken in part-exchange from customers | unknown until tested |

Each sits at a different purity, and **10 grams of 22K is not 10 grams of 24K.** 22K is 0.916
fine, so 10g of it contains 9.16g of actual gold. Adding gross weights across those four
buckets produces a number that does not mean anything.

This is not a pedantic point. It is the reason a shop that counts in gross weight cannot tell
a genuine loss from a shift in its karat mix. If last month's stock was heavier in 22K and
this month's is heavier in 18K, gross weight falls with no metal missing at all. A real theft
and a mix shift look identical on the only number most shops actually track.

So every quantity in this model is converted to **fine gold** before anything is added or
compared:

```
fine gold (g) = gross weight (g) x purity
```

On that basis the shelf, the karigar book, the scrap tray and the old gold drawer are all
denominated in the same unit, and the daily reconciliation becomes arithmetic instead of
guesswork.

---

## What it answers every day

### 1. Does counted physical stock match the book?

```
closing book = opening + purchases - issued to karigar + returned from karigar
               - sold + old gold taken in
variance     = physical counted - closing book
```

Reported in grams fine and converted to rupees at that day's spot, against a tolerance set in
`Setup`. Issues, returns, sales and old gold are all pulled from their source registers by
`SUMIFS` on date, so the ledger cannot silently drift away from the underlying records.

### 2. Is the board rate still tracking the market?

A jeweller quotes customers from a board rate, which somebody has to update by hand. The model
compares it to spot on a common fine basis:

```
board implied fine = board rate 22K / 0.916
drift              = board implied fine - spot fine
cost of the drift  = drift x fine gold sold that day
```

A board rate left stale while spot rallies means **every ticket that day sold metal below
replacement cost.** The sample data resets the board rate only every fourth day on purpose, to
generate exactly that failure.

### 3. How much making charge is being given away?

Making charge is the labour component, and it is where almost all of a jeweller's real margin
lives. Metal is close to a pass-through. The model tracks list making charge against realised,
line by line:

```
MC given away = MC list - MC realised
realisation % = MC realised / MC list
```

This is the most commonly unmeasured leak in a jewellery P&L, because a discount at the counter
is invisible unless somebody is differencing list against realised on every line.

---

## What the sample data surfaces

Running the 22 business days of synthetic October data:

**One karigar is quietly running 63% over his wastage norm.**

| Karigar | Trade | Fine gold issued | Wastage | Norm | vs Norm | Lines flagged |
|---|---|---|---|---|---|---|
| K-01 | Chain making | 1,645.9 g | 1.88% | 1.80% | +0.08pp | 0 |
| K-02 | Chain making | 1,048.4 g | 1.84% | 1.80% | +0.04pp | 0 |
| **K-03** | **Stone setting** | **1,571.3 g** | **5.72%** | **3.50%** | **+2.22pp** | **10** |
| K-04 | Casting | 1,346.3 g | 2.61% | 2.50% | +0.11pp | 0 |
| K-05 | Hand finishing | 1,359.9 g | 2.58% | 2.80% | -0.22pp | 0 |

Every other karigar sits within 0.11pp of his norm. K-03 is 2.22pp over, on 1.57 kg of fine
gold issued, which is roughly 35 grams of unexplained loss in a single month. Nothing about any
individual job looks wrong. It is only visible in aggregate, measured against a norm that is
specific to his trade, on a fine basis.

The rest of the month:

- **4 of 22 days** breached the stock variance tolerance, the worst at -4.18 g fine, about ₹40,000
- **4 of 22 days** needed the board rate repriced, with drift reaching **-₹243/g fine** against spot
- Making charge realisation landed at **85.0%**, right on the floor, with **₹201,658** given away in a month

---

## Sheets

| Sheet | What it holds |
|---|---|
| `README` | In-workbook version of this explanation |
| `Setup` | Purity table, wastage norms by trade, variance and rate tolerances. Yellow cells are the inputs |
| `Rate_Sheet` | Daily spot, board rate, drift, reprice flag |
| `Metal_Ledger` | Opening, purchases, issues, returns, sales, old gold, closing book vs counted physical |
| `Karigar_Ledger` | Metal out per karigar, what returned, wastage against the norm for that trade |
| `SKU_Register` | Product master: gross, stone, net gold, fine gold, list making charge |
| `Sales_Register` | Line-level sales: metal value, MC list vs realised, old gold credit, net invoice |
| `Daily_Recon` | The one-screen daily output. Change the date in the yellow cell |
| `Monthly_Analysis` | Monthly rollup plus per-karigar wastage for the whole period |

Conventions follow standard financial-model practice: blue text for hardcoded inputs, black for
formulas, green for cross-sheet links, yellow fill for cells meant to be edited. Every
tolerance and norm is a labelled cell referenced by the formulas that consume it, never a
number buried inside one.

---

## Design notes

**Wastage is measured on fine gold issued, not gross weight.** Measured on gross, a karat mix
shift is indistinguishable from a loss, which is the whole failure mode this model exists to
fix. Applying a gross-basis norm to a fine-basis quantity would reintroduce it.

**Norms are per trade, not per shop.** Stone setting legitimately loses more metal than chain
making. A single shop-wide norm either clears genuinely bad chain work or condemns perfectly
normal stone setting. `Setup` holds one norm per trade and `Karigar_Ledger` looks up the right
one by the karigar's trade.

**Metal is billed at the board rate, not at spot.** That is deliberate, and it is what makes
the rate drift measurable instead of theoretical. `Sales_Register` column H links to the board
rate for that date, so the drift flows straight through to the margin on each line.

**Cumulative karigar exposure is tracked separately** from the daily issue and return. The
running unreturned balance is the shop's actual credit exposure to its goldsmiths, and it is
not visible from any single day's movement.

### One bug worth recording

The first version of `Monthly_Analysis` used `SUMPRODUCT` with boolean multiplication to do the
conditional sums. It returned `#VALUE!` across the sheet.

The cause: `SUMPRODUCT` coerces its arguments to numbers, and the ranges being multiplied
included text columns (karat labels, status flags). A text cell inside a `SUMPRODUCT` argument
is not treated as zero, it fails the whole expression. The fix was to restructure so every
conditional sum runs over a strictly numeric column and to switch to `SUMIFS`, which handles
mixed-type criteria ranges natively and is also considerably faster across 1,675 formulas.

General lesson, and it cost an hour to learn: `SUMPRODUCT` is the wrong tool the moment any
column in the range can contain text.

---

## Running it

```bash
pip install openpyxl
python build_workbook.py
```

Writes `gold_reconciliation_system.xlsx`. The script is the single source of truth; the
workbook is regenerated from it rather than edited by hand.

To verify every formula evaluates:

```bash
python verify.py gold_reconciliation_system.xlsx
```

---

## On the data

**Everything in this repository is synthetic and generated by `build_workbook.py` under a fixed
random seed.** Spot and board rates are illustrative and are not quotes from any exchange.
Karigars are coded `K-01` to `K-05`. SKUs, invoice numbers and the business name are invented.
No real business, counterparty or customer appears anywhere in it.

The four stock variance breaches and K-03's wastage drift are seeded deliberately, so that the
exception logic demonstrates itself rather than sitting idle on clean data. They are marked as
such in the script.

The modelling approach, the fine-gold basis, the per-trade wastage norms and the three daily
questions come from work on physical gold inventory for a retail jeweller. The data does not.

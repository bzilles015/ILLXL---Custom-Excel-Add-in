# ILLXL
### The Illest Excel Add-In for Finance Pros

**50+ keyboard shortcuts. Zero cost. All ILL.**

Tired of paying $125–200/year for WSP Boost, WST Macros, or Macabacus? ILLXL is a free, open-source Excel add-in built specifically for finance professionals — covering everything the paid tools do, plus features they don't.

---

## What's Inside

| Module | What It Does |
|---|---|
| **modCore** | Performance mode, navigate to blanks/errors, break external links, absolute/relative refs |
| **modFormatCycles** | Cycle number, date, currency, percent formats. Adjust decimals. Scale. Toggle sign. |
| **modFormulas** | CAGR, % change, equals-left, growth rate, quick SUM/AVERAGE |
| **modStyles** | Auto-color, font/fill/color cycles, input styles, header styles, indent, zero-check CF |
| **modBorders** | Cycle individual borders, apply outline + inside, sum bars |
| **modUnitTags** | Cycle FAST-standard unit tags ([mln $], [%], [x], [bps], etc.) |

---

## Why ILLXL?

- **Free. Forever.** No subscription, no trial, no upsell.
- **Custom undo engine** — every shortcut supports Ctrl+Z across up to 5,000 cells simultaneously (values, formulas, and number formats).
- **Performance mode** — disables calc, screen updates, events, and CF for blazing-fast bulk formatting. One keystroke on, one keystroke off.
- **FAST-standard unit tags** — cycle through [mln $], [thd $], [bn $], [%], [x], [pp], [bps] and more in two keystrokes. No more typing tags manually.
- **Zero-Check CF** — green if zero (balanced), red if not. Two keystrokes vs a full dialog workflow. Built for balance sheet checks.
- **Auto-Color** — instantly color-code an entire model: blue=hardcode, black=formula, green=sheet ref, red=external link.
- **Center Across Selection** — one keystroke. No more Merge & Center.

---

## Installation

1. Download `ILLXL.xlam` from this repo
2. Open Excel → **File → Options → Add-ins**
3. At the bottom, set **Manage: Excel Add-ins** → click **Go...**
4. Click **Browse**, navigate to `ILLXL.xlam`, click OK
5. Make sure the ILLXL checkbox is checked → **OK**

Shortcuts load automatically every time Excel opens.

---

## Keyboard Shortcuts

📄 **[Download the full shortcut reference PDF](ILLXL_Shortcut_Reference.pdf)**

### modCore — Utilities
| Shortcut | Action |
|---|---|
| Ctrl+Alt+Shift+M | Toggle Performance Mode |
| Ctrl+Alt+Shift+A | Make References Absolute ($A$1) |
| Ctrl+Alt+Shift+R | Make References Relative (A1) |
| Ctrl+Alt+Shift+N | Go To Next Blank in Selection |
| Ctrl+Alt+Shift+E | Go To Next Error in Selection |
| Ctrl+Alt+Shift+L | Break External Links (convert to values) |

### modFormatCycles — Number Formats
| Shortcut | Action |
|---|---|
| Ctrl+Shift+1 | Cycle Number Format (#,##0 → K → M) |
| Ctrl+Shift+3 | Cycle Date Format (m/d/yyyy → m/d/yy → mmm-yy → d-mmm-yy) |
| Ctrl+Shift+4 | Cycle Currency Format ($#,##0 → $K → $M) |
| Ctrl+Shift+5 | Cycle Percent Format (0.0% → 0% → +/- → ...) |
| Ctrl+Shift+8 | Cycle Other Formats (A=Actual / B=Budget / F=Forecast / Q / P / E / x) |
| Ctrl+Shift+> | Increase Decimal Places |
| Ctrl+Shift+< | Decrease Decimal Places |
| Alt+Shift+< | Scale Up ÷1,000 |
| Alt+Shift+> | Scale Down ×1,000 |
| Ctrl+Alt+Shift+\ | Toggle Sign (positive ↔ negative) |
| Ctrl+Alt+2 | Divide by 100 (5 → 0.05) |
| Ctrl+Alt+Shift+2 | Multiply by 100 (0.05 → 5) |

### modFormulas — Formula Tools
| Shortcut | Action |
|---|---|
| Ctrl+Alt+Shift+C | Insert CAGR — n = end year minus start year (2023→2026 = 3) |
| Ctrl+Alt+Shift+W | Insert % Change =(Current−Prior)/ABS(Prior) |
| Ctrl+Alt+D | Equals Left — link each cell to the cell on its left |
| Ctrl+Alt+Shift+G | Apply Growth Rate =LEFT×(1+rate) |
| Ctrl+Alt+= | Quick SUM |
| Ctrl+Alt+Shift+= | Quick Average |

### modStyles — Colors, Fonts & Layout
| Shortcut | Action |
|---|---|
| Ctrl+Alt+A | Auto-Color entire selection |
| Ctrl+' | Cycle Font (Aptos Narrow → Poppins → Times New Roman) |
| Ctrl+Shift+K | Cycle Fill Color |
| Ctrl+Alt+Shift+I | Cycle Text Case (Title → lower → UPPER) |
| Ctrl+Shift+C | Cycle Font Color |
| Ctrl+Shift+F | Increase Font Size |
| Ctrl+Shift+G | Decrease Font Size |
| Ctrl+Alt+> | Indent In |
| Ctrl+Alt+< | Indent Out |
| Ctrl+Alt+E | Center Across Selection (no more Merge & Center) |
| Ctrl+Shift+N | Insert Static Timestamp |
| Ctrl+Alt+Shift+U | Cycle Input Style (Yellow → Lt Yellow → Gray → Peach → Pale Blue) |
| Ctrl+Alt+Shift+H | Cycle Header Style |
| Ctrl+Alt+Shift+Y | Insert Headers from Prompt |
| Ctrl+Alt+Shift+D | Insert Variance Headers (AvF % / AvB % / Var AvB) |
| Ctrl+Alt+Shift+Z | Apply Zero-Check CF |
| Ctrl+Alt+Shift+X | Clear Zero-Check CF |

### modBorders — Border Cycles
| Shortcut | Action |
|---|---|
| Ctrl+Alt+Shift+↑ | Cycle Top Border (Thin → None → Medium → Hairline) |
| Ctrl+Alt+Shift+↓ | Cycle Bottom Border |
| Ctrl+Alt+Shift+← | Cycle Left Border |
| Ctrl+Alt+Shift+→ | Cycle Right Border |
| Ctrl+Alt+Shift+B | Outline (Medium) + Inside (Thin) |
| Ctrl+Alt+− | Apply Sum Bar |
| Ctrl+Alt+_ | Apply Double Sum Bar |

### modUnitTags — Unit Tags
| Shortcut | Action |
|---|---|
| Ctrl+Alt+Shift+T | Cycle Value Tag: [#] [%] [mln $] [thd $] [bn $] [x] [pp] [bps] |
| Ctrl+Alt+Shift+O | Cycle Duration Tag: [d] [m] [q] [y] |
| Ctrl+Alt+Shift+P | Cycle Rate Tag: [%/y] [$/unit] [$/FTE] [$/yr] |
| Ctrl+Alt+Shift+Backspace | Remove Last Unit Tag |

---

## Input Style Reference

| Step | Style | Fill | Font | Use Case |
|---|---|---|---|---|
| 1 | Yellow | #FFF2CC | Blue | Primary hardcoded assumptions |
| 2 | Light Yellow | #FFFFCC | Blue | Secondary / supporting assumptions |
| 3 | Gray | #D9D9D9 | Blue | Linked / locked cells |
| 4 | Peach | #FFC799 | Teal | Special flagged inputs |
| 5 | Pale Blue | #DDEBF7 | Dark Blue | Override cells |

---

## Compared to Paid Alternatives

| Feature | ILLXL | WSP Boost | Macabacus |
|---|---|---|---|
| Price | **Free** | ~$150/yr | ~$200/yr |
| Number format cycles | ✅ | ✅ | ✅ |
| Border cycles | ✅ | ✅ | ✅ |
| Input styles | ✅ | ✅ | ✅ |
| Custom undo engine | ✅ | ❌ | ❌ |
| Performance mode | ✅ | ❌ | ❌ |
| FAST unit tags | ✅ | ❌ | ❌ |
| Zero-check CF | ✅ | ❌ | ❌ |
| Action logging | ✅ | ❌ | ❌ |
| Open source | ✅ | ❌ | ❌ |

---

## Built By

**Bruce Zachary Illes**
[LinkedIn](https://linkedin.com/in/brucezacharyilles) · [GitHub](https://github.com/bzilles015)

---

*ILLXL — because mediocre tools cost too much.*

# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a communications document project for the **Lake Ontario Atlantic Salmon Restoration Program (LOASRP)**, a binational partnership between Ontario MNR and OFAH (Ontario Federation of Anglers and Hunters). The primary deliverable is a print-ready HTML one-pager program brief.

## File Structure

| File | Purpose |
|---|---|
| `Bring Back the Salmon One-Pager.html` | **Current working version** — the main one-pager to edit |
| `uploads/Bring Back the Salmon One-Pager v4.html` | Earlier draft (reference only) |
| `export/Bring-Back-the-Salmon-One-Pager.html` | Bundled export (regenerated, don't hand-edit) |
| `export/src/Bring Back the Salmon One-Pager.html` | Source copy used for export |

## Source Documents

| File | Contents |
|---|---|
| `uploads/_ATS_workplan.txt` | MNR 5-Year Workplan 2025–2029 (plain text extract) — authoritative program priorities and actions |
| `uploads/_LOTC_status.txt` | LOTC Atlantic Salmon Working Group status paper, March 2026 — science, stocking, and assessment data |
| `uploads/FINAL_DRAFT_ATS_Workplan.docx` | Original Word version of the workplan |
| `uploads/LOTC_Status of Atlantic Salmon Restoration*.docx` | Original Word version of the status paper |

When updating statistics, program priorities, or factual claims in the HTML, verify against these source documents first.

## HTML Document Architecture

The one-pager is a self-contained HTML file with all CSS inlined in `<style>` tags. It renders as a letter-portrait page (816 × 1056 px) with a `@media print` rule for direct browser printing.

**Layout sections (top to bottom):**
1. `.masthead` — logos (OFAH + Fish & Wildlife) and program title
2. `.hero` — full-bleed photo with overlaid headline
3. `.intro` — two-column "Our Goal / Why This Matters" grid
4. `.priorities` — five-column numbered priority cards
5. `.lake` — two-column layout: copy + `.report-cta` | `.sidebar` KPI stats
6. `.doing` — four-column action cards (Fish Production, Habitat, Science, Education)
7. `.foot` — dark navy footer with website CTA

**CSS custom properties (color palette):**
```css
--navy: #0b3a53        /* deep navy — headers, footer, card icons */
--forest: #2e5b3a      /* conservation green — section bars, kickers */
--sand: #ede6d2        /* warm cream — sidebar, CTA backgrounds */
--accent: #b7651f      /* copper/amber — CTA border, superscripts */
--ink: #0f1d24         /* near-black body text */
--mute: #607078        /* secondary/caption text */
```

## Assets

All images are in `assets/` and referenced with relative paths:
- `assets/ofah-logo.png` — OFAH logo (square, 44×44px in masthead)
- `assets/fish-wildlife-logo.png` — Ontario Fish and Wildlife logo
- `assets/hero-angler.jpg` — hero section background photo
- `assets/salmon-icon-white.png` — small salmon silhouette (white, used in cards and CTA)

The `uploads/` folder contains additional reference images (electrofishing photos, alternate logos) not currently used in the one-pager.

## Key Program Facts (for content accuracy)

- **3M+** Atlantic Salmon stocked since 2006
- **300+** returning adults sampled in the past three years
- **150+** habitat projects delivered
- **1,500+** students reached through the Classroom Hatchery
- **50+** partner organizations
- Five priorities: increase adult abundance, improve habitat, enhanced yearling stocking, enhance angling opportunities, improve communications
- Focus area: tributaries **east of Toronto** + Lakefront Promenade (Credit River restoration paused)
- Website: **bringbackthesalmon.ca** | Contact: **info@bringbackthesalmon.ca**
- Partners: Ontario MNR · OFAH · NYSDEC · USFWS · DFO

## Previewing Changes

Open `Bring Back the Salmon One-Pager.html` directly in a browser. No build step required. For print output, use browser Print → Save as PDF (letter, no margins).

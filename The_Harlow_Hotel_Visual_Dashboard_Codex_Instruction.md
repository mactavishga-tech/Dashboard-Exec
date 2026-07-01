# THE HARLOW HOTEL
# Final Visual Excel Dashboard Build Specification

## Codex Instruction — Phase 1 Visual Prototype Only

Create a production-ready Excel dashboard visual prototype for **The Harlow Hotel, Harlow, Essex**.

This build is **visual-first only**. It must reproduce the reference dashboard’s appearance, hierarchy, density, colours, spacing, card layout, chart style, and executive reporting feel as closely as possible, while applying the Harlow Hotel amendments listed below.

Do **not** build the full Power Query architecture, data model, governance framework, or refresh automation at this stage. Those belong to a later phase.

The purpose of this workbook is to provide a polished Excel design reference that can be reviewed visually and then used as the front-end prototype for a future Power BI implementation.

---

## 1. Core Visual Amendments

Apply these amendments throughout the workbook:

- Replace all previous/legacy branding with:
  - **THE HARLOW HOTEL**
  - **Harlow, Essex**
- Do not use any legacy hotel name, legacy city, legacy country, or India-specific contact details
- Use placeholder Harlow Hotel contact details:
  - `+44 Placeholder`
  - `info@harlowhotel.com`
  - `Harlow, Essex, UK`
- Use **GBP (£)** throughout.
- Do not use any non-GBP currency symbols, non-GBP currency codes, or South Asian large-number abbreviations
- Replace all room types with **Twin** and **Double only**.
- Do not use any legacy four-category room mix
- Replace the legacy guest-satisfaction metric with **Review Performance Score (RPS)**.
- Add the following visual sections not present in the original dashboard:
  - Daily Operations Strip
  - Executive Alerts
  - Data Health

---

## 2. Deliverable

Create one Excel workbook:

`The_Harlow_Hotel_Executive_Dashboard_Visual_Prototype.xlsm`

The workbook must be macro-enabled because the later implementation may use macros for refresh/navigation support. For this visual phase, include only lightweight macros if useful, such as navigation helpers or reset-filter placeholders.

The workbook must contain:

- Complete dashboard visual prototype
- The Harlow Hotel branding
- Dark navy luxury sidebar
- Gold accent styling
- GBP currency formatting
- Twin/Double room types only
- Six KPI cards
- Daily Operations Strip
- Analytics grid with charts and tables
- Review Performance Score panel
- Housekeeping Status panel
- Executive Alerts panel
- Data Health widget
- Bottom Key Insights bar
- Placeholder data stored in named Excel tables
- Simple navigation hyperlinks between sheets

---

## 3. Required Workbook Sheets

Create the following sheets:

1. `01_Dashboard`
2. `02_Placeholder_Data`
3. `03_Settings`

Optional support sheets may be created if needed:

4. `04_Helper`
5. `05_Pivots`

Keep the workbook simple. This is not the full data architecture stage.

---

## 4. Overall Visual Style

Create a premium hospitality operations dashboard with:

- Dark navy branded sidebar
- Gold luxury accent colour
- White and light-grey analytics workspace
- Rounded white KPI and chart cards
- Dense executive BI layout
- Compact spacing
- Minimal whitespace
- Clean professional chart styling
- Subtle shadows only
- Clear executive information hierarchy

The dashboard should feel like a **luxury hotel management dashboard**, not a marketing landing page.

Avoid:

- Bright gradients in the main canvas
- Oversized visuals
- Thick borders
- Heavy shadows
- 3D charts
- Large decorative elements
- Purple/blue AI-style gradients
- SaaS-style pill overuse
- Backend architecture content

---

## 5. Canvas and Layout

Use a **16:9 wide dashboard canvas**.

Target size:

- Approximately `1536 × 864 px`

Primary structure:

- Left sidebar: approximately `205 px` wide
- Main dashboard area: remaining width
- Header/filter bar: approximately `85 px` high
- KPI row: six cards
- Daily Operations Strip
- Main analytics grid
- Executive Alerts/Data Health section
- Bottom Key Insights bar

Background colours:

- Body background: `#f5f6fa`
- Main canvas: `#f7f8fb`
- Sidebar: `#061936` to `#031126`
- Card background: `#ffffff`
- Card border: `#e7eaf0`

Card styling:

- Border radius: `10–14 px`
- Border: `1 px solid #e7eaf0`
- Subtle shadow only:
  - `0 4px 14px rgba(15, 23, 42, 0.06)`

---

## 6. Colour Palette

Use this palette consistently:

| Purpose | Colour |
|---|---:|
| Deep navy | `#061936` |
| Darker navy | `#031126` |
| Navy blue | `#092b66` |
| Gold | `#d4a84f` |
| Bright gold | `#f2c261` |
| Amber | `#d8a22d` |
| Green | `#16a05d` |
| Soft green | `#eaf8ef` |
| Purple | `#8b4fd6` |
| Rose | `#d94b63` |
| Cyan | `#5cc7ad` |
| Canvas | `#f5f6fa` |
| Card | `#ffffff` |
| Border | `#e7eaf0` |
| Gridline | `#edf0f5` |
| Primary text | `#071326` |
| Secondary text | `#4b5563` |
| Muted text | `#7b8494` |

Dominant visual language:

- Navy
- Gold
- White
- Soft grey

---

## 7. Typography

Use a clean modern sans-serif font.

Preferred:

- Manrope

Fallbacks:

- Inter
- Segoe UI
- Calibri

Font weights:

- Dashboard title: `800`
- Card title: `700–800`
- KPI values: `800`
- Table headers: `700`
- Body text: `500`

Text colours:

- Primary: `#071326`
- Secondary: `#4b5563`
- Muted: `#7b8494`

---

## 8. Left Sidebar

Create a full-height deep navy vertical sidebar.

Approximate width:

- `205 px`

Background:

- Dark navy gradient from `#061936` to `#031126`

### 8.1 Branding Block

Position at top of sidebar.

Height:

- Approximately `125 px`

Content:

- Gold luxury mark, crest, crown, wreath, or monogram
- Large gold hotel initial/monogram
- `THE HARLOW HOTEL`
- `Harlow, Essex`

Alignment:

- Horizontally centred

### 8.2 Navigation

Navigation starts around `140 px` from top.

Navigation items:

1. Overview
2. Reservations
3. Guests
4. Rooms
5. Revenue
6. F&B Sales
7. Housekeeping
8. Reviews
9. Reports

Each navigation item:

- Height: approximately `46 px`
- Horizontal padding: `16 px`
- Margin: `4 px` vertical and `15 px` horizontal
- Border radius: `7 px`
- Font size: `13 px`
- Font weight: `500`
- Icon + text layout
- Gap between icon and text: approximately `12 px`

Active item:

- `Overview`

Active style:

- Gold gradient background
- Text colour: `#071326`
- Font weight: `700`

Inactive style:

- White text
- Opacity approximately `92%`

### 8.3 Sidebar Contact Card

Position:

- Bottom of sidebar

Style:

- Left/right margin: `15 px`
- Bottom margin: `18 px`
- Border: `1 px solid rgba(255,255,255,0.12)`
- Border radius: `8 px`
- Padding: `16 px`
- Background: `rgba(255,255,255,0.03)`

Contact text:

- `THE HARLOW HOTEL`
- `+44 Placeholder`
- `info@harlowhotel.com`
- `Harlow, Essex, UK`

---

## 9. Header Area

Main content begins to the right of the sidebar.

Header padding:

- Top: `20 px`
- Left/right: `20 px`
- Bottom: `0 px`

Header layout:

- Left: title and subtitle
- Right: filters, last updated block, refresh icon

Title:

- `Dashboard Overview`

Title style:

- Font size: `23 px`
- Font weight: `800`
- Colour: `#071326`

Subtitle:

- `Real-time Performance Snapshot`

Subtitle style:

- Font size: `13 px`
- Colour: `#4b5563`

### 9.1 Filter Boxes

Required filters:

- Date Range
- Department
- Room Type
- Booking Source

Optional if space allows:

- Market Segment
- Rate Code

Filter box style:

- Height: `46 px`
- Background: `#ffffff`
- Border: `1 px solid #e4e7ec`
- Border radius: `8 px`
- Padding: `8 px 14 px`
- Minimum width: approximately `150–190 px`

Filter label:

- Font size: `10 px`
- Font weight: `700`
- Colour: `#111827`

Filter value:

- Font size: `11 px`
- Colour: `#111827`

Example Date Range:

- `1 Apr 2026 - 31 May 2026`

Room Type filter values:

- `All`
- `Twin`
- `Double`

### 9.2 Last Updated and Refresh

Last updated block:

- Label: `Last Updated`
- Value: `1 Jun 2026 08:30 AM`

Refresh icon:

- Far right
- Navy outline style
- Clean circular or square icon button

---

## 10. KPI Card Row

Create six KPI cards in one row.

Grid:

- Six equal columns
- Gap: `10 px`

Each card:

- Height: approximately `142 px`
- Background: `#ffffff`
- Border: `1 px solid #e7eaf0`
- Border radius: `10 px`
- Padding: `15 px 17 px`
- Subtle shadow only

Each KPI card contains:

- Label top-left
- Circular icon top-right
- Large value
- Green percentage change
- Comparison period text
- Small sparkline at bottom

KPI cards:

| KPI | Value | Change | Accent |
|---|---:|---:|---:|
| Total Bookings | `1,652` | `+18.7%` | `#1f70c1` |
| Total Guests | `3,142` | `+16.4%` | `#1a9858` |
| Occupancy Rate | `76.8%` | `+8.9%` | `#8b4fd6` |
| Average Daily Rate | `£92.40` | `+12.3%` | `#e2a72e` |
| RevPAR | `£70.96` | `+22.6%` | `#d94b63` |
| Total Revenue | `£148K` | `+21.9%` | `#2476c7` |

Comparison text:

- `vs 1 Mar - 31 Mar 2026`

KPI label style:

- Font size: `12 px`
- Font weight: `700`
- Colour: `#111827`

KPI value style:

- Font size: `24 px`
- Font weight: `800`
- Colour: `#071326`

Uplift style:

- Colour: `#16a05d`
- Font size: `11 px`
- Font weight: `700`

Comparison style:

- Font size: `9 px`
- Colour: `#111827`

Icon circles:

- Use pale tint of the KPI accent colour
- Thin outline icons

Sparklines:

- Tiny line charts
- Matching KPI accent colour
- No axes
- No gridlines

---

## 11. Daily Operations Strip

Add a compact operational strip directly below the KPI cards.

It should appear as a row of smaller operational cards and must not dominate the page.

Metrics:

| Metric | Value |
|---|---:|
| Occupied Rooms | `91` |
| Available Rooms | `119` |
| Arrivals | `38` |
| Departures | `42` |
| Stayovers | `49` |
| Out of Order | `3` |
| Day Use | `2` |
| Current Occupancy | `76.8%` |

Visual style:

- White cards
- Small navy/gold icons
- Compact text
- Border: `#e7eaf0`
- Border radius: `8 px`
- Height: approximately `64–76 px`
- Tight spacing

---

## 12. Main Analytics Grid

All analytics cards:

- White background
- Light grey border
- Border radius: `10 px`
- Padding: `15 px`
- Subtle shadow only

Card title style:

- Font size: `13 px`
- Font weight: `800`
- Colour: `#071326`

Card gaps:

- `8–10 px`

---

## 13. Analytics Row 1

### 13.1 Occupancy Rate Trend

Approximate width:

- `37%`

Approximate height:

- `230–250 px`

Title:

- `Occupancy Rate Trend`

Chart type:

- Line chart

X-axis labels:

- `1 Apr`
- `8 Apr`
- `15 Apr`
- `22 Apr`
- `29 Apr`
- `6 May`
- `13 May`
- `20 May`
- `27 May`
- `31 May`

Y-axis:

- `0%`
- `25%`
- `50%`
- `75%`
- `100%`

Series:

- This Period: solid dark navy line with circular markers
- Last Period: grey-blue dashed line

Colours:

- This Period: `#092b66`
- Last Period: `#8a9bb3`
- Gridlines: `#edf0f5`

Legend:

- Small, clean, approximately `9 px`

### 13.2 Room Type Wise Bookings

Approximate width:

- `34%`

Approximate height:

- `230–250 px`

Title:

- `Room Type Wise Bookings`

Chart type:

- Donut chart

Segments:

| Room Type | Count | Share | Colour |
|---|---:|---:|---:|
| Twin | `742` | `44.9%` | `#092f78` |
| Double | `910` | `55.1%` | `#f2b22e` |

Centre text:

- `1,652`
- `Total Bookings`

Donut hole:

- `58–65%`

Legend:

- Right side
- Clean labels
- Show count and percentage

### 13.3 Revenue Summary

Approximate width:

- `29%`

Approximate height:

- `230–250 px`

Title:

- `Revenue Summary`

Top-right dropdown text:

- `This Period`

Rows:

| Revenue Line | Value | Change |
|---|---:|---:|
| Room Revenue | `£109.20K` | `↑ 20.2%` |
| F&B Revenue | `£27.40K` | `↑ 19.8%` |
| Other Services | `£11.40K` | `↑ 18.6%` |
| Total Revenue | `£148.00K` | `↑ 21.9%` |

Row height:

- Approximately `45 px`

Total row:

- Background: `#f7f8fb`
- Font weight: `800`

---

## 14. Analytics Row 2

### 14.1 Revenue Trend

Approximate width:

- `34%`

Approximate height:

- `220–235 px`

Title:

- `Revenue Trend`

Chart type:

- Combo chart

Bars:

- Revenue (`£`)
- Colour: `#092f78`

Line:

- ADR (`£`)
- Colour: `#d8a22d`
- Stroke: `2 px`
- Small circular markers

Left y-axis:

- `£0`
- `£2K`
- `£4K`
- `£6K`
- `£8K`

Right y-axis:

- `£0`
- `£25`
- `£50`
- `£75`
- `£100`
- `£125`
- `£150`

### 14.2 Bookings by Source

Approximate width:

- `19%`

Approximate height:

- `220–235 px`

Title:

- `Bookings by Source`

Chart type:

- Donut chart

Segments:

| Source | Bookings | Share | Colour |
|---|---:|---:|---:|
| Direct Website | `612` | `37.0%` | `#092f78` |
| OTA | `528` | `31.9%` | `#f2b22e` |
| Walk-in | `276` | `16.7%` | `#5cc7ad` |
| Corporate | `168` | `10.2%` | `#6aa12b` |
| Others | `68` | `4.1%` | `#a14bdc` |

Legend:

- Compact
- Clean
- No clutter

### 14.3 Top Booking Sources

Approximate width:

- `24%`

Approximate height:

- `220–235 px`

Title:

- `Top Booking Sources`

Table columns:

- Source
- Bookings
- % Share

Rows:

| Source | Bookings | % Share |
|---|---:|---:|
| Direct Website | `612` | `37.0%` |
| OTA Online | `528` | `31.9%` |
| Walk-in | `276` | `16.7%` |
| Corporate | `168` | `10.2%` |
| Others | `68` | `4.1%` |
| Total | `1,652` | `100%` |

Table style:

- Header font size: `10 px`
- Body font size: `11 px`
- Row bottom border: `#edf0f5`
- Total row bold

### 14.4 Top Performing Room Types

Approximate width:

- `23%`

Approximate height:

- `220–235 px`

Title:

- `Top Performing Room Types`

Table columns:

- Room Type
- Occupancy %
- ADR (£)
- RevPAR (£)

Rows:

| Room Type | Occupancy % | ADR (£) | RevPAR (£) |
|---|---:|---:|---:|
| Twin | `74.5%` | `£88.20` | `£65.71` |
| Double | `78.6%` | `£95.80` | `£75.30` |
| Total / Avg | `76.8%` | `£92.40` | `£70.96` |

Do not include any other room types.

---

## 15. Analytics Row 3

### 15.1 Monthly Performance Overview

Approximate width:

- `36%`

Approximate height:

- `190–210 px`

Title:

- `Monthly Performance Overview`

Table columns:

- Month
- Bookings
- Guests
- Occupancy %
- ADR (£)
- RevPAR (£)
- Revenue (£)

Rows:

| Month | Bookings | Guests | Occupancy % | ADR (£) | RevPAR (£) | Revenue (£) |
|---|---:|---:|---:|---:|---:|---:|
| Apr 2026 | `812` | `1,542` | `74.1%` | `£89.40` | `£66.25` | `£70.21K` |
| May 2026 | `840` | `1,600` | `79.4%` | `£95.40` | `£75.75` | `£77.79K` |
| Total / Avg | `1,652` | `3,142` | `76.8%` | `£92.40` | `£70.96` | `£148.00K` |

Total row:

- Background: `#f7f8fb`
- Font weight: `800`

### 15.2 Average Daily Rate Trend

Approximate width:

- `27%`

Approximate height:

- `190–210 px`

Title:

- `Average Daily Rate (ADR) Trend`

Chart type:

- Line chart only

Line colour:

- `#d8a22d`

Marker colour:

- `#d8a22d`

Y-axis:

- `£50`
- `£75`
- `£100`
- `£125`
- `£150`

### 15.3 Review Performance Score

Approximate width:

- `16%`

Approximate height:

- `190–210 px`

Title:

- `Review Performance Score`

Visual type:

- Gauge-style panel, progress ring, or compact score block

Preferred:

- Semi-circle gauge or progress ring

Display:

- `82%`
- `RPS`

Subtext:

- `Target: 80%`

Bottom text:

- `↑ 6.5% vs 1 Mar - 31 Mar 2026`

Gauge colours:

- Filled: `#d8a22d`
- Remaining: `#d9dce4`
- Value text: `#071326`
- Uplift: `#16a05d`

### 15.4 Housekeeping Status

Approximate width:

- `21%`

Approximate height:

- `190–210 px`

Title:

- `Housekeeping Status`

Four mini status cards in a `2 × 2` grid:

| Status | Value | Background | Text/Icon |
|---|---:|---:|---:|
| Cleaned | `91` | `#eaf8ef` | `#16a05d` |
| In Progress | `14` | `#fff7e6` | `#d8a22d` |
| Pending | `11` | `#fdecec` | `#d94b63` |
| Out of Order | `3` | `#f1f2f6` | `#374151` |

Footer text:

- `Total Rooms: 119`

---

## 16. Executive Alerts Panel

Add a compact Executive Alerts card.

Position:

- Bottom-right lower dashboard area, near Data Health and above the Key Insights bar.

Title:

- `Executive Alerts`

Show 4–6 compact alert rows:

- Revenue below budget
- ADR below target
- Occupancy below target
- High OOO rooms
- Housekeeping backlog
- RPS decline

Each alert row:

- Small status dot
- Alert text
- RAG colour

RAG colours:

- Green: `#16a05d`
- Amber: `#d8a22d`
- Red: `#d94b63`

Style:

- White card
- Light border
- Compact rows
- Not overly large

---

## 17. Data Health Widget

Add a compact Data Health card.

Position:

- Near Executive Alerts in the bottom-right area above the Key Insights bar.

Title:

- `Data Health`

Display:

| Item | Value |
|---|---:|
| Last Refresh | `1 Jun 2026 08:30` |
| Records Processed | `4,862` |
| Failed Records | `0` |
| Missing Revenue | `0` |
| Missing Bookings | `0` |
| Refresh Status | `Green` |

Visual style:

- Small operational control card
- RAG status indicator
- White background
- Light border
- Compact text
- Not visually dominant

---

## 18. Bottom Key Insights Bar

Create a full-width dark navy bar at the bottom of the main dashboard area.

Height:

- Approximately `56 px`

Background:

- `#061936`

Text colour:

- `#ffffff`

Layout:

- Left label
- Insight items across the bar
- Thin vertical dividers between insight items
- Right note

Left label:

- `Key Insights`

Insight items:

- `Occupancy rate improved by 8.9% compared to last period`
- `Revenue increased by 21.9% compared to last period`
- `ADR increased by 12.3% compared to last period`
- `Direct bookings contributed 37.0% of total bookings`

Highlight percentages in:

- `#7bd88f`

Right note:

- `All values are in GBP (£)`

---

## 19. Icon Style

Use thin outline icons.

Preferred style:

- Lucide
- Feather
- Font Awesome outline

Stroke width equivalent:

- `1.8 px`

Suggested icons:

- calendar-check
- users
- bed
- tag
- chart-line
- pound-sign
- calendar-days
- door-open
- utensils
- broom
- star
- file-chart
- refresh-cw

Sidebar icons:

- White

KPI icons:

- Inside circular pastel backgrounds
- Match KPI accent palette

---

## 20. Chart Styling

General chart styling:

- Gridline colour: `#edf0f5`
- Axis label colour: `#4b5563`
- Axis label font size: `9 px`
- Legend font size: `9 px`
- No heavy chart borders
- Charts should feel embedded inside cards

Line charts:

- Stroke width: `2 px`
- Marker radius: `2.5–3 px`

Bars:

- Fill: `#092f78`
- Rounded top corners if possible

Donuts:

- Cutout/hole: `58–65%`
- Segment border: white `1 px`

Avoid:

- 3D charts
- Heavy chart outlines
- Excess labels
- Large legends
- Bright or clashing colours

---

## 21. Placeholder Data Requirements

Although this is a visual prototype, all visible values must come from placeholder data tables rather than being manually hardcoded into dashboard display cells.

Create named Excel tables on `02_Placeholder_Data`:

- `tbl_KPI_Summary`
- `tbl_Daily_Operations`
- `tbl_Occupancy_Trend`
- `tbl_Room_Type_Distribution`
- `tbl_Revenue_Summary`
- `tbl_Revenue_Trend`
- `tbl_Booking_Sources`
- `tbl_Top_Booking_Sources`
- `tbl_Room_Performance`
- `tbl_Monthly_Performance`
- `tbl_ADR_Trend`
- `tbl_RPS_Summary`
- `tbl_Housekeeping_Status`
- `tbl_Executive_Alerts`
- `tbl_Data_Health`
- `tbl_Insights`

Data table rules:

- Use clean headers
- Use consistent field names
- Use proper numeric/date data types where possible
- Do not merge cells in data tables
- Keep source tables query-ready for later Power Query replacement

---

## 22. Navigation and Interactivity

Create light navigation support using Excel-native functionality:

- Sidebar navigation items should hyperlink to relevant sheets or placeholder sections.
- `Overview` should link to `01_Dashboard`.
- Other nav items may link to placeholder anchors or supporting sheets if those pages are not yet built.
- Add a simple refresh/reset visual button in the header if practical.
- If macros are added, keep them simple and documented in `03_Settings`.

This phase should support visual review and future Power BI multi-page mapping, not full operational automation.

---

## 23. Visual Priority Order

Build in this priority:

1. Dark navy branded sidebar
2. Gold luxury hotel identity
3. Header with filters and last updated area
4. Six compact KPI cards
5. Daily Operations Strip
6. Analytics Row 1
7. Analytics Row 2
8. Analytics Row 3
9. Review Performance Score using the Harlow reporting metric
10. Executive Alerts card
11. Data Health widget
12. Bottom navy Key Insights bar
13. Named placeholder data tables
14. Hyperlinks/navigation polish

---

## 24. Acceptance Criteria

The final workbook is acceptable when:

- The dashboard visually matches the supplied reference design style.
- The brand reads clearly as **The Harlow Hotel, Harlow, Essex**.
- The sidebar is dark navy with gold luxury styling.
- The main canvas uses white/light-grey executive BI cards.
- There are six KPI cards across the top.
- All values use GBP.
- Only Twin and Double appear as room types.
- RPS appears as the review-quality metric.
- The Daily Operations Strip is present.
- Executive Alerts are present.
- Data Health is present.
- The Key Insights bar is present at the bottom.
- The workbook is saved as `.xlsm`.
- Placeholder data is stored in named Excel tables.
- No full Power Query/data model architecture has been added yet.

---

## 25. Explicit Exclusions for Phase 1

Do not include:

- Full Power Query architecture
- Fact/dimension modelling
- Source system mapping
- KPI governance framework
- Refresh automation design
- Data quality rules beyond visual Data Health placeholders
- Enterprise deployment notes
- Security/RLS design
- Detailed DAX measures
- Power BI semantic model documentation

Those items belong to Phase 2 after the visual prototype is reviewed.

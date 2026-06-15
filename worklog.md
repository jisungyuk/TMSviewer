# TMSviewer — Work Log

---

## 2026-06-15

### ActiveWindow — MVC / Hold Task popup (realtime_viewer.py)

#### Layout & UI
- Added "Maximum Voluntary Contraction" title at top; switches to "Hold Task" when Hold task is checked
- 3-column target row: left inputs | "targets" label | right inputs (each column center-aligned)
- Horizontal divider line between avg stats and Hold task section
- Y axis label parentheses removed (`<-` / `->`)
- LabChart status "RUNNING" → "PLAY"

#### Stats & History
- n counter starts from 1 during first streaming trial (was showing 0)
- avg MAX / MIN shown as clickable buttons → copies value; adjacent editable QLineEdit for manual override
- Per-side Redo buttons; Redo blocked while Hold task is active

#### Switch mode
- Normal click: toggles left-only ↔ right-only only
- Shift+click: both-channel mode (`<--->`)
- Button tooltip explains the two modes
- Status/countdown text only shown on active side(s)

#### Y axis
- Per-side Y axis spinbox (hidden/shown per switch mode)
- WAIT/GO/RELAX text repositioned proportionally when Y max changes

#### Hold Task mode
- Checkbox enabled only when active side(s) have avg MAX measured
- On check: title → "Hold Task", MVC button → "HOLD", channels/Y axis/Redo disabled
- Y axis auto-set to 0–100 (%MVC); ball position converted to (raw / avg_MAX × 100)
- Auto-applies default targets (50% ± 10%) and sets duration to 0 (infinite)
- On uncheck: everything restored, target bands cleared
- Switch blocked with warning if target side lacks MVC data

#### Target bands
- Per-side: `[50 %] ± [10 %] [APPLY]` — no spinbox arrows, centered
- APPLY draws gray semi-transparent horizontal band + thick white center line on graph (0–100 axis)
- Ball turns green inside band, red outside (during HOLD streaming)
- Bands cleared on Hold task uncheck

#### Streaming behavior
- Duration 0 = infinite (no countdown shown); range 0–999 s; default 10 s
- HOLD mode: no stats commit on stop, no "GO" instruction shown
- F1 shortcut: MVC/HOLD/Stop button; F2 shortcut: Switch button
- Ball size increased 22 → 26 px

---

## 2026-06-05

### Window & Layout
- Fixed window size to **2000 × 1125 px** (`setFixedWidth` + `setFixedHeight`); removed `adjustSize()`
- White background for entire central widget
- Black 2px border on both plot widgets
- Horizontal separator lines between: Hunt row / MSO row, and MSO row / Tables row

### Mode & Analysis Controls
- **Chart / Scope** mode radio buttons (top row)
  - Chart mode: shows Window spinbox + Apply
  - Scope mode: shows Pre / Post spinbox + Apply (defaults: 0.2 s / 0.8 s)
- **None / MEP** analysis radio buttons (second row)
  - Disabled in Chart mode
  - Default: MEP selected when switching to Scope

### Centre Column (between graphs)
- **Trial #** large spinbox (28 pt bold), range 0–9999, user-editable; auto-increments on Sample
- **`<--->`** arrow label (14 pt bold) above TMS param set 1
- **TMS parameter set 1** (`_make_mep_params_widget`):
  - MEP window: start — end (ms)
  - Threshold: QDoubleSpinBox, default 0.05 mV
  - Prestim window: start — end (ms)
  - Spinbox pairs: first box has no unit suffix, second has ` ms`
- **Reset** button: resets both param sets to defaults (10/50/0.05/−200/−50)
- **Extend / Collapse** toggle button:
  - Expands a second TMS param set below (for right graph)
  - Arrow changes: `<--->` → `<---`, and `--->`  appears above set 2
- **TMS parameter set 2** (hidden until Extend): same layout as set 1

### Graphs
- Left EMG (blue) and Right EMG (red), side by side
- Stats labels (MEP amp / Prestim RMS) flanking each graph (9 pt, 90 px wide)
- Y-axis label: `mV`
- **Scope mode shading**: yellow semi-transparent (`rgba(255,220,0,60)`) regions for Prestim and MEP windows
  - Left graph always uses param set 1
  - Right graph uses set 2 if Extend is active, otherwise set 1
  - Shading not shown in Chart mode
- Orange dashed vertical line at trigger (t = 0) in scope mode
- Orange dashed vertical lines at each trigger time in chart mode

### Hunt Panel (Scope + MEP mode only)
- One panel per graph (left / right), placed below graphs
- **Hunt** checkbox enables: `num / den` labels + **Clear** / **Redo** buttons (12 pt)
- Denominator increments on each Sample press (when Hunt active)
- Numerator increments after scope data confirms MEP amp > threshold
- When both panels active, Clear/Redo operate on both sides simultaneously
- History stack supports Redo (restores previous num/den)
- Panels disabled in Chart mode or when analysis = None

### MSO / Location Row
- Large font (23 pt): **MSO %:** spinbox + **Location ID:** spinbox
- Both spinboxes: wheel-scrollable without focus, no arrow buttons, click → select all (`_WheelSpinBox`)
- **New Best** / **Equal Best** buttons (16 pt):
  - New Best: adds entry to dropdown in red
  - Equal Best: adds entry in orange (`#e65100`)
  - Format: `{MSO}%at{LocID}`
- **Dropdown** (QComboBox, 16 pt, min 200 px): shows history of best entries with colours

### Data Tables
- One `QTableWidget` per graph, fixed height 180 px, scrollable
- Columns: `Channel | Loc ID | %MSO | Trial | MEP amp | Prestim RMS`
- Row added on each Sample press: Channel/Loc/MSO/Trial filled immediately; MEP amp/RMS filled after scope data is ready
- **Save** button between the two tables: opens file dialog (default: Desktop), saves `.txt` with tab-separated data for both tables

### Play / Sample Buttons
- **Play / Stop** (▶ / ⏹, 13 pt, 48 px tall): starts/stops LabChart recording
- **Sample** (green when enabled, 13 pt, 48 px tall): logs trigger timestamp, adds comment to LabChart, increments Trial #
- Auto-start viewer when LabChart begins recording externally

### LabChart Integration (`labchart_client.py`)
- COM API: `GetActiveObject("ADIChart.Application")`
- `_active_rec()`: caches last valid `SamplingRecord` so data retrieval works after sampling stops
- `get_scope_data()`: fetches fixed window around trigger timestamp
- `get_latest_data()`: fetches last N seconds for chart mode
- Fix: stop streaming → navigating to past triggers now works correctly (user-selected bypass + `_active_rec` fallback)

### Styling
- All buttons: gray (`#a0a0a0`) with hover (`#909090`), pressed (`#787878`), disabled (`#d0d0d0`) states
  - Exceptions: Sample (green `#2e7d32` when enabled), Play (inherits gray, changes text only)
- Default (unsized) fonts scaled up by 10% via `central.setFont(scaled_app_font)`
  - Explicitly sized widgets (Trial #, MSO, arrows, Hunt, stats, etc.) retain their own sizes

# TMSviewer — Work Log

---

## 2026-06-16

### LabChart FRO + TriggerBox 연결 테스트 (test_fro_pulse.py)

#### 셋업
- TMS: Magstim BiStim (Independent triggering mode)
  - PowerLab Output 1 → Conditioning machine 트리거
  - PowerLab Output 2 → Testing machine 트리거
  - Conditioning machine feedback → ch6
- TriggerBox: Brain Products (COM5, 115200 baud), ch1 → PowerLab Input 5 (analog)
- FRO 트리거 조건: Input 5 above 3.5V (hex 내부: `Input = 4`, `Level = 35`, 0-indexed)

#### 검증 완료
- Python → `win32com.GetActiveObject("ADIChart.Application")` → `PlayMessage(hex)` → FRO 설정
- `pyserial` COM5 → TriggerBox `0x01` → PowerLab Input 5 → FRO 발화
- DoublePulse / SinglePulse 템플릿 발화 확인
- ISI 동적 변경 (`set_fro_output_delay`) 확인

#### 주요 발견 및 수정

**COM 연결**
- `CreateObject("ADIChart.Document")` → 새 문서 생성 (잘못된 방식) → `.vbs` 파일을 `GetObject` 방식으로 수정
- COM 객체 캐시 시 `RPC_E_DISCONNECTED (-2147417848)` 발생 → `play_message()` 내부에서 매번 fresh 연결하도록 수정

**FRO hex 구조**
- PlayMessage hex = 바이너리 헤더 + UTF-16LE 인코딩 FRO 설정 텍스트
- ISI 변경: `PulseDelay` 값 바이트 in-place 치환 (6자 `X.XXXX` 고정 길이)
- byte 20–23: 체크섬, 치환 후 byte sum delta만큼 업데이트

**트러블슈팅**
- SinglePulse 발화 안 됨 + ch6 신호 없음 → ch6 = Conditioning 피드백만 수신 (Testing 피드백 미연결), 별개로 Output 2 케이블 불량
- DoublePulse에서 클릭 한 번만 들림 → Output 2 케이블 불량 (Conditioning만 발화 중이었음)
- **근본 원인: MAGIC 셋업 케이블이 Output 2 → Testing machine 연결을 방해** → MAGIC 제거 후 정상 동작 확인

---

## 2026-06-17

### FRO 모드 (Paired-Pulse) 안정화 및 그래프 정렬 수정 (realtime_viewer.py)

#### 해결한 문제

**1. PowerLab FRO 초기화 실패 (에러/발화 안 됨)**

- **증상**: 첫 번째 Play 후 SP/DP 버튼 클릭 시 TMS 발화 안 됨. Stop → Play 후에는 정상 동작. 이후 SP↔DP 전환 시 disconnect 에러 발생.
- **근본 원인**: PowerLab FRO는 첫 `StartSampling` 직후 내부 초기화가 완료되지 않은 상태임. PlayMessage를 받아도 첫 번째 StartSampling 사이클에서는 제대로 반영되지 않는 하드웨어 동작.
- **오해했던 원인**: 초기에 COM background thread의 `PumpWaitingMessages`가 PlayMessage와 충돌한다고 의심했으나, COM thread를 꺼도 동일한 에러 발생 → COM thread는 원인이 아니었음.
- **해결책: Preload + Auto-bounce** (`_fro_preload` → `_fro_preload_dp` → `_fro_preload_done` chain):
  1. Play 누름 → StartSampling 전에 SP PlayMessage 전송 (300ms 대기) → DP PlayMessage 전송
  2. StartSampling #1 실행
  3. 500ms 대기 → 자동 StopSampling (`_fro_auto_bounce`)
  4. SP PlayMessage 재전송 → DP PlayMessage 재전송
  5. StartSampling #2 → 버튼 활성화, data_timer 시작
  - 총 약 2초 후 사용 가능. SP↔DP 전환 및 반복 Stop/Play 모두 안정적으로 동작.

**2. 그래프 x=0 정렬 불일치**

- **증상**: TMS 발화 시 EMG artifact가 x=0(trigger line)에서 ±30~130ms 오차로 찍히며, 매번 다름.
- **원인 분석**:
  - 기존 방식: 버튼 클릭 시 `current_time() + 0.2001`을 trigger time으로 사전 계산.
  - 오차 원인 ①: `play_message()` COM 호출 지연(5~50ms 가변) → QTimer 시작 시점이 매번 달라짐 → 실제 발화 시각 가변.
  - 오차 원인 ②: `GetRecordLength` 기반 `current_time()`이 LabChart 내부 버퍼링으로 인해 실제 녹화 시각보다 ~130ms 뒤처진 값을 반환함. COM API 한계.
  - `_fire_triggerbox`로 포인트를 옮겨서 `current_time() + 0.0501`을 사용해도 동일 문제 (GetRecordLength 지연이 근본 원인).
- **해결책: `OnCommentAdded` 이벤트 활용**
  - LabChart는 FRO output 발화 직후 comment를 자동으로 기록함 (SP: Output1 직후, DP: Output2 직후).
  - `OnCommentAdded` 이벤트의 `tick` 파라미터 → `tick * spt` = 정확한 발화 시각.
  - COM background thread를 preload 완료 후(`_fro_preload_done` 2차 pass)에 시작하여 이 이벤트를 수신.

**3. COM thread 운용 전략 (최종)**

| 구간 | COM thread 상태 | 이유 |
|------|----------------|------|
| FRO 모드 선택 시 | 중단 | 초기에는 불필요 |
| Play → Preload 전 구간 | 중단 | PlayMessage와 PumpWaitingMessages 겹침 방지 |
| Preload 완료, 버튼 활성화 시 | **시작** | OnCommentAdded 수신 준비 |
| 버튼 클릭 시 (`_fro_firing=True`) | 일시 중지 | PlayMessage 전송 구간 보호 |
| TMS 발화 후 (t≈200ms~500ms) | **재개** | comment 이벤트 처리 |
| Stop 누름 | 중단 | 다음 Play까지 불필요 |

- `_fro_firing` 플래그가 이미 있었고, COM thread 내부 루프가 이 플래그를 체크하므로 추가 코드 최소화.

**4. LabChart COM API — Comment 읽기 불가 확인**

- `doc.NumberOfComments` 없음. `dir(doc)`로 확인: comment 관련 API = `AddCommentAtSelection`, `AppendComment` 두 개뿐.
- 즉, COM API로는 comment를 쓸 수만 있고 읽을 수 없음. WithEvents(`OnCommentAdded`)를 통한 실시간 이벤트 수신만 가능.

#### 시도 기록 (시간순)

| # | 시도 | 결과 |
|---|------|------|
| 1 | `data_timer` 중단/재개 타이밍 조정 (발화 전 중단, 500ms 후 재개) | 에러 패턴 동일 |
| 2 | COM thread 중단 (FRO 모드에서 `PumpWaitingMessages` 완전 제거) | 에러 동일. COM thread가 원인 아님을 확인 |
| 3 | `_fro_firing` 플래그로 발화 구간만 PumpWaitingMessages 차단 | 에러 동일 |
| 4 | `_fro_pending_trigger_t` 를 버튼 클릭 직전에 계산 (COM 호출 최소화) | 에러 동일 |
| 5 | Warmup: StartSampling 후 SP+DP PlayMessage를 500ms 간격으로 먼저 전송 | 첫 Play 버튼 안 먹힘, Stop→Play 에러 |
| 6 | `_fro_last_config` 로 중복 PlayMessage 스킵 (SP→DP 전환 시에만 전송) | 에러 동일. 다만 Warmup DP와 버튼 DP 중복 문제 해결 |
| 7 | Bounce: StartSampling → 500ms → StopSampling → StartSampling (수동 Stop/Play 자동화 시도) | 첫 Play SP/DP 안 먹힘, 정지→재생 에러 |
| 8 | Preload: StartSampling **전에** SP+DP PlayMessage 전송 | 첫 Play SP/DP 발화 안 됨, Stop→Play 에러/앱 종료 |
| 9 | **Preload + Auto-bounce** (StartSampling 전 SP+DP → Start → 500ms → Stop → SP+DP → Start) | **성공. SP/DP 모두 안정 동작. Stop/Play 반복도 정상** |
| 10 | `_fire_triggerbox`에서 `current_time() + 0.0501` 로 trigger time 개선 시도 | 그래프 정렬 변화 없음. `GetRecordLength` 자체가 ~130ms 지연됨을 확인 |
| 11 | LabChart COM API comment 읽기 탐색 (`NumberOfComments` 등) | API 없음. `dir(doc)` 결과: `AddCommentAtSelection`, `AppendComment` 만 존재 |
| 12 | **COM thread + `OnCommentAdded` 활용** (preload 완료 후 thread 시작, 이벤트로 정확한 tick 수신) | **성공. artifact가 x=0에 정확히 정렬됨** |

#### 핵심 발견 사항

- **PowerLab FRO는 Stop/Start 사이클이 1회 필요**: 첫 StartSampling 이후 FRO가 제대로 초기화되지 않음. 이건 Python/GUI 코드 버그가 아니라 PowerLab 하드웨어 동작.
- **`GetRecordLength` 기반 `current_time()`은 ~130ms 지연**: LabChart가 데이터를 내부 버퍼에 쌓고 나중에 기록하기 때문. trigger time 추정에 사용 불가.
- **COM thread는 에러 원인이 아니었음**: COM thread 유무와 관계없이 동일 에러 패턴. 억울하게 의심받음.
- **`OnCommentAdded` tick 값이 가장 정확한 trigger time**: LabChart가 FRO output 발화 시 즉시 comment를 기록하며 이 tick은 버퍼 지연 없이 실제 발화 시각을 반영.

#### 변경 파일 및 함수

- `realtime_viewer.py`:
  - `_start()` → FRO 모드 시 `_fro_preload()` 호출
  - `_fro_preload()`, `_fro_preload_dp()`, `_fro_preload_done()`: preload + auto-bounce chain
  - `_fro_auto_bounce()`: 자동 stop → preload 재시작
  - `_fro_preload_done()` 2차 pass: `_register_com_events()` 호출
  - `_stop()`: FRO 모드 시 `_stop_com_thread()` 호출
  - `_on_tms_trigger()`: FRO 모드 guard 제거 → SP/DP 모두 OnCommentAdded로 trigger time 등록
  - `_fire_triggerbox()`: FRO 모드 직접 `_register_trigger` 제거 (OnCommentAdded가 처리)
  - `_artifact_free_range()`: NaN 데이터 방어 처리 추가

- `test_fro_pulse.py`:
  - `start_sampling()`, `stop_sampling()` 함수 추가, menu option s/x 추가
  - option 7: `dir(doc)` comment 관련 속성 탐색 (확인 용도)

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

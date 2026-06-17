using UnityEngine;
using TMPro;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO.Ports;

/// <summary>
/// Real World Reaching trial controller.
///
/// State machine:
///   MoveToStart → HoldInStart → ShowDirection → WaitForGo → Executing → Feedback → TrialDone
///
/// Instruction codes:  0 = REST,  1 = REACH,  2 = REACH+GRASP
/// False start rule:   zone exit before Go → full reset to MoveToStart
/// After Go:           REST zone-exit is noted for GOOD/BAD; REACH/R+G zone exit is fine
/// Outcome at t=N:
///   REST        → MCP inside start zone     → GOOD
///   REACH       → MCP inside target zone    → GOOD
///   REACH+GRASP → index tip AND thumb tip inside target zone → GOOD
/// </summary>
public class TrialGameController_RWR : MonoBehaviour
{
    public event Action OnTrialFinished;

    // ── Instruction constants ───────────────────────────────────────
    const int INST_REST  = 0;
    const int INST_REACH = 1;
    const int INST_RG    = 2;

    // ── State machine ───────────────────────────────────────────────
    enum TrialState
    {
        Idle,
        MoveToStart,
        HoldInStart,
        ShowDirection,
        WaitForGo,
        Executing,
        Feedback,
        TrialDone
    }

    // ── Inspector ───────────────────────────────────────────────────

    [Header("UI")]
    [SerializeField] TMP_Text instructionText;   // large centre text
    [SerializeField] TMP_Text debugText;         // side debug overlay

    [Header("Core References")]
    [SerializeField] Transform startSphere;      // visualised start zone
    [SerializeField] Transform targetSphere;     // visualised target zone

    [Header("Finger Transforms (from LeapFingerInput)")]
    [SerializeField] Transform indexTip;         // index fingertip
    [SerializeField] Transform thumbTip;         // thumb fingertip
    [SerializeField] Transform indexMcp;         // index MCP

    [Header("Radii (metres)")]
    [SerializeField] float startRadius  = 0.03f;
    [SerializeField] float targetRadius = 0.03f;
    public float StartRadius => startRadius;

    [Header("Timing (seconds)")]
    [SerializeField] float holdDuration      = 0.5f;   // hold in start before direction cue
    [SerializeField] float goDelay           = 2.0f;   // direction cue → go
    [SerializeField] float executionDuration = 3.0f;   // fixed N-second execution window
    [SerializeField] float feedbackDuration  = 1.0f;   // GOOD/BAD display time

    [Header("Rendering")]
    [SerializeField] Renderer startRenderer;
    [SerializeField] Renderer targetRenderer;
    [SerializeField] Color startIdleColor   = new Color(0.5f, 0.5f, 0.5f, 0.3f);
    [SerializeField] Color startReadyColor  = new Color(0f,   1f,   0f,   0.3f);
    [SerializeField] Color targetIdleColor  = new Color(0.5f, 0.5f, 0.5f, 0.3f);
    [SerializeField] Color targetActiveColor= new Color(1f,   1f,   0f,   0.3f);
    [SerializeField] Color targetGoodColor  = new Color(0f,   1f,   0f,   0.3f);
    [SerializeField] Color targetBadColor   = new Color(1f,   0f,   0f,   0.3f);

    [Header("Cursor Objects")]
    [SerializeField] GameObject mcpCursor;
    [SerializeField] GameObject indexTipCursor;
    [SerializeField] GameObject thumbTipCursor;
    [SerializeField] bool showMcpCursor        = true;
    [SerializeField] bool showIndexTipCursor   = false;
    [SerializeField] bool showThumbTipCursor   = false;

    [Header("Audio")]
    [SerializeField] AudioSource audioSource;
    [SerializeField] AudioClip   readyClip;
    [SerializeField] AudioClip   goClip;
    [SerializeField] AudioClip   goodClip;
    [SerializeField] AudioClip   badClip;

    [Header("TTL")]
    [SerializeField] float  ttlOffsetMs       = 0f;
    [SerializeField] float  ttl2OffsetMs      = 2.5f;
    [SerializeField] string ttlComPort        = "COM5";
    [SerializeField] int    ttlChannel        = 1;       // 1–8, maps to bit (ch1=1, ch2=2, ch3=4 ...)
    [SerializeField] float  ttlPulseDurationMs = 100f;   // how long the line stays high (ms) — matches Python reference (TRIG_PULSE_MS=100)
    [SerializeField] Renderer ttlLampRenderer;
    [SerializeField] Color    ttlLampOffColor  = Color.black;
    [SerializeField] Color    ttlLampOnColor   = Color.yellow;
    [SerializeField] float    ttlLampDuration  = 0.1f;

    [Header("LabChart FRO")]
    [SerializeField] LabChartFro froController;

    [Header("Zone Cylinder Heights")]
    [SerializeField] float startHeight  = 0.05f;
    [SerializeField] float targetHeight = 0.05f;

    [Header("References")]
    [SerializeField] LeapFingerInput leapInput;
    [SerializeField] Gettinghanddata handData;
    [SerializeField] TrialDataLogger dataLogger;

    // ── Private state ───────────────────────────────────────────────
    TrialState state = TrialState.Idle;

    int  currentInstruction = INST_REST;
    int  currentTrialIndex  = 0;
    int  currentTargetId    = 0;
    int  currentHandMode    = 1;

    float holdTimer      = 0f;
    float readyTime      = -1f;
    float goTime         = -1f;
    float execTimer      = 0f;
    float feedbackTimer  = 0f;
    float ttlLampTimer   = 0f;

    bool ttlEnabled    = true;
    bool ttlPending    = false;
    bool ttlFired      = false;
    float ttlFiredTime = -1f;
    float ttlPlannedTime = 0f;

    bool notifiedFinished = false;
    bool outcomeGood      = false;

    bool   paused          = false;
    float  pauseStartTime  = 0f;
    string textBeforePause = "";

    bool cursorsOverrideHidden = false;
    bool spheresOverrideHidden = false;

    SerialPort ttlPort = null;

    // ── Public entry point ──────────────────────────────────────────
    public void ConfigureAndBegin(
        Vector3 startPos,
        Vector3 targetPos,
        float   targetRadiusMeters,
        bool    ttlEnabledForTrial,
        float   ttlMs,
        float   ttl2Ms,
        int     trialIndex,
        int     targetId,
        int     handMode,
        int     instruction,
        float   perTrialHoldDuration      = 0f,  // 0 = use Inspector value
        float   perTrialWaitForGo         = 0f,  // 0 = use Inspector value
        float   perTrialExecutingDuration = 0f,  // 0 = use Inspector value
        float   perTrialStartRadiusCm     = 0f)  // 0 = use Inspector value
    {
        if (startSphere)  startSphere.position  = startPos;
        if (targetSphere) targetSphere.position = targetPos;

        targetRadius        = targetRadiusMeters;
        ttlEnabled          = ttlEnabledForTrial;
        ttlOffsetMs         = ttlMs;
        ttl2OffsetMs        = ttl2Ms;
        currentTrialIndex   = trialIndex;
        currentTargetId     = targetId;
        currentHandMode     = handMode;
        currentInstruction  = Mathf.Clamp(instruction, 0, 2);

        // Apply per-trial timing overrides (0 = keep Inspector default)
        if (perTrialHoldDuration      > 0f) holdDuration      = perTrialHoldDuration;
        if (perTrialWaitForGo         > 0f) goDelay           = perTrialWaitForGo;
        if (perTrialExecutingDuration > 0f) executionDuration = perTrialExecutingDuration;
        if (perTrialStartRadiusCm     > 0f) startRadius       = perTrialStartRadiusCm / 100f;

        if (startSphere)
            startSphere.localScale  = new Vector3(startRadius * 2f, startHeight, startRadius * 2f);
        if (targetSphere)
            targetSphere.localScale = new Vector3(targetRadiusMeters * 2f, targetHeight, targetRadiusMeters * 2f);

        InitTrial();
    }

    void InitTrial()
    {
        state            = TrialState.MoveToStart;
        holdTimer        = 0f;
        readyTime        = -1f;
        goTime           = -1f;
        execTimer        = 0f;
        feedbackTimer    = 0f;
        ttlPending       = false;
        ttlFired         = false;
        ttlFiredTime     = -1f;
        ttlPlannedTime   = 0f;
        notifiedFinished = false;
        outcomeGood      = false;

        if (!spheresOverrideHidden)
        {
            if (startSphere)  startSphere.gameObject.SetActive(true);
            if (targetSphere) targetSphere.gameObject.SetActive(true);
        }

        SetStartColor(startIdleColor);
        SetTargetColor(targetIdleColor);
        SetCursors(true);

        if (instructionText)
            instructionText.text = "Put your hand on home position";

        if (ShouldLog() && leapInput)
            dataLogger.Setup(leapInput.leapProvider, indexTip, thumbTip, indexMcp, null);
    }

    void Awake()
    {
        if (!startRenderer  && startSphere)  startRenderer  = startSphere.GetComponentInChildren<Renderer>();
        if (!targetRenderer && targetSphere) targetRenderer = targetSphere.GetComponentInChildren<Renderer>();

        OpenTtlPort();
    }

    void OnDestroy()
    {
        CloseTtlPort();
    }

    void OpenTtlPort()
    {
        if (string.IsNullOrEmpty(ttlComPort)) return;
        try
        {
            ttlPort = new SerialPort(ttlComPort, 115200);
            ttlPort.Open();
            // Reset all channels to 0 on open
            ttlPort.Write(new byte[] { 0 }, 0, 1);
            Debug.Log($"[TrialGameController_RWR] TTL port {ttlComPort} opened.");
        }
        catch (Exception e)
        {
            Debug.LogWarning($"[TrialGameController_RWR] Could not open TTL port {ttlComPort}: {e.Message}");
            ttlPort = null;
        }
    }

    void CloseTtlPort()
    {
        if (ttlPort != null && ttlPort.IsOpen)
        {
            try { ttlPort.Write(new byte[] { 0 }, 0, 1); } catch { }
            ttlPort.Close();
        }
        ttlPort = null;
    }

    void Update()
    {
        if (Input.GetKeyDown(KeyCode.P))
            TogglePause();

        if (paused)
        {
            UpdateCursors();
            UpdateDebug();
            return;
        }

        // TTL lamp countdown
        if (ttlLampTimer > 0f)
        {
            ttlLampTimer -= Time.deltaTime;
            if (ttlLampTimer <= 0f && ttlLampRenderer)
                ttlLampRenderer.material.color = ttlLampOffColor;
        }

        // TTL fire check
        if (ttlPending && !ttlFired && Time.time >= ttlPlannedTime)
            FireTtlPulse();

        switch (state)
        {
            case TrialState.MoveToStart:    Update_MoveToStart();    break;
            case TrialState.HoldInStart:    Update_HoldInStart();    break;
            case TrialState.ShowDirection:  Update_ShowDirection();  break;
            case TrialState.WaitForGo:      Update_WaitForGo();      break;
            case TrialState.Executing:      Update_Executing();      break;
            case TrialState.Feedback:       Update_Feedback();       break;
            case TrialState.TrialDone:      Update_TrialDone();      break;
        }

        UpdateCursors();
        UpdateDebug();
    }

    // ── State handlers ──────────────────────────────────────────────

    void Update_MoveToStart()
    {
        if (McpInStart())
        {
            state     = TrialState.HoldInStart;
            holdTimer = 0f;
            if (instructionText) instructionText.text = "+";
        }
    }

    void Update_HoldInStart()
    {
        if (!McpInStart())
        {
            ResetToMoveToStart();
            return;
        }

        holdTimer += Time.deltaTime;
        if (holdTimer >= holdDuration)
        {
            SetStartColor(startReadyColor);
            PlaySound(readyClip);
            readyTime = Time.time;

            // Begin kinematic recording at ShowDirection
            if (ShouldLog() && leapInput)
            {
                dataLogger.BeginTrial(
                    currentTrialIndex, currentTargetId, currentHandMode,
                    leapInput.useLeftHand, leapInput.lastTimestampUs,
                    startSphere ? startSphere.position : Vector3.zero, startRadius,
                    targetSphere ? targetSphere.position : Vector3.zero, targetRadius,
                    holdDuration, goDelay, executionDuration, feedbackDuration,
                    ttlOffsetMs, readyTime, -1f
                );
            }

            state = TrialState.ShowDirection;
            ShowDirectionCue();
        }
    }

    void Update_ShowDirection()
    {
        if (!McpInStart())
        {
            ResetToMoveToStart();
            return;
        }

        if (Time.time - readyTime >= goDelay)
        {
            goTime = Time.time;
            EnterGo();
        }
    }

    void Update_WaitForGo()
    {
        // Kept for potential future use; currently ShowDirection transitions directly to Executing via EnterGo
    }

    void Update_Executing()
    {
        execTimer += Time.deltaTime;

        // REST trial: remind subject to return if they leave; otherwise keep fixation + cue
        if (currentInstruction == INST_REST && !McpInStart())
        {
            if (instructionText)
                instructionText.text = "REST — please return to home position";
        }
        else if (currentInstruction == INST_REST && McpInStart())
        {
            if (instructionText)
                instructionText.text = "+\n<color=#888888><size=65%><i>Rest</i></size></color>";
        }

        if (execTimer >= executionDuration)
        {
            EvaluateOutcome();
            EnterFeedback();
        }
    }

    void Update_Feedback()
    {
        feedbackTimer += Time.deltaTime;
        if (feedbackTimer >= feedbackDuration)
        {
            if (startSphere)  startSphere.gameObject.SetActive(false);
            if (targetSphere) targetSphere.gameObject.SetActive(false);

            if (ShouldLog()) dataLogger.EndAndSave();

            state = TrialState.TrialDone;
        }
    }

    void Update_TrialDone()
    {
        if (!notifiedFinished)
        {
            notifiedFinished = true;
            OnTrialFinished?.Invoke();
        }
    }

    // ── Helpers ─────────────────────────────────────────────────────

    void ShowDirectionCue()
    {
        string instrLabel = currentInstruction switch
        {
            INST_REST  => "<i>Rest</i>",
            INST_REACH => "<i>Reach</i>",
            INST_RG    => "<i>Reach & Grasp</i>",
            _          => "<i>Rest</i>"
        };

        // Fixation cross on top; dimmed italic direction cue below.
        // At Go: non-REST replaces this with "GO"; REST keeps it as-is.
        if (instructionText)
            instructionText.text = $"+\n<color=#888888><size=65%>{instrLabel}</size></color>";
        PlaySound(readyClip);

        // TTL trigger fires 500ms before Go cue (fixed reference point).
        // ttlOffsetMs is expressed relative to Go cue (0 = at Go, -50 = 50ms before Go).
        // Internal: trigger fires at goPlanned - 0.5s regardless of ttlOffsetMs.
        float goPlanned = readyTime + goDelay;
        ttlPlannedTime  = goPlanned - 0.5f;
        ttlPending      = true;   // trigger always fires (LabChart event marker), regardless of ttlEnabled
        ttlFired        = false;
        ttlFiredTime    = -1f;

        // Always arm FRO to reset LabChart state from previous trial.
        // ttlEnabled=false → NoPulse (both outputs disabled, clears previous state).
        // ttlEnabled=true  → normal stimulation settings.
        //   Output1 (Conditioning) = (500 + ttl1 + ttl2) ms from trigger  (variable, fires first when ttl2 < 0)
        //   Output2 (Testing)      = (500 + ttl1) ms from trigger          (fixed reference)
        // ttl2 == 0 means SinglePulse (Output1/Conditioning disabled).
        if (froController != null)
        {
            // Cancel any in-progress coroutine from a previous trial
            froController.CancelPrepare();

            if (!ttlEnabled)
            {
                froController.activeCoroutine = StartCoroutine(froController.PrepareNoPulse());
            }
            else
            {
                float out1Abs     = 500f + ttlOffsetMs + ttl2OffsetMs;  // Conditioning
                float out2Abs     = 500f + ttlOffsetMs;                  // Testing
                bool  doublePulse = ttl2OffsetMs != 0f;

                if (doublePulse && (out1Abs < 0f || out1Abs >= 9900f))
                    Debug.LogWarning($"[TrialGameController_RWR] Output1 (Conditioning) absolute delay out of range: {out1Abs:F1}ms. FRO skipped.");
                else if (out2Abs < 0f || out2Abs >= 9900f)
                    Debug.LogWarning($"[TrialGameController_RWR] Output2 (Testing) absolute delay out of range: {out2Abs:F1}ms. FRO skipped.");
                else
                    froController.activeCoroutine = StartCoroutine(froController.PrepareOutputs(out1Abs, out2Abs, doublePulse));
            }
        }
    }

    void EnterGo()
    {
        // REST: keep fixation cross + "Rest" cue (subject already knows to stay).
        // Non-REST: replace with "GO" so the subject knows to start moving.
        if (currentInstruction != INST_REST)
        {
            if (instructionText) instructionText.text = "<color=#00CC00>GO</color>";
        }
        SetTargetColor(targetActiveColor);
        PlaySound(goClip);

        // Recalculate TTL trigger time from actual go time — only if not already fired.
        // Trigger always fires 1 second before Go cue.
        if (!ttlFired)
        {
            ttlPlannedTime = goTime - 0.5f;
            ttlPending     = true;   // trigger always fires
        }

        if (ShouldLog()) dataLogger.SetGoTime(goTime);

        execTimer = 0f;
        state     = TrialState.Executing;
    }

    void EvaluateOutcome()
    {
        switch (currentInstruction)
        {
            case INST_REST:
                outcomeGood = McpInStart();
                break;
            case INST_REACH:
                outcomeGood = McpInTarget();
                break;
            case INST_RG:
                outcomeGood = TipInTarget(indexTip) && TipInTarget(thumbTip);
                break;
            default:
                outcomeGood = false;
                break;
        }
    }

    void EnterFeedback()
    {
        if (instructionText) instructionText.text = outcomeGood ? "<color=#00CC00>GOOD</color>" : "<color=#FF3333>BAD</color>";
        SetTargetColor(outcomeGood ? targetGoodColor : targetBadColor);
        PlaySound(outcomeGood ? goodClip : badClip);
        feedbackTimer = 0f;
        state         = TrialState.Feedback;
    }

    void TogglePause()
    {
        paused = !paused;
        if (paused)
        {
            pauseStartTime  = Time.time;
            textBeforePause = instructionText ? instructionText.text : "";
            if (instructionText) instructionText.text = "<color=#FFFF44>PAUSE</color>";
        }
        else
        {
            // Shift all time references forward so trial timing stays correct
            float elapsed = Time.time - pauseStartTime;
            if (readyTime      > 0f) readyTime      += elapsed;
            if (goTime         > 0f) goTime         += elapsed;
            if (ttlPlannedTime > 0f) ttlPlannedTime += elapsed;

            if (instructionText) instructionText.text = textBeforePause;
        }
    }

    void ResetToMoveToStart()
    {
        state     = TrialState.MoveToStart;
        holdTimer = 0f;
        readyTime = -1f;
        goTime    = -1f;
        execTimer = 0f;
        ttlPending = false;

        SetStartColor(startIdleColor);
        SetTargetColor(targetIdleColor);
        SetCursors(true);

        if (instructionText)
            instructionText.text = "Put your hand on home position";

        Debug.Log("[TrialGameController_RWR] False start — reset to MoveToStart.");
    }

    void FireTtlPulse()
    {
        ttlFired     = true;
        ttlPending   = false;
        ttlFiredTime = Time.time;

        if (ShouldLog() && leapInput)
            dataLogger.NoteTtlFired(leapInput.lastTimestampUs);

        if (ttlLampRenderer)
        {
            ttlLampRenderer.material.color = ttlLampOnColor;
            ttlLampTimer = ttlLampDuration;
        }

        Debug.Log($"[TrialGameController_RWR] TTL fired — {(ttlFiredTime - goTime) * 1000f:F1} ms from Go (target offset: {ttlOffsetMs} ms)");

        // NoPulse trials have no FRO event — append a text comment so Event mode shows a label.
        if (!ttlEnabled && froController != null)
            StartCoroutine(froController.AppendCommentCoroutine($"Trial {currentTrialIndex}"));

        // Hardware pulse: write channel byte high, then reset after pulse duration
        if (ttlPort != null && ttlPort.IsOpen)
        {
            byte channelByte = (byte)(1 << (ttlChannel - 1)); // ch1=0x01, ch2=0x02, ch3=0x04 ...
            try
            {
                ttlPort.Write(new byte[] { channelByte }, 0, 1);
                ttlPort.BaseStream.Flush();
                StartCoroutine(ResetTtlAfterDelay(ttlPulseDurationMs / 1000f));
            }
            catch (Exception e)
            {
                Debug.LogWarning($"[TrialGameController_RWR] TTL write failed: {e.Message}");
            }
        }
    }

    IEnumerator ResetTtlAfterDelay(float delaySeconds)
    {
        yield return new WaitForSecondsRealtime(delaySeconds);
        if (ttlPort != null && ttlPort.IsOpen)
        {
            try { ttlPort.Write(new byte[] { 0 }, 0, 1); ttlPort.BaseStream.Flush(); }
            catch (Exception e) { Debug.LogWarning($"[TrialGameController_RWR] TTL reset failed: {e.Message}"); }
        }
    }

    // ── Zone checks (XZ-plane, same convention as R/RG) ─────────────

    bool McpInStart()
    {
        if (!indexMcp || !startSphere) return false;
        if (leapInput != null && !leapInput.hasIndexJointData) return false;
        float dx = indexMcp.position.x - startSphere.position.x;
        float dz = indexMcp.position.z - startSphere.position.z;
        return dx * dx + dz * dz <= startRadius * startRadius;
    }

    bool McpInTarget()
    {
        if (!indexMcp || !targetSphere) return false;
        if (leapInput != null && !leapInput.hasIndexJointData) return false;
        float dx = indexMcp.position.x - targetSphere.position.x;
        float dz = indexMcp.position.z - targetSphere.position.z;
        return dx * dx + dz * dz <= targetRadius * targetRadius;
    }

    bool TipInTarget(Transform tip)
    {
        if (!tip || !targetSphere) return false;
        float dx = tip.position.x - targetSphere.position.x;
        float dz = tip.position.z - targetSphere.position.z;
        return dx * dx + dz * dz <= targetRadius * targetRadius;
    }

    // ── Visuals ──────────────────────────────────────────────────────

    void SetStartColor(Color c)  { if (startRenderer)  SetTransparentColor(startRenderer.material,  c); }
    void SetTargetColor(Color c) { if (targetRenderer) SetTransparentColor(targetRenderer.material, c); }

    void SetTransparentColor(Material mat, Color c)
    {
        mat.SetFloat("_Mode", 3);
        mat.SetInt("_SrcBlend", (int)UnityEngine.Rendering.BlendMode.SrcAlpha);
        mat.SetInt("_DstBlend", (int)UnityEngine.Rendering.BlendMode.OneMinusSrcAlpha);
        mat.SetInt("_ZWrite", 0);
        mat.DisableKeyword("_ALPHATEST_ON");
        mat.EnableKeyword("_ALPHABLEND_ON");
        mat.DisableKeyword("_ALPHAPREMULTIPLY_ON");
        mat.renderQueue = 3000;
        mat.color = c;
    }

    void SetCursors(bool visible)
    {
        if (cursorsOverrideHidden) visible = false;
        if (mcpCursor)       mcpCursor.SetActive(showMcpCursor && visible);
        if (indexTipCursor)  indexTipCursor.SetActive(showIndexTipCursor && visible);
        if (thumbTipCursor)  thumbTipCursor.SetActive(showThumbTipCursor && visible);
    }

    public void SetCursorsVisible(bool visible)
    {
        cursorsOverrideHidden = !visible;
        if (mcpCursor)       mcpCursor.SetActive(visible && showMcpCursor);
        if (indexTipCursor)  indexTipCursor.SetActive(visible && showIndexTipCursor);
        if (thumbTipCursor)  thumbTipCursor.SetActive(visible && showThumbTipCursor);
    }

    public void SetSpheresVisible(bool visible)
    {
        spheresOverrideHidden = !visible;
        if (startSphere)  startSphere.gameObject.SetActive(visible);
        if (targetSphere) targetSphere.gameObject.SetActive(visible);
    }

    public void SetInstructionFontSize(float size)
    {
        if (instructionText) instructionText.fontSize = size;
    }

    public void SetExperimentingMode(bool value)
    {
        if (dataLogger) dataLogger.SetExperimentingMode(value);
    }

    void UpdateCursors()
    {
        if (mcpCursor      && showMcpCursor      && indexMcp)  mcpCursor.transform.position      = indexMcp.position;
        if (indexTipCursor && showIndexTipCursor && indexTip)  indexTipCursor.transform.position = indexTip.position;
        if (thumbTipCursor && showThumbTipCursor && thumbTip)  thumbTipCursor.transform.position = thumbTip.position;
    }

    void PlaySound(AudioClip clip)
    {
        if (!audioSource || !clip) return;
        audioSource.PlayOneShot(clip);
    }

    // ── Debug overlay ────────────────────────────────────────────────

    void UpdateDebug()
    {
        if (!debugText) return;

        string instrName = currentInstruction switch
        {
            INST_REST  => "REST",
            INST_REACH => "REACH",
            INST_RG    => "REACH+GRASP",
            _          => "?"
        };

        string ttlStatus = !ttlEnabled
            ? "TTL: none"
            : ttlFired
                ? $"TTL: fired ({(ttlFiredTime - goTime) * 1000f:F1} ms from Go)"
                : ttlPending
                    ? $"TTL: pending (offset {ttlOffsetMs} ms)"
                    : "TTL: waiting";

        string mcpLine = "MCP: (no data)";
        if (handData != null)
        {
            bool useLeft = leapInput != null && leapInput.useLeftHand;
            Vector3 mcp  = useLeft ? handData.leftMcpPos : handData.rightMcpPos;
            mcpLine = $"MCP ({(useLeft ? "L" : "R")}): {mcp:F3}";
        }

        float elapsed = goTime > 0f ? Mathf.Max(0f, Time.time - goTime) : 0f;

        debugText.text =
            $"State: {state}\n" +
            $"Instruction: {instrName}\n" +
            $"Trial: {currentTrialIndex}  Target: {currentTargetId}\n" +
            $"Exec: {(state == TrialState.Executing ? execTimer : 0f):F2} / {executionDuration:F1} s\n" +
            $"Outcome: {(state >= TrialState.Feedback ? (outcomeGood ? "GOOD" : "BAD") : "-")}\n" +
            ttlStatus + "\n" +
            mcpLine;
    }

    // ── Logging helpers ───────────────────────────────────────────────

    bool ShouldLog() =>
        dataLogger != null &&
        RuntimeConfigStore.Instance != null &&
        RuntimeConfigStore.Instance.enableTrialLogging;

    public int GetStateCode() => (int)state;
}

using System;
using System.Collections;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using UnityEngine;

/// <summary>
/// Controls LabChart Fast Response Output (FRO) per trial via COM automation.
///
/// The TTL trigger (TriggerBox pulse) always fires 1 second before the Go cue.
/// FRO delays are measured from that trigger:
///   Output1 (Testing Stimulus)     delay = (1000 + ttl1) ms
///   Output2 (Conditioning Stimulus) delay = (1000 + ttl1 + ttl2) ms
///
/// ttl2 == 0  → SinglePulse (Output2 disabled)
/// ttl2 != 0  → DoublePulse (both outputs active)
///
/// Template files (must be recorded in LabChart):
///   DoublePulse: Firstt.vbs      — Output1 = 0.0501s placeholder, Output2 = 0.0525s placeholder
///   SinglePulse: SinglePulse.vbs — Output1 = 0.0501s placeholder, Output2 disabled (On=0)
///
/// Requires: .NET Framework API level (Project Settings → Player → Api Compatibility Level)
/// </summary>
public class LabChartFro : MonoBehaviour
{
    [Header("FRO Templates")]
    [Tooltip("DoublePulse VBS: Output1=0.0501s placeholder, Output2=0.0525s placeholder, both enabled")]
    [SerializeField] string vbsDoublePulsePath = @"C:\Users\Jisung Yuk\Documents\leapUnity\Source\DoublePulse.vbs";

    [Tooltip("SinglePulse VBS: Output1=0.0501s placeholder, Output2 disabled (On=0)")]
    [SerializeField] string vbsSinglePulsePath = @"C:\Users\Jisung Yuk\Documents\leapUnity\Source\SinglePulse.vbs";

    [Tooltip("NoPulse VBS: Output1 and Output2 both disabled (On=0). Used when ttlEnabled=false to reset FRO state.")]
    [SerializeField] string vbsNoPulsePath = @"C:\Users\Jisung Yuk\Documents\leapUnity\Source\NoPulse.vbs";

    [Tooltip("Delay between template restore and modified message (ms). ~50ms recommended.")]
    [SerializeField] float interMessageDelayMs = 50f;

    // 6-char placeholders — must match exactly what LabChart encoded in the VBS files
    const string TEMPLATE_DELAY_OUT1 = "0.0501";  // Output1 placeholder (6 chars)
    const string TEMPLATE_DELAY_OUT2 = "0.0525";  // Output2 placeholder (6 chars)

    // Cached template bytes for each mode
    byte[] templateBytesDouble  = null;
    byte[] templateBytesSingle  = null;
    byte[] templateBytesNoPulse = null;
    bool   doubleLoaded  = false;
    bool   singleLoaded  = false;
    bool   noPulseLoaded = false;

    // Track active coroutine so it can be cancelled if a new trial starts before it finishes
    [HideInInspector] public Coroutine activeCoroutine = null;

    // ── Public API ───────────────────────────────────────────────────

    /// <summary>
    /// Call this during ShowDirectionCue (before the Go cue).
    /// out1AbsoluteMs and out2AbsoluteMs are ABSOLUTE delays from the TTL trigger pulse.
    ///   out1AbsoluteMs = 500 + ttl1   (ms)
    ///   out2AbsoluteMs = 500 + ttl1 + ttl2   (ms)
    /// doublePulse=false → SinglePulse template (Output2 disabled).
    /// Use SendNoPulse() instead when ttlEnabled=false.
    /// Returns a coroutine — start it with StartCoroutine().
    /// </summary>
    /// <summary>Cancels any in-progress PrepareOutputs/PrepareNoPulse coroutine.</summary>
    public void CancelPrepare()
    {
        if (activeCoroutine != null)
        {
            StopCoroutine(activeCoroutine);
            activeCoroutine = null;
            Debug.Log("[LabChartFro] Previous prepare coroutine cancelled.");
        }
    }

    public IEnumerator PrepareOutputs(float out1AbsoluteMs, float out2AbsoluteMs, bool doublePulse)
    {
        byte[] templateBytes = doublePulse
            ? EnsureTemplate(ref templateBytesDouble, ref doubleLoaded, vbsDoublePulsePath, "DoublePulse")
            : EnsureTemplate(ref templateBytesSingle, ref singleLoaded, vbsSinglePulsePath, "SinglePulse");

        if (templateBytes == null) yield break;

        // Output1 = Conditioning (variable), Output2 = Testing (fixed reference)
        float out1_s = out1AbsoluteMs / 1000f;
        float out2_s = out2AbsoluteMs / 1000f;

        // Output2 (Testing) always needed — validate first
        string out2Str = FormatDelay(out2_s);
        if (out2Str == null)
        {
            Debug.LogWarning($"[LabChartFro] Output2 (Testing) delay out of range: {out2AbsoluteMs:F1}ms ({out2_s:F4}s). Must be 0–9.9999s.");
            yield break;
        }

        // Output1 (Conditioning) only needed for DoublePulse
        string out1Str = null;
        if (doublePulse)
        {
            out1Str = FormatDelay(out1_s);
            if (out1Str == null)
            {
                Debug.LogWarning($"[LabChartFro] Output1 (Conditioning) delay out of range: {out1AbsoluteMs:F1}ms ({out1_s:F4}s). Must be 0–9.9999s.");
                yield break;
            }
        }

        // Step 1: restore known-good FRO state with template
        string templateHex = "0x" + BitConverter.ToString(templateBytes).Replace("-", "");
        SendPlayMessage(templateHex);

        // Step 2: wait so LabChart can process
        yield return new WaitForSecondsRealtime(interMessageDelayMs / 1000f);

        // Step 3: build modified message and send
        byte[] modified = doublePulse
            ? BuildModifiedDouble(templateBytes, out1Str, out2Str)
            : BuildModifiedSingle(templateBytes, out2Str);

        if (modified == null) yield break;

        string modifiedHex = "0x" + BitConverter.ToString(modified).Replace("-", "");
        SendPlayMessage(modifiedHex);

        if (doublePulse)
            Debug.Log($"[LabChartFro] FRO armed (DoublePulse): Conditioning(Out1)={out1AbsoluteMs:F1}ms  Testing(Out2)={out2AbsoluteMs:F1}ms  (from TTL trigger)");
        else
            Debug.Log($"[LabChartFro] FRO armed (SinglePulse): Testing(Out2)={out2AbsoluteMs:F1}ms  Conditioning=disabled");
    }

    /// <summary>
    /// Sends the NoPulse template to reset FRO — both outputs disabled.
    /// Call this when ttlEnabled=false so previous trial's FRO state is cleared.
    /// Returns a coroutine — start it with StartCoroutine().
    /// </summary>
    public IEnumerator PrepareNoPulse()
    {
        byte[] templateBytes = EnsureTemplate(ref templateBytesNoPulse, ref noPulseLoaded, vbsNoPulsePath, "NoPulse");
        if (templateBytes == null) yield break;

        // NoPulse template needs no modification — send as-is (both outputs are On=0)
        string templateHex = "0x" + BitConverter.ToString(templateBytes).Replace("-", "");
        SendPlayMessage(templateHex);

        // No second message needed — template itself is the final state
        Debug.Log("[LabChartFro] FRO reset (NoPulse): both outputs disabled.");
    }

    /// <summary>
    /// Appends a comment to the active LabChart document at the current recording position.
    /// Runs on a background thread — non-blocking, no TMS or FRO hardware interaction.
    /// </summary>
    public IEnumerator AppendCommentCoroutine(string text)
    {
        yield return null; // defer one frame so TTL pulse timing is not affected

        string vbs = null;
        try
        {
            vbs = Path.Combine(Path.GetTempPath(), "labchart_comment.vbs");
            File.WriteAllText(vbs,
                "On Error Resume Next\r\n" +
                "Set App = GetObject(,\"ADIChart.Application\")\r\n" +
                "If Err.Number <> 0 Then WScript.Quit 1\r\n" +
                "On Error GoTo 0\r\n" +
                "Set Doc = App.ActiveDocument\r\n" +
                $"Doc.AppendComment \"{text}\"\r\n",
                Encoding.ASCII);

            var psi = new System.Diagnostics.ProcessStartInfo("cscript.exe", $"//Nologo \"{vbs}\"")
            {
                CreateNoWindow        = true,
                UseShellExecute       = false,
                RedirectStandardError = true
            };
            using (var proc = System.Diagnostics.Process.Start(psi))
            {
                string err = proc.StandardError.ReadToEnd();
                proc.WaitForExit(2000);
                if (proc.ExitCode != 0 || !string.IsNullOrEmpty(err))
                    Debug.LogWarning($"[LabChartFro] AppendComment error: {err}");
                else
                    Debug.Log($"[LabChartFro] AppendComment OK: \"{text}\"");
            }
        }
        catch (Exception e)
        {
            Debug.LogWarning($"[LabChartFro] AppendComment failed: {e.Message}");
        }
    }

    // ── Internal ─────────────────────────────────────────────────────

    byte[] EnsureTemplate(ref byte[] cache, ref bool loaded, string path, string label)
    {
        if (loaded) return cache;
        loaded = true;

        if (string.IsNullOrEmpty(path) || !File.Exists(path))
        {
            Debug.LogWarning($"[LabChartFro] {label} VBS template not found: {path}");
            return null;
        }

        try
        {
            string txt = File.ReadAllText(path, Encoding.UTF8);
            var match = Regex.Match(txt, @"PlayMessage\s*\(\s*""(0x[0-9A-Fa-f]+)""\s*\)");
            if (!match.Success)
            {
                Debug.LogWarning($"[LabChartFro] {label}: Could not find PlayMessage hex in VBS template.");
                return null;
            }

            string hex = match.Groups[1].Value.Substring(2);
            cache = HexToBytes(hex);
            Debug.Log($"[LabChartFro] {label} template loaded ({cache.Length} bytes).");

            // Verify Output1 placeholder
            byte[] needle1 = Encoding.Unicode.GetBytes("PulseDelay = " + TEMPLATE_DELAY_OUT1);
            int count1 = CountOccurrences(cache, needle1);
            if (count1 < 1)
                Debug.LogWarning($"[LabChartFro] {label}: Output1 placeholder 'PulseDelay = {TEMPLATE_DELAY_OUT1}' not found. Re-record VBS with Output1 = {TEMPLATE_DELAY_OUT1}s.");
            else
                Debug.Log($"[LabChartFro] {label}: Output1 placeholder OK ({count1} occurrence).");

            if (label == "DoublePulse")
            {
                byte[] needle2 = Encoding.Unicode.GetBytes("PulseDelay = " + TEMPLATE_DELAY_OUT2);
                int count2 = CountOccurrences(cache, needle2);
                if (count2 < 1)
                    Debug.LogWarning($"[LabChartFro] {label}: Output2 placeholder 'PulseDelay = {TEMPLATE_DELAY_OUT2}' not found. Re-record VBS with Output2 = {TEMPLATE_DELAY_OUT2}s.");
                else
                    Debug.Log($"[LabChartFro] {label}: Output2 placeholder OK ({count2} occurrence).");
            }

            return cache;
        }
        catch (Exception e)
        {
            Debug.LogWarning($"[LabChartFro] {label}: Failed to load template: {e.Message}");
            return null;
        }
    }

    /// <summary>Replaces Output1 and Output2 placeholders, fixes checksum.</summary>
    byte[] BuildModifiedDouble(byte[] template, string out1Str, string out2Str)
    {
        byte[] raw = (byte[])template.Clone();

        byte[] oldOut1 = Encoding.Unicode.GetBytes("PulseDelay = " + TEMPLATE_DELAY_OUT1);
        byte[] newOut1 = Encoding.Unicode.GetBytes("PulseDelay = " + out1Str);
        byte[] oldOut2 = Encoding.Unicode.GetBytes("PulseDelay = " + TEMPLATE_DELAY_OUT2);
        byte[] newOut2 = Encoding.Unicode.GetBytes("PulseDelay = " + out2Str);

        if (oldOut1.Length != newOut1.Length)
        {
            Debug.LogWarning($"[LabChartFro] Output1 length mismatch: placeholder={oldOut1.Length} new={newOut1.Length}. out1Str='{out1Str}' must be {TEMPLATE_DELAY_OUT1.Length} chars.");
            return null;
        }
        if (oldOut2.Length != newOut2.Length)
        {
            Debug.LogWarning($"[LabChartFro] Output2 length mismatch: placeholder={oldOut2.Length} new={newOut2.Length}. out2Str='{out2Str}' must be {TEMPLATE_DELAY_OUT2.Length} chars.");
            return null;
        }

        long origSum = SumBytes(raw);

        if (!ReplaceOccurrence(raw, oldOut1, newOut1, 1))
        {
            Debug.LogWarning($"[LabChartFro] Could not find Output1 placeholder in DoublePulse template.");
            return null;
        }
        if (!ReplaceOccurrence(raw, oldOut2, newOut2, 1))
        {
            Debug.LogWarning($"[LabChartFro] Could not find Output2 placeholder in DoublePulse template.");
            return null;
        }

        FixChecksum(raw, origSum);
        return raw;
    }

    /// <summary>Replaces Output2 placeholder only (SinglePulse — Output1/Conditioning is disabled in template).</summary>
    byte[] BuildModifiedSingle(byte[] template, string out2Str)
    {
        byte[] raw = (byte[])template.Clone();

        byte[] oldOut2 = Encoding.Unicode.GetBytes("PulseDelay = " + TEMPLATE_DELAY_OUT2);
        byte[] newOut2 = Encoding.Unicode.GetBytes("PulseDelay = " + out2Str);

        if (oldOut2.Length != newOut2.Length)
        {
            Debug.LogWarning($"[LabChartFro] Output2 length mismatch: placeholder={oldOut2.Length} new={newOut2.Length}. out2Str='{out2Str}' must be {TEMPLATE_DELAY_OUT2.Length} chars.");
            return null;
        }

        long origSum = SumBytes(raw);

        if (!ReplaceOccurrence(raw, oldOut2, newOut2, 1))
        {
            Debug.LogWarning($"[LabChartFro] Could not find Output2 placeholder in SinglePulse template.");
            return null;
        }

        FixChecksum(raw, origSum);
        return raw;
    }

    static void FixChecksum(byte[] raw, long origSum)
    {
        long newSum   = SumBytes(raw);
        long delta    = newSum - origSum;
        uint oldCsum  = BitConverter.ToUInt32(raw, 20);
        long newCsumL = ((((long)oldCsum + delta) % 0x100000000L) + 0x100000000L) % 0x100000000L;
        uint newCsum  = (uint)newCsumL;
        byte[] cb = BitConverter.GetBytes(newCsum);
        Array.Copy(cb, 0, raw, 20, 4);
        Debug.Log($"[LabChartFro] Checksum: delta={delta} old=0x{oldCsum:X8} new=0x{newCsum:X8}");
    }

    void SendPlayMessage(string hexMessage)
    {
        string vbs = null;
        try
        {
            vbs = Path.Combine(Path.GetTempPath(), "labchart_playmsg.vbs");
            File.WriteAllText(vbs,
                "On Error Resume Next\r\n" +
                "Set App = GetObject(,\"ADIChart.Application\")\r\n" +
                "If Err.Number <> 0 Then WScript.Quit 1\r\n" +
                "On Error GoTo 0\r\n" +
                "Set Doc = App.ActiveDocument\r\n" +
                $"Call Doc.PlayMessage(\"{hexMessage}\")\r\n",
                Encoding.ASCII);

            var psi = new System.Diagnostics.ProcessStartInfo("cscript.exe", $"//Nologo \"{vbs}\"")
            {
                CreateNoWindow        = true,
                UseShellExecute       = false,
                RedirectStandardError = true
            };
            using (var proc = System.Diagnostics.Process.Start(psi))
            {
                string err = proc.StandardError.ReadToEnd();
                proc.WaitForExit(2000);
                if (proc.ExitCode != 0 || !string.IsNullOrEmpty(err))
                    Debug.LogWarning($"[LabChartFro] cscript error (exit {proc.ExitCode}): {err}");
                else
                    Debug.Log("[LabChartFro] cscript OK.");
            }
        }
        catch (Exception e)
        {
            Debug.LogWarning($"[LabChartFro] SendPlayMessage failed: {e.GetType().Name} — {e.Message}");
        }
    }

    // ── Helpers ──────────────────────────────────────────────────────

    /// <summary>
    /// Formats a delay in seconds to exactly 6 characters (e.g. 1.0000, 0.9900).
    /// Supports 0.0000–9.9999 s. Returns null if out of range.
    /// </summary>
    static string FormatDelay(float seconds)
    {
        if (seconds < 0f || seconds >= 10f) return null;
        string s = seconds.ToString("0.0000", CultureInfo.InvariantCulture);
        return s.Length == 6 ? s : null;
    }

    static byte[] HexToBytes(string hex)
    {
        byte[] bytes = new byte[hex.Length / 2];
        for (int i = 0; i < bytes.Length; i++)
            bytes[i] = Convert.ToByte(hex.Substring(i * 2, 2), 16);
        return bytes;
    }

    static long SumBytes(byte[] data)
    {
        long sum = 0;
        foreach (byte b in data) sum += b;
        return sum;
    }

    static int CountOccurrences(byte[] haystack, byte[] needle)
    {
        int count = 0, start = 0;
        while (true)
        {
            int idx = IndexOf(haystack, needle, start);
            if (idx < 0) break;
            count++;
            start = idx + needle.Length;
        }
        return count;
    }

    static bool ReplaceOccurrence(byte[] buffer, byte[] oldBytes, byte[] newBytes, int occurrence)
    {
        int start = 0;
        for (int i = 0; i < occurrence; i++)
        {
            int idx = IndexOf(buffer, oldBytes, start);
            if (idx < 0) return false;
            if (i == occurrence - 1)
                Array.Copy(newBytes, 0, buffer, idx, newBytes.Length);
            start = idx + oldBytes.Length;
        }
        return true;
    }

    static int IndexOf(byte[] haystack, byte[] needle, int start = 0)
    {
        for (int i = start; i <= haystack.Length - needle.Length; i++)
        {
            bool match = true;
            for (int j = 0; j < needle.Length; j++)
                if (haystack[i + j] != needle[j]) { match = false; break; }
            if (match) return i;
        }
        return -1;
    }
}

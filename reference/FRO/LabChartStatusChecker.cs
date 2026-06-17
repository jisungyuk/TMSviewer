using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using UnityEngine;

/// <summary>
/// Polls LabChart open/recording status every few seconds on a background thread.
/// - Open: detected via process name "LabChart8"
/// - Recording: detected by checking all window titles of LabChart's process via EnumWindows
/// </summary>
public class LabChartStatusChecker : MonoBehaviour
{
    [SerializeField] float pollInterval = 2f;

    public bool IsOpen { get; private set; }

    bool busy = false;

    // ── Win32 P/Invoke ───────────────────────────────────────────────
    delegate bool EnumWindowsProc(System.IntPtr hWnd, System.IntPtr lParam);

    [DllImport("user32.dll")]
    static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, System.IntPtr lParam);

    [DllImport("user32.dll")]
    static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll")]
    static extern int GetWindowText(System.IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    static extern bool IsWindowVisible(System.IntPtr hWnd);

    // ── Unity lifecycle ──────────────────────────────────────────────
    void OnEnable()
    {
        IsOpen = false;
        InvokeRepeating(nameof(TriggerCheck), 0f, pollInterval);
    }

    void OnDisable() => CancelInvoke(nameof(TriggerCheck));

    void TriggerCheck()
    {
        if (!busy) _ = CheckAsync();
    }

    async Task CheckAsync()
    {
        busy = true;
        try
        {
            IsOpen = await Task.Run(() => CheckLabChart());
        }
        catch (System.Exception ex)
        {
            UnityEngine.Debug.LogWarning($"[LabChartStatus] Exception: {ex.Message}");
        }
        finally
        {
            busy = false;
        }
    }

    // Runs on background thread — returns true if LabChart process is running
    bool CheckLabChart()
    {
        var procs = Process.GetProcessesByName("LabChart8");
        if (procs.Length == 0) procs = Process.GetProcessesByName("LabChart");
        if (procs.Length == 0) return false;

        // Collect all PIDs belonging to LabChart
        var pidSet = new HashSet<uint>();
        foreach (var p in procs) pidSet.Add((uint)p.Id);

        // Enumerate all visible windows and collect titles for those PIDs
        var titles = new List<string>();
        EnumWindows((hWnd, _) =>
        {
            if (!IsWindowVisible(hWnd)) return true;

            GetWindowThreadProcessId(hWnd, out uint pid);
            if (!pidSet.Contains(pid)) return true;

            var sb = new StringBuilder(256);
            GetWindowText(hWnd, sb, 256);
            string title = sb.ToString();
            if (!string.IsNullOrWhiteSpace(title))
                titles.Add(title);

            return true;
        }, System.IntPtr.Zero);

        return true;
    }
}

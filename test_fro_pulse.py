"""
Standalone FRO + TriggerBox test script.
No GUI required. Run from terminal:

    python test_fro_pulse.py

Tests:
  1  SinglePulse.vbs  — template, no modification
  2  DoublePulse.vbs  — template, no modification  (verifies hardware double-pulse)
  3  DoublePulse.vbs  — Output 2 modified to 0.0501 + ISI_ms/1000
  4  Print decoded FRO values (no fire)
  q  Quit
"""

import os
import re
import sys
import time
import struct

# ── Optional serial ────────────────────────────────────────────────────────────
try:
    import serial
    _SERIAL_OK = True
except ImportError:
    _SERIAL_OK = False
    print("[warn] pyserial not available — TriggerBox disabled")

# ── Optional LabChart COM ──────────────────────────────────────────────────────
try:
    import win32com.client
    _COM_OK = True
except ImportError:
    _COM_OK = False
    print("[warn] win32com not available — LabChart PlayMessage disabled")

TRIGGERBOX_PORT    = "COM5"
TRIGGERBOX_BAUD    = 115200
TRIGGERBOX_CHANNEL = 1        # ch1 = 0x01; wired to LabChart ch5
TRIGGERBOX_PULSE_S = 0.100    # 100 ms pulse

VBS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reference", "FRO")


# ── Helpers ────────────────────────────────────────────────────────────────────

def load_vbs_hex(filename):
    path = os.path.join(VBS_DIR, filename)
    with open(path, "r") as f:
        content = f.read()
    m = re.search(r'PlayMessage\s*\("(0x[0-9A-Fa-f]+)"\)', content)
    if not m:
        raise ValueError(f"PlayMessage hex not found in {filename}")
    return m.group(1)


def set_fro_output_delay(hex_str, output_num, delay_s):
    raw = bytearray(bytes.fromhex(hex_str[2:]))
    orig_sum      = sum(raw)
    output_marker = f"Output = {output_num}\n".encode("utf-16-le")
    key           = "PulseDelay = ".encode("utf-16-le")
    newline_le    = b"\x0a\x00"

    pos = raw.find(bytes(output_marker))
    if pos == -1:
        raise ValueError(f"Output = {output_num} not found in hex")
    pd_pos = raw.find(bytes(key), pos)
    if pd_pos == -1:
        raise ValueError(f"PulseDelay not found after Output = {output_num}")

    val_start = pd_pos + len(key)
    val_end   = raw.find(newline_le, val_start)
    if val_end == -1:
        raise ValueError("newline not found after PulseDelay value")

    old_val = raw[val_start:val_end].decode("utf-16-le")
    old_len = val_end - val_start
    new_val = f"{delay_s:.4f}".encode("utf-16-le")
    new_len = len(new_val)

    if new_len < old_len:
        new_val = new_val + b"\x20\x00" * ((old_len - new_len) // 2)
    elif new_len > old_len:
        new_val = new_val[:old_len]

    raw[val_start:val_end] = new_val

    delta    = sum(raw) - orig_sum
    old_csum = int.from_bytes(raw[20:24], "little")
    new_csum = (old_csum + delta) % 0x100000000
    raw[20:24] = new_csum.to_bytes(4, "little")

    print(f"  [hex] Output {output_num}: {old_val.strip()} → {delay_s:.4f}  (checksum delta={delta:+d})")
    return "0x" + raw.hex().upper()


def decode_fro_outputs(hex_str):
    """Print all active outputs and their PulseDelay values."""
    try:
        raw = bytes.fromhex(hex_str[2:])
        text = raw.decode("utf-16-le", errors="replace")
    except Exception as e:
        print(f"  [decode error] {e}")
        return
    for line in text.splitlines():
        line = line.strip()
        if line.startswith("Output = ") or line.startswith("On = ") or line.startswith("PulseDelay = "):
            print(f"    {line}")


def open_triggerbox():
    if not _SERIAL_OK:
        return None
    try:
        tb = serial.Serial(TRIGGERBOX_PORT, TRIGGERBOX_BAUD, timeout=0.1)
        tb.write(bytes([0]))
        tb.flush()
        print(f"  [triggerbox] opened {TRIGGERBOX_PORT}")
        return tb
    except Exception as e:
        print(f"  [triggerbox] open failed: {e}")
        return None


def fire_triggerbox(tb):
    if tb is None or not tb.is_open:
        print("  [triggerbox] not open — skipping fire")
        return
    ch_byte = 1 << (TRIGGERBOX_CHANNEL - 1)
    tb.write(bytes([ch_byte]))
    tb.flush()
    print(f"  [triggerbox] fired ch{TRIGGERBOX_CHANNEL} (0x{ch_byte:02X})")
    time.sleep(TRIGGERBOX_PULSE_S)
    tb.write(bytes([0]))
    tb.flush()
    print(f"  [triggerbox] reset")


def get_labchart_doc():
    if not _COM_OK:
        return None
    try:
        app = win32com.client.GetActiveObject("ADIChart.Application")
        doc = app.ActiveDocument
        print("  [labchart] connected to active document")
        return doc
    except Exception as e:
        print(f"  [labchart] connect failed: {e}")
        return None


def play_message(doc, hex_str, label=""):
    tag = f" ({label})" if label else ""
    try:
        app = win32com.client.GetActiveObject("ADIChart.Application")
        doc = app.ActiveDocument
        doc.PlayMessage(hex_str)
        print(f"  [labchart] PlayMessage sent{tag}")
    except Exception as e:
        print(f"  [labchart] PlayMessage FAILED{tag}: {e}")
        print("  → LabChart가 실행 중이고 sampling 중인지 확인하세요.")


# ── Test sequences ─────────────────────────────────────────────────────────────

def test_single_template(doc, tb):
    print("\n--- Test 1: SinglePulse template (no modification) ---")
    hex_str = load_vbs_hex("SinglePulse.vbs")
    print("  Decoded FRO:")
    decode_fro_outputs(hex_str)
    play_message(doc, hex_str, "SinglePulse template")
    time.sleep(0.05)
    fire_triggerbox(tb)


def test_double_template(doc, tb):
    print("\n--- Test 2: DoublePulse template (no modification) ---")
    hex_str = load_vbs_hex("DoublePulse.vbs")
    print("  Decoded FRO:")
    decode_fro_outputs(hex_str)
    play_message(doc, hex_str, "DoublePulse template")
    time.sleep(0.05)
    fire_triggerbox(tb)


def test_double_modified(doc, tb, isi_ms):
    isi_s  = isi_ms / 1000.0
    out2_s = 0.0501 + isi_s
    print(f"\n--- Test 3: DoublePulse modified  ISI={isi_ms} ms  Out2={out2_s:.4f}s ---")
    hex_str  = load_vbs_hex("DoublePulse.vbs")
    modified = set_fro_output_delay(hex_str, 2, out2_s)
    print("  Decoded FRO (modified):")
    decode_fro_outputs(modified)
    play_message(doc, modified, f"DoublePulse ISI={isi_ms}ms")
    time.sleep(0.05)
    fire_triggerbox(tb)


def test_double_two_step(doc, tb):
    """Two-step: template → 50ms → template again → 50ms → fire."""
    print("\n--- Test 5: DoublePulse two-step (template → template) ---")
    hex_str = load_vbs_hex("DoublePulse.vbs")
    play_message(doc, hex_str, "step 1: template")
    time.sleep(0.05)
    play_message(doc, hex_str, "step 2: template again")
    time.sleep(0.05)
    fire_triggerbox(tb)


def test_alternate_single_double(doc, tb, n=10, isi_ms=2.5, delay_s=5.0):
    """Alternate SinglePulse and DoublePulse to reproduce PowerLab disconnect."""
    sp_hex = load_vbs_hex("SinglePulse.vbs")
    dp_hex = load_vbs_hex("DoublePulse.vbs")
    isi_s  = isi_ms / 1000.0
    dp_mod = set_fro_output_delay(dp_hex, 2, 0.0501 + isi_s)
    print(f"\n--- Test 6: Alternate Single/Double x{n}  ISI={isi_ms}ms  interval={delay_s}s ---")
    for i in range(n):
        is_single = (i % 2 == 0)
        label = "Single" if is_single else f"Double ISI={isi_ms}ms"
        hex_str = sp_hex if is_single else dp_mod
        print(f"  [{i+1}/{n}] {label}")
        play_message(doc, hex_str, label)
        time.sleep(0.15)
        fire_triggerbox(tb)
        print(f"  waiting {delay_s}s...")
        time.sleep(delay_s)
    print("  Done.")


def test_decode_only():
    print("\n--- Test 4: Decode only (no fire) ---")
    for fname in ("SinglePulse.vbs", "DoublePulse.vbs"):
        print(f"\n  {fname}:")
        try:
            h = load_vbs_hex(fname)
            decode_fro_outputs(h)
        except Exception as e:
            print(f"    error: {e}")


# ── Main ───────────────────────────────────────────────────────────────────────

def start_sampling(doc):
    if doc is None:
        return
    try:
        doc.StartSampling(0, False, 0)
        print("  [labchart] sampling started")
    except Exception as e:
        print(f"  [labchart] StartSampling failed: {e}")


def stop_sampling(doc):
    if doc is None:
        return
    try:
        doc.StopSampling()
        print("  [labchart] sampling stopped")
    except Exception as e:
        print(f"  [labchart] StopSampling failed: {e}")


def main():
    print("FRO + TriggerBox test script")
    print("=" * 40)

    doc = get_labchart_doc()
    tb  = open_triggerbox()

    MENU = """
Options:
  1  SinglePulse template (no modification)
  2  DoublePulse template (no modification)
  3  DoublePulse modified  (enter ISI in ms)
  4  Decode hex values only (no fire)
  5  DoublePulse two-step (template → template → fire)
  6  Alternate Single/Double (reproduce disconnect bug)
  s  Start sampling (재생)
  x  Stop sampling (정지)
  q  Quit
"""

    while True:
        print(MENU)
        choice = input("Choice: ").strip().lower()

        if choice == "q":
            break
        elif choice == "s":
            start_sampling(doc)
        elif choice == "x":
            stop_sampling(doc)
        elif choice == "1":
            test_single_template(doc, tb)
        elif choice == "2":
            test_double_template(doc, tb)
        elif choice == "3":
            try:
                isi = float(input("  ISI (ms): ").strip())
            except ValueError:
                print("  Invalid number.")
                continue
            test_double_modified(doc, tb, isi)
        elif choice == "5":
            test_double_two_step(doc, tb)
        elif choice == "6":
            try:
                n   = int(input("  Repeat count (default 10): ").strip() or "10")
                isi = float(input("  ISI ms (default 2.5): ").strip() or "2.5")
                sec = float(input("  Interval between pulses sec (default 5): ").strip() or "5")
            except ValueError:
                print("  Invalid input.")
                continue
            test_alternate_single_double(doc, tb, n=n, isi_ms=isi, delay_s=sec)
        elif choice == "4":
            test_decode_only()
        else:
            print("  Unknown choice.")

    stop_sampling(doc)

    if tb and tb.is_open:
        tb.write(bytes([0]))
        tb.flush()
        tb.close()
        print("[triggerbox] closed")

    print("Done.")


if __name__ == "__main__":
    main()

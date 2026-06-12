"""
Magstim BiStim serial communication test script.
Run this standalone to verify connection before integrating into TMSviewer.

Requirements:
    pip install pyserial

Usage:
    python magstim_test.py        (prompts for COM port)
    python magstim_test.py COM3   (specify port directly)
"""

import sys
import time
import threading
import serial


# ---------------------------------------------------------------------------
# CRC
# ---------------------------------------------------------------------------

def calc_crc(command: str) -> str:
    """Single CRC byte: bitwise NOT of (sum of ASCII values & 0xFF)."""
    return chr(~sum(ord(c) for c in command) & 0xFF)


def build_cmd(command: str) -> bytes:
    return (command + calc_crc(command)).encode("latin-1")


# ---------------------------------------------------------------------------
# Low-level send / receive
# ---------------------------------------------------------------------------

def send_recv(port: serial.Serial, command: str, expected_bytes: int) -> bytes | None:
    """Send a command and read back expected_bytes. Returns None on error."""
    try:
        port.reset_input_buffer()
        port.write(build_cmd(command))
        response = port.read(expected_bytes)
        return response if len(response) == expected_bytes else None
    except serial.SerialException as e:
        print(f"  [serial error] {e}")
        return None


# ---------------------------------------------------------------------------
# Commands
# ---------------------------------------------------------------------------

def enable_remote(port: serial.Serial) -> bool:
    resp = send_recv(port, "Q@", 3)
    return resp is not None and resp[0:1] != b"?"


def disable_remote(port: serial.Serial):
    send_recv(port, "R@", 3)


def keepalive(port: serial.Serial) -> bool:
    resp = send_recv(port, "Q@", 3)
    return resp is not None


def get_parameters(port: serial.Serial) -> dict | None:
    """
    J@ query — returns PowerA, PowerB, IPI and status flags.
    Response: 1 echo byte (J) + 1 status + 3 PowerA + 3 PowerB + 3 IPI + 1 CRC = 12 bytes total.
    """
    resp = send_recv(port, "J@", 12)
    if resp is None:
        return None
    try:
        status = resp[1]
        power_a = int(resp[2:5].decode("latin-1"))
        power_b = int(resp[5:8].decode("latin-1"))
        ipi     = int(resp[8:11].decode("latin-1"))
        flags = {
            "standby":        bool(status & 0x01),
            "armed":          bool(status & 0x02),
            "ready":          bool(status & 0x04),
            "coil_present":   bool(status & 0x08),
            "replace_coil":   bool(status & 0x10),
            "error":          bool(status & 0x20),
            "remote_enabled": bool(status & 0x80),
        }
        return {"power_a": power_a, "power_b": power_b, "ipi": ipi, "flags": flags}
    except (ValueError, UnicodeDecodeError):
        return None


def get_temperature(port: serial.Serial) -> dict | None:
    """
    F@ query — returns coil temperatures.
    Response: 1 echo (F) + 1 status + 3 Temp1 + 3 Temp2 + 1 CRC = 9 bytes.
    """
    resp = send_recv(port, "F@", 9)
    if resp is None:
        return None
    try:
        temp1 = int(resp[2:5].decode("latin-1"))
        temp2 = int(resp[5:8].decode("latin-1"))
        return {"temp1": temp1, "temp2": temp2}
    except (ValueError, UnicodeDecodeError):
        return None


def set_power_a(port: serial.Serial, value: int) -> bool:
    """Set Channel A power (0–100)."""
    cmd = f"@{value:03d}"
    resp = send_recv(port, cmd, 3)
    return resp is not None and resp[0:1] != b"?"


def set_power_b(port: serial.Serial, value: int) -> bool:
    """Set Channel B power (0–100)."""
    cmd = f"A{value:03d}"
    resp = send_recv(port, cmd, 3)
    return resp is not None and resp[0:1] != b"?"


def set_ipi(port: serial.Serial, value: int) -> bool:
    """Set inter-pulse interval (0–999 in normal mode)."""
    cmd = f"C{value:03d}"
    resp = send_recv(port, cmd, 3)
    return resp is not None and resp[0:1] != b"?"


# ---------------------------------------------------------------------------
# Keepalive thread
# ---------------------------------------------------------------------------

class KeepaliveThread(threading.Thread):
    def __init__(self, port: serial.Serial):
        super().__init__(daemon=True)
        self.port    = port
        self._stop   = threading.Event()
        self.healthy = True

    def run(self):
        while not self._stop.wait(0.5):
            self.healthy = keepalive(self.port)

    def stop(self):
        self._stop.set()


# ---------------------------------------------------------------------------
# Interactive test CLI
# ---------------------------------------------------------------------------

HELP = """
Commands:
  status          — read PowerA, PowerB, IPI, status flags
  a <0-100>       — set Channel A power
  b <0-100>       — set Channel B power
  ipi <0-999>     — set inter-pulse interval
  poll            — start polling status every 1s (press Enter to stop)
  help            — show this message
  quit            — release remote control and exit
"""


def print_params(params: dict):
    f = params["flags"]
    state = ("ARMED/READY" if f["ready"] else
             "ARMED" if f["armed"] else
             "STANDBY" if f["standby"] else "UNKNOWN")
    print(f"  Power A : {params['power_a']:3d}%")
    print(f"  Power B : {params['power_b']:3d}%")
    print(f"  IPI     : {params['ipi']:3d}")
    print(f"  State   : {state}")
    print(f"  Remote  : {'ON' if f['remote_enabled'] else 'OFF'}")
    if f["error"]:
        print("  *** DEVICE ERROR FLAG SET ***")
    if f["replace_coil"]:
        print("  *** REPLACE COIL FLAG SET ***")


def run_interactive(port: serial.Serial):
    print("\nEnabling remote control...", end=" ", flush=True)
    if not enable_remote(port):
        print("FAILED — check device is on and not in local-only mode.")
        return
    print("OK")

    ka = KeepaliveThread(port)
    ka.start()
    print("Keepalive thread started (500 ms interval).")
    print(HELP)

    try:
        while True:
            try:
                line = input("bistim> ").strip()
            except EOFError:
                break

            if not line:
                continue

            parts = line.split()
            cmd   = parts[0].lower()

            if cmd in ("quit", "exit", "q"):
                break

            elif cmd == "help":
                print(HELP)

            elif cmd == "status":
                params = get_parameters(port)
                if params:
                    print_params(params)
                else:
                    print("  Failed to read parameters.")

            elif cmd == "a" and len(parts) == 2:
                try:
                    val = int(parts[1])
                    assert 0 <= val <= 100
                except (ValueError, AssertionError):
                    print("  Usage: a <0-100>")
                    continue
                ok = set_power_a(port, val)
                print(f"  Power A → {val}%  {'OK' if ok else 'FAILED'}")

            elif cmd == "b" and len(parts) == 2:
                try:
                    val = int(parts[1])
                    assert 0 <= val <= 100
                except (ValueError, AssertionError):
                    print("  Usage: b <0-100>")
                    continue
                ok = set_power_b(port, val)
                print(f"  Power B → {val}%  {'OK' if ok else 'FAILED'}")

            elif cmd == "ipi" and len(parts) == 2:
                try:
                    val = int(parts[1])
                    assert 0 <= val <= 999
                except (ValueError, AssertionError):
                    print("  Usage: ipi <0-999>")
                    continue
                ok = set_ipi(port, val)
                print(f"  IPI → {val}  {'OK' if ok else 'FAILED'}")

            elif cmd == "poll":
                print("  Polling every 1s — press Enter to stop.")
                stop_poll = threading.Event()

                def _poll():
                    while not stop_poll.wait(1.0):
                        params = get_parameters(port)
                        if params:
                            print(f"\r  A={params['power_a']:3d}%  "
                                  f"B={params['power_b']:3d}%  "
                                  f"IPI={params['ipi']:3d}  "
                                  f"state={'RDY' if params['flags']['ready'] else 'SBY'}  ",
                                  end="", flush=True)

                t = threading.Thread(target=_poll, daemon=True)
                t.start()
                input()
                stop_poll.set()
                t.join(timeout=2)
                print()

            else:
                print(f"  Unknown command: '{line}'. Type 'help'.")

    finally:
        ka.stop()
        print("\nDisabling remote control...", end=" ", flush=True)
        disable_remote(port)
        print("done.")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    if len(sys.argv) >= 2:
        port_name = sys.argv[1]
    else:
        port_name = input("Enter COM port (e.g. COM3): ").strip()

    print(f"\nOpening {port_name} at 9600 baud...", end=" ", flush=True)
    try:
        port = serial.Serial(
            port=port_name,
            baudrate=9600,
            bytesize=serial.EIGHTBITS,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE,
            timeout=0.3,
        )
    except serial.SerialException as e:
        print(f"FAILED\n  {e}")
        sys.exit(1)
    print("OK")

    try:
        run_interactive(port)
    finally:
        port.close()
        print("Port closed.")


if __name__ == "__main__":
    main()

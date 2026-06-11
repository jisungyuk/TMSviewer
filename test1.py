import time
import threading
import pythoncom
import win32com.client
import numpy as np
import matplotlib.pyplot as plt

CHANNEL = 3  # Ch3

lc_doc = None


def read_channel_data():
    try:
        num_blocks = lc_doc.NumberOfRecords
        secs_per_tick = lc_doc.GetRecordSecsPerTick(num_blocks)
        sample_rate = round(1 / secs_per_tick)

        raw = lc_doc.GetScopeChannelData(1, CHANNEL, num_blocks, 1, -1, 1)
        data = np.array(raw)
        units = lc_doc.GetUnits(CHANNEL, 1)

        total_samples = len(data)
        time = np.arange(total_samples) / sample_rate * 1000  # ms

        print(f"  샘플 수: {total_samples}, 샘플레이트: {sample_rate} Hz, 단위: {units}")
        print(f"  Peak-to-peak: {data.max() - data.min():.4f} {units}")

        plt.figure(figsize=(10, 4))
        plt.plot(time, data)
        plt.xlabel("Time (ms)")
        plt.ylabel(units)
        plt.title(f"Ch{CHANNEL} - TMS 발화 후 EMG (Block {num_blocks})")
        plt.tight_layout()
        plt.show(block=False)
        plt.pause(0.1)

    except Exception as e:
        print(f"  [오류] 데이터 읽기 실패: {e}")


class LabChartEvents:

    def OnCommentAdded(self, *args):
        print("[TMS 감지] Comment 추가됨 -> 1초 후 Ch3 데이터 읽기...")
        threading.Timer(1.0, read_channel_data).start()

    def OnStartSamplingBlock(self, *_):
        print("[블록 시작] 샘플링 시작")

    def OnNewSamples(self, *_):
        pass


def main():
    global lc_doc
    pythoncom.CoInitialize()

    print("LabChart 연결 중...")
    try:
        lc_app = win32com.client.GetActiveObject("ADIChart.Application")
    except Exception:
        print("[오류] LabChart가 실행 중이지 않습니다. LabChart를 먼저 켜주세요.")
        return

    lc_doc = lc_app.ActiveDocument
    if lc_doc is None:
        print("[오류] LabChart에 열린 문서가 없습니다.")
        return

    print(f"연결 성공: {lc_doc.Name}")
    print(f"현재 블록 수: {lc_doc.NumberOfRecords}")
    print("TMS 발화 대기 중... (Ctrl+C로 종료)\n")

    win32com.client.WithEvents(lc_doc, LabChartEvents)

    try:
        while True:
            pythoncom.PumpWaitingMessages()
            time.sleep(0.05)
    except KeyboardInterrupt:
        print("\n종료합니다.")


if __name__ == "__main__":
    main()

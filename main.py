import sys
from PyQt5.QtWidgets import QApplication
from realtime_viewer import RealTimeViewer


def main():
    app = QApplication(sys.argv)
    viewer = RealTimeViewer()
    viewer.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()

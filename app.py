import os
import os.path
import platform
from subprocess import Popen, PIPE
from threading import Thread

from PyQt5.QtCore import *
from PyQt5.QtWidgets import *

label: QLabel
wnd: QMainWindow
currentFilename: str = None
dlg: QProgressDialog = None


def loadFile(filename):
    if filename is None:
        return
    if platform == "Windows":
        filename = filename[1:]
    global currentFilename
    currentFilename = filename
    label.setText("当前文件: " + filename)
    wnd.statusBar().showMessage("文件载入成功: " + filename)


class Converter(QThread):
    done = pyqtSignal(bool)

    def __init__(self, fmt):
        super().__init__()
        self.fmt = fmt

    def run(self):
        fmt = self.fmt
        fromFilename = currentFilename
        toFilename = os.path.splitext(fromFilename)[0] + "." + fmt
        args = ["pandoc", fromFilename]
        if fmt == "pdf":
            args.extend(["--pdf-engine", "wkhtmltopdf"])
        args.extend(["-o", toFilename])
        process = Popen(args=args, stdout=PIPE, stderr=PIPE)
        err = process.communicate()[1].decode()
        code = process.returncode
        wnd.statusBar().showMessage(("转换成功: " + toFilename) if code == 0 else "转换失败: " + err)
        self.done.emit(True)


def convertFile(fmt):
    if currentFilename is None:
        return
    global dlg
    dlg = QProgressDialog()
    dlg.setWindowTitle("转换中...")
    dlg.setMaximum(0)
    dlg.setCancelButton(None)
    dlg.setWindowFlags(Qt.Window | Qt.WindowTitleHint | Qt.CustomizeWindowHint)
    dlg.setWindowModality(Qt.WindowModal)
    dlg.show()
    wnd.statusBar().showMessage("转换中...")
    global converter
    converter = Converter(fmt)
    converter.done.connect(dlg.close)
    converter.start()


app = QApplication([])
wnd = QMainWindow()
wnd.resize(400, 400 * 0.618)
wnd.setAcceptDrops(True)
wnd.dragEnterEvent = lambda e: e.accept()
wnd.dropEvent = lambda e: print(e.mimeData().urls())
wnd.dropEvent = lambda e: loadFile(next(iter([x.path() for x in e.mimeData().urls()]), None))
wnd.statusBar().showMessage("就绪.")
payload = QWidget()
label = QLabel("请拖入文件.转换后会自动覆盖同名文件,请谨慎使用")

buttons = QHBoxLayout()
mapper = QSignalMapper()
for k, v in dict(docx="Word", pdf="PDF", pptx="PowerPoint").items():
    button = QPushButton(v)
    button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
    mapper.setMapping(button, k)
    button.clicked.connect(mapper.map)
    buttons.addWidget(button)
mapper.mapped["QString"].connect(convertFile)

layout = QVBoxLayout()
layout.addWidget(label)
layout.addLayout(buttons)
payload.setLayout(layout)
wnd.setCentralWidget(payload)
wnd.setWindowTitle("麦多文档转换器")
wnd.show()
font = app.font()
font.setPointSize(font.pointSize() / 0.618)
app.setFont(font)
app.exec()

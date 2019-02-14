import os
import os.path
import platform
from subprocess import Popen, PIPE

from PyQt5.QtCore import *
from PyQt5.QtWidgets import *

label: QLabel
wnd: QMainWindow
currentFilename: str = None


def loadFile(filename):
    if filename is None:
        return
    if platform == "Windows":
        filename = filename[1:]
    global currentFilename
    currentFilename = filename
    label.setText("当前文件: " + filename)
    wnd.statusBar().showMessage("文件载入成功: " + filename)


def convertFile(fmt):
    if currentFilename is None:
        return
    wnd.statusBar().showMessage("转换中...")
    fromFilename = currentFilename
    toFilename = os.path.splitext(fromFilename)[0] + "." + fmt
    args = ["pandoc", fromFilename]
    if fmt == "pdf":
        args.extend(["--pdf-engine", "wkhtmltopdf"])
    args.extend(["-o", toFilename])
    process = Popen(args=args, stdout=PIPE, stderr=PIPE)
    _, err = process.communicate()
    code = process.returncode
    wnd.statusBar().showMessage(("转换成功: " + toFilename) if code == 0 else "转换失败: " + err.decode())


app = QApplication([])
wnd = QMainWindow()
wnd.resize(400, 400 * 0.618)
wnd.setAcceptDrops(True)
wnd.dragEnterEvent = lambda e: e.accept()
wnd.dropEvent = lambda e: print(e.mimeData().urls())
wnd.dropEvent = lambda e: loadFile(next(iter([x.path() for x in e.mimeData().urls()]), None))
wnd.statusBar().showMessage("就绪.")
payload = QWidget()
label = QLabel("请拖入文件.")

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
wnd.setWindowTitle("麦多")
wnd.show()
font = app.font()
font.setPointSize(font.pointSize() * 2)
app.setFont(font)
app.exec()

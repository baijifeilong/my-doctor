import os.path

import pypandoc
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *

label: QLabel
wnd: QMainWindow
currentFilename: str = None


def loadFile(filename):
    global currentFilename
    currentFilename = filename
    print("Done")
    label.setText("当前文件: " + filename)
    wnd.statusBar().showMessage("文件载入成功: " + filename)


def convertFile(fmt):
    if currentFilename is None:
        return
    wnd.statusBar().showMessage("转换中...")
    wnd.repaint()
    filename = currentFilename
    to = os.path.splitext(filename)[0] + "." + fmt
    pypandoc.convert_file(filename, fmt, outputfile=to)
    import time
    time.sleep(3)
    wnd.statusBar().showMessage("转换成功: " + to)


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
wnd.show()
font = app.font()
font.setPointSize(font.pointSize() * 2)
app.setFont(font)
app.exec()

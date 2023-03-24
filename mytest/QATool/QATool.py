import whisper
import sys
from PySide6.QtWidgets import QWidget, QApplication, QLineEdit, QMainWindow, QTextBrowser, QPushButton, QMenu
import os
import difflib
import openpyxl
import zhconv

class Window(QMainWindow):
    def __init__(self):
        super(Window, self).__init__()
        self.text = ""  # ==> 默认文本内容
        self.setWindowTitle('CV QA Assistant')  # ==> 窗口标题
        self.resize(500, 400)  # ==> 定义窗口大小
        self.textBrowser = QTextBrowser()
        self.setCentralWidget(self.textBrowser)  # ==> 定义窗口主题内容为textBrowser
        self.setAcceptDrops(True)  # ==> 设置窗口支持拖动（必须设置）
        #self.model = whisper.load_model("medium")
        self.model = whisper.load_model("small")
        self.ModelLevel = 'small'
        self.Sheetname = 'Sheet1'
        self.TextCol = 'A'
        self.WavCol = 'B'
        self.textBrowser.setText('drag audio files folder in to set audio path\n'+'fill out the blanks below before drag in .xlsx please\n')
        self.audiopath = ''
        self.gate = '0.9'
        self.initUI()

    # 鼠标拖入事件
    def dragEnterEvent(self, event):
        file = event.mimeData().urls()[0].toLocalFile()
        if file.endswith('.xlsx'):
            event.accept()
        else:
            self.audiopath = file
            self.textBrowser.setText('audio files path set to: '+ self.audiopath +'\n'+'now drag in xlsx file please!!!\n')
            event.ignore()

    # 鼠标放开
    def dropEvent(self, event):
        #self.setWindowTitle('FindingSimilarAudio')
        file = event.mimeData().urls()[0].toLocalFile()  # ==> 获取文件路径
        #print("拖拽的文件 ==> {}".format(file))
        #self.text += file + "\n"
        #self.textBrowser.setText(self.text)
        if file.endswith('.xlsx'):
            self.text = 'audio files path set to: '+ self.audiopath +'\n'
            self.wb = openpyxl.load_workbook(file)
            self.text += 'sheet: '+self.Sheetname + ' opened'
            #self.textBrowser.setText('sheet: '+self.Sheetname + ' opened')
        else:
            self.text = 'drag in xlsx file please!!!'

        self.textBrowser.setText(self.text)
        event.accept()#事件处理完毕,不向上转发,ignore()则向上转发

    def initUI(self):
        CModelBtn = QPushButton('Start checking', self)
        CModelBtn.setCheckable(True)
        CModelBtn.move(10, 350)
        PlusBtn = QPushButton('change model', self)
        PlusBtn.setCheckable(True)
        PlusBtn.move(110, 350)
        #MinusBtn = QPushButton('-', self)
        #MinusBtn.setCheckable(True)
        #MinusBtn.move(210, 350)
        CModelBtn.clicked[bool].connect(self.checkCV)
        PlusBtn.clicked[bool].connect(self.PlusNum)
        #MinusBtn.clicked[bool].connect(self.MinusNum)
        #self.setGeometry(300, 300, 280, 170)
        #self.setWindowTitle('切换按钮') 
        #self.show()
        TextColwidget = QLineEdit(self)
        # 设置最长输入字符10个
        TextColwidget.setMaxLength(10)
        # 输入框提示
        TextColwidget.setPlaceholderText("Text Col, like A, B or CD etc.")
        # 设置输入框只读
        # widget.setReadOnly(True)
        # 按下回车键
        #widget.returnPressed.connect(self.return_pressed)
        # 鼠标选中
        #widget.selectionChanged.connect(self.selection_changed)
        # 文本发生变化
        TextColwidget.textChanged.connect(self.TextColtext_changed)
        # 正在编辑
        #widget.textEdited.connect(self.text_edited)
        TextColwidget.move(10, 300)
        
        CVColwidget = QLineEdit(self)
        # 设置最长输入字符10个
        CVColwidget.setMaxLength(10)
        # 输入框提示
        CVColwidget.setPlaceholderText("CV Col, like A, B or CD etc.")
        CVColwidget.textChanged.connect(self.CVColtext_changed)
        # 正在编辑
        #widget.textEdited.connect(self.text_edited)
        CVColwidget.move(110, 300)

        Sheetwidget = QLineEdit(self)
        # 设置最长输入字符10个
        Sheetwidget.setMaxLength(10)
        # 输入框提示
        Sheetwidget.setPlaceholderText("sheet name")
        Sheetwidget.textChanged.connect(self.Sheettext_changed)
        # 正在编辑
        #widget.textEdited.connect(self.text_edited)
        Sheetwidget.move(210, 300)

        Gatewidget = QLineEdit(self)
        # 设置最长输入字符10个
        Gatewidget.setMaxLength(10)
        # 输入框提示
        Gatewidget.setPlaceholderText("accuracy gate, default 0.9")
        Gatewidget.textChanged.connect(self.gatechange)
        # 正在编辑
        #widget.textEdited.connect(self.text_edited)
        Gatewidget.move(310, 300)
        
    def checkCV(self, pressed):
        if self.audiopath == '':
            self.textBrowser.setText('drag in folder to set the audio path first!!!')
            return
        self.text = 'audio files path set to: '+ self.audiopath +'\n'
        self.AudioSheet = self.wb[self.Sheetname]
        for i in range(self.AudioSheet.max_row):
            if (self.AudioSheet[self.TextCol + str(i+1)].value != '') and (self.AudioSheet[self.WavCol + str(i+1)].value != ''):
                path = self.AudioSheet[self.WavCol + str(i+1)].value
                result = self.model.transcribe(os.path.join(self.audiopath, path))
                #print(result)
                #print(self.AudioSheet[self.TextCol + str(i+1)].value)
                ratio = self.Compare(self.AudioSheet[self.TextCol + str(i+1)].value, result["text"])
                if ratio < float(self.gate):
                    self.text += 'mismatching'+self.AudioSheet[self.TextCol + str(i+1)].value+'\n'+'xlsx line '+ str(i+1)+'\n'
                    self.textBrowser.setText(self.text)
        self.text += 'completed!!!'
        self.textBrowser.setText(self.text)

        #result = model.transcribe(r"G:\SpeechRecognition\whisper\mytest\untitled_1.wav")
        #print(result["text"])

    def PlusNum(self, pressed):
        if self.ModelLevel == 'medium':
            self.ModelLevel = 'small'
        else:
            self.ModelLevel = 'medium'
        self.model = whisper.load_model(self.ModelLevel)
        self.textBrowser.setText('model changed to '+self.ModelLevel+'\n'+'简中请不要使用medium!!!\n'+'medium与繁体台词配合使用')


    def TextColtext_changed(self, s):
        self.TextCol = s

    def CVColtext_changed(self, s):
        self.WavCol = s

    def Sheettext_changed(self, s):
        self.Sheetname = s

    def gatechange(self, s):
        self.gate = s

    def Compare(self, A,B):
        #if self.is_chinese(A[0]) and self.ModelLevel == 'medium':
            #self.textBrowser.setText('aaaaaaaaa')
            #B = zhconv.convert(B, 'zh-cn')
            #self.textBrowser.setText('bbbbbbbbbbbb')
        return difflib.SequenceMatcher(None, A, B).quick_ratio()

    def is_chinese(self, char):
        if '\u4e00' <= char <= '\u9fff':
            return True
        else:
            return False


'''
    def return_pressed(self):
        print("Return pressed!")
        #self.centralWidget().setText("Boom!")

    def selection_changed(self):
        print("Selection changed!")
        #print(self.centralWidget().selectedText())

    def text_edited(self, s):
        print("Text edited...")
        #print(s)
'''

app = QApplication(sys.argv)
window = Window()
window.show()
sys.exit(app.exec())
import sys
import os
from PyQt5.QtWidgets import QMainWindow, QApplication, qApp, QAction, QFileDialog, QTextEdit
from PyQt5.QtGui import QIcon
from openpyxl import load_workbook
import datetime
import win32com.client as win32


class App(QMainWindow):
    def __init__(self):
        super().__init__()
        self.console = QTextEdit()
        self.headers = []
        self.contents = []
        self.outlook = win32.Dispatch('outlook.application')
        self.initAct()
        self.initUI()
        # for debug only, turn off line below for product use
        self.parse_input_excel('./assets/input_example.xlsx')

    def initAct(self):
        exitAct = QAction('&Exit', self)
        exitAct.setShortcut('Ctrl+Q')
        exitAct.setStatusTip('Exit application')
        exitAct.triggered.connect(qApp.quit)
        openAct = QAction(QIcon('assets/excel.ico'), '&Open Excel', self)
        openAct.setShortcut('Ctrl+O')
        openAct.setStatusTip('Open Excel as input')
        openAct.triggered.connect(self.on_open_excel)
        sendAct = QAction(QIcon('assets/send.ico'), '&Send Mails', self)
        sendAct.setShortcut('Ctrl+Shit+S')
        sendAct.setStatusTip('Send mail')
        sendAct.triggered.connect(self.on_send_mail)
        helpAct = QAction('&Help', self)
        helpAct.setShortcut('Ctrl+H')
        helpAct.triggered.connect(self.on_help)

        menubar = self.menuBar()
        fileMenu = menubar.addMenu('&File')
        fileMenu.addAction(openAct)
        fileMenu.addAction(sendAct)
        fileMenu.addAction(exitAct)
        helpMenu = menubar.addMenu('&Help')
        helpMenu.addAction(helpAct)

        self.toolbar = self.addToolBar('toolbar')
        self.toolbar.addAction(openAct)
        self.toolbar.addAction(sendAct)

    def initUI(self):
        self.statusBar().showMessage('Ready')
        self.setGeometry(300, 300, 640, 400)
        self.setWindowTitle('MailBot')
        self.setWindowIcon(QIcon('assets/app.ico'))
        self.setCentralWidget(self.console)
        self.console.setReadOnly(True)
        self.show()

    def log_to_console(self, log_text):
        self.console.insertPlainText(log_text + '\n')

    def on_open_excel(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "Select Excel File",
            os.path.expanduser("~/Desktop"),  # 选择Excel文件对话框默认在桌面
            "Excel File(*.xlsx)"
        )
        if file_name and os.path.splitext(file_name)[-1] == '.xlsx':
            self.parse_input_excel(file_name)

    def parse_input_excel(self, file_name):
        """解析输入的Excel文档
        :param file_name: Excel文档路径
        """
        excel_handle = load_workbook(file_name)  # 打开Excel文档
        worksheet = excel_handle.worksheets[0]  # 默认只支持从第一个页签加载数据
        for idx,row in enumerate(worksheet.rows):
            row_content = [cell.value for cell in row]
            if idx == 0:
                self.parse_header(row_content)
            else:
                self.append_content(row_content)
        self.log_to_console('Ready to send!')

    def parse_header(self, row_content):
        """解析表头行
        :param row_content: 表头行list
        """
        for idx,title in enumerate(row_content):
            if title.strip() == '邮箱':
                self.mail_col_idx = idx
        self.headers = row_content

    def append_content(self, row_content):
        """解析内容行, 并存储
        :param row_content: 内容行list
        """
        self.contents.append((
            row_content[self.mail_col_idx],
            row_content
        ))

    def on_send_mail(self):
        for idx,content in enumerate(self.contents):
            mail = self.outlook.CreateItem(0)
            if not content[0]:
                continue
            mail.To = content[0]
            content_value = content[1]
            for i in range(len(content_value)):
                if type(content_value[i]) is datetime.datetime:
                    content_value[i] = content_value[i].strftime('%Y-%m-%d')
            self.log_to_console('#{}: send to {}'.format(idx+1, mail.To))
            mail.Subject = '【重要请核对】请您核对已有信息，截止时间：2018年11月21日17:00前，谢谢 :)'  # 邮件主题
            # 接下来填入邮件正文(HTML)
            mail.HTMLBody = '<p>各位员工：</p>'
            mail.HTMLBody += '<p>为了进一步提升公司员工关怀的服务质量，请您核对如下个人相关信息，如缺失或错误请补充改正，直接回复此邮件修改即可，谢谢您的支持~</p>'
            mail.HTMLBody += '<table style="table-layout:fixed;" cellpadding="0" cellspacing="0" border="1">'
            mail.HTMLBody += ('<tr>' + '<th style="white-space:nowrap;">{}</th>' * (len(self.headers)-2) + '</tr>').format(*self.headers[:-2])
            mail.HTMLBody += ('<tr>' + '<th style="white-space:nowrap;">{}</th>' * (len(content_value)-2) + '</tr>').format(*content_value[:-2])
            mail.HTMLBody += '</table>'
            mail.HTMLBody += '<p>OnceMore2020</p>'
            mail.Send()
        self.log_to_console('发送完成')

    def on_help(self):
        self.log_to_console('{}\nDark Powered Mail Bot\nAuthor:OnceMore2020\n{}'.format('* + '*10, '* + '*10))

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())
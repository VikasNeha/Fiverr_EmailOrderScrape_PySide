from PySide.QtGui import QApplication, QMainWindow, QMessageBox
from PySide.QtCore import QTimer
import sys
from ui_ScrapeTool import Ui_MainWindow
from config import Config
from datetime import date
from Scraping_Main import doScrape
from multiprocessing import Process, Queue, freeze_support
#import threading
#from Queue import Queue


# noinspection PyCallByClass
class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)
        self.setupUi(self)
        self.btnScrape.clicked.connect(self.doScrape)

        self.q = Queue()
        self.timer = QTimer()
        self.timer.timeout.connect(self.updateProgress)

    def doScrape(self):
        if not self.validateInput():
            return
        Con = self.buildInput()
        QMessageBox.information(self, 'Starting...', 'Starting Scrape, Please be patient !!!')
        self.timer.start(100)
        p = Process(target=doScrape, args=(Con, self.q))
        p.daemon = True
        p.start()

    def updateProgress(self):
        if self.q.empty():
            return
        q_msg, q_progress = self.q.get()
        self.lblProgress.setText(q_msg)
        self.progressBar.setValue(q_progress)

        if q_progress == 100:
            self.timer.stop()
            if q_msg == 'SUCCESS':
                QMessageBox.information(self, 'Done', 'Scrape Successful !!!')
            else:
                QMessageBox.critical(self, 'Error Occurred', 'Error Occurred in Scraping')
        self.updateProgress()

    def validateInput(self):
        if self.txtUsername.text().strip() == '':
            QMessageBox.critical(self, 'Validation Error', 'Please enter User Name')
            self.txtUsername.setFocus()
            return False
        elif self.txtPassword.text().strip() == '':
            QMessageBox.critical(self, 'Validation Error', 'Please enter Password')
            self.txtPassword.setFocus()
            return False
        elif self.txtFromDate.date() > self.txtToDate.date():
            QMessageBox.critical(self, 'Validation Error', 'From Date sould not be greater than To Date')
            self.txtFromDate.setFocus()
            return False
        else:
            return True

    def buildInput(self):
        Con = Config()
        Con.Username = self.txtUsername.text().strip()
        Con.Password = self.txtPassword.text().strip()
        d = self.txtFromDate.date()
        Con.FromDate = date(d.year(), d.month(), d.day())
        d = self.txtToDate.date()
        Con.ToDate = date(d.year(), d.month(), d.day())
        Con.Label = str(self.txtLabel.text().strip())
        if Con.Label == '':
            Con.Label = 'Inbox'
        return Con

if __name__ == '__main__':
    freeze_support()
    app = QApplication(sys.argv)
    frame = MainWindow()
    frame.show()
    sys.exit(app.exec_())
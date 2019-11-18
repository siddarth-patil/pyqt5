import sys
import win32com.client as win32

from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtWebEngineWidgets import *

class MainPage(QWidget):

    def __init__(self):
        super().__init__()
        self.web = QWebEngineView()
        self.initUI()

    def initUI(self):
        geometry = app.desktop().availableGeometry()

        #Button for URL 1
        URL1_button = QPushButton('URL1', self)
        URL1_button.setToolTip('Check URL1')
        URL1_button.clicked.connect(self.URL1_on_click)

        #Button for URL 2
        URL2_button = QPushButton('URL2', self)
        URL2_button.setToolTip('Check URL2')
        URL2_button.clicked.connect(self.URL2_on_click)

        #Button for capturing URL 1
        URL1_Capture = QPushButton('URL1_Capture', self)
        URL1_Capture.setToolTip('Capture URL1')
        URL1_Capture.clicked.connect(self.URL1_Capture)

        #Button for capturing URL 2
        URL2_Capture = QPushButton('URL2_Capture', self)
        URL2_Capture.setToolTip('Capture URL2')
        URL2_Capture.clicked.connect(self.URL2_Capture)

        #Button for sending Mail
        Send_Mail = QPushButton('Send Mail', self)
        Send_Mail.setToolTip('Send Mail')
        Send_Mail.clicked.connect(self.Send_Mail)

        #To divide the whole application into grid form
        grid = QGridLayout()
        grid.setSpacing(10)

        #Positiong the buttons
        grid.addWidget(URL1_button, 1, 0)
        grid.addWidget(URL1_Capture, 1, 1)

        grid.addWidget(URL2_button, 2, 0)
        grid.addWidget(URL2_Capture, 2, 1)

        grid.addWidget(Send_Mail, 3, 0)

        #Positioning the space for the browser with size(30,30)
        grid.addWidget(self.web, 1, 2, 30, 30)

        self.setLayout(grid)

        self.setGeometry(geometry)
        self.setWindowTitle('Try')
        self.show()

    def URL1_on_click(self):
        self.web.load(QUrl("https://google.com"))
        self.web.show()

    def URL2_on_click(self):
        self.web.load(QUrl("https://duckduckgo.com"))
        self.web.show()
    
    #Defining the screenshot function which is invoked on the click of "Url1 Capture" and "Url2 Capture"
    def URL1_Capture(self):
        shot1 = QScreen.grabWindow(app.primaryScreen(), app.desktop().winId())
        shot1.save("C:\\Users\\xyz\\Desktop\\shot1.jpg", 'jpg')
        print("Done")

    def URL2_Capture(self):
        shot2 = QScreen.grabWindow(app.primaryScreen(), app.desktop().winId())
        shot2.save("C:\\Users\\xyz\\Desktop\\shot2.jpg", 'jpg')

    #Function defining Send_Mail which is invoked on the click of "Send Mail" Button
    def Send_Mail(self):
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        # mail.To = 'xyz@gmail.com;abc@gmail.com'
        mail.To = 'xyz@gmail.com'
        mail.CC = "abc@gmail.com"
        mail.BCC = "pqr@gmail.com"
        mail.Subject = 'Test'
        mail.Body = 'Automated Mail using Pyhton \n \n Next Line'
        # mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional

        # To attach a file to the email (optional):
        attachment1 = "C:\\Users\\xyz\\Desktop\\shot1.jpg"
        attachment2 = "C:\\Users\\xyz\\Desktop\\shot2.jpg"
        mail.Attachments.Add(attachment1)
        mail.Attachments.Add(attachment2)

        try:
            mail.Send()
        except:
            print("An exception occurred")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MainPage()
    sys.exit(app.exec_())

# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'EntryWindow.ui'
#
# Created by: PyQt5 UI code generator 5.15.4
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_EntryWindow(object):
    def setupUi(self, EntryWindow):
        EntryWindow.setObjectName("EntryWindow")
        EntryWindow.resize(320, 240)
        self.centralwidget = QtWidgets.QWidget(EntryWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.textBrowser = QtWidgets.QTextBrowser(self.centralwidget)
        self.textBrowser.setGeometry(QtCore.QRect(50, 80, 211, 31))
        self.textBrowser.setObjectName("textBrowser")
        EntryWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(EntryWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 320, 22))
        self.menubar.setObjectName("menubar")
        self.menuFile = QtWidgets.QMenu(self.menubar)
        self.menuFile.setObjectName("menuFile")
        EntryWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(EntryWindow)
        self.statusbar.setObjectName("statusbar")
        EntryWindow.setStatusBar(self.statusbar)
        self.actionNew_Project = QtWidgets.QAction(EntryWindow)
        self.actionNew_Project.setObjectName("actionNew_Project")
        self.actionOpen_Project = QtWidgets.QAction(EntryWindow)
        self.actionOpen_Project.setObjectName("actionOpen_Project")
        self.actionExit = QtWidgets.QAction(EntryWindow)
        self.actionExit.setObjectName("actionExit")
        self.menuFile.addAction(self.actionNew_Project)
        self.menuFile.addAction(self.actionOpen_Project)
        self.menuFile.addAction(self.actionExit)
        self.menubar.addAction(self.menuFile.menuAction())

        self.retranslateUi(EntryWindow)
        QtCore.QMetaObject.connectSlotsByName(EntryWindow)

    def retranslateUi(self, EntryWindow):
        _translate = QtCore.QCoreApplication.translate
        EntryWindow.setWindowTitle(_translate("EntryWindow", "Stock Manager"))
        self.textBrowser.setHtml(_translate("EntryWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'Ubuntu\'; font-size:11pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\">Open or create a new project</p></body></html>"))
        self.menuFile.setTitle(_translate("EntryWindow", "File"))
        self.actionNew_Project.setText(_translate("EntryWindow", "New Project"))
        self.actionOpen_Project.setText(_translate("EntryWindow", "Open Project"))
        self.actionExit.setText(_translate("EntryWindow", "Exit"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    EntryWindow = QtWidgets.QMainWindow()
    ui = Ui_EntryWindow()
    ui.setupUi(EntryWindow)
    EntryWindow.show()
    sys.exit(app.exec_())
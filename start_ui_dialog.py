# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'Start_Dialog.ui'
#
# Created by: PyQt5 UI code generator 5.15.6
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_Dialog(object):
    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(600, 350)
        Dialog.setMinimumSize(QtCore.QSize(600, 350))
        Dialog.setMaximumSize(QtCore.QSize(600, 350))
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(10)
        Dialog.setFont(font)
        self.verticalLayout = QtWidgets.QVBoxLayout(Dialog)
        self.verticalLayout.setObjectName("verticalLayout")
        self.groupBox = QtWidgets.QGroupBox(Dialog)
        self.groupBox.setObjectName("groupBox")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(self.groupBox)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.progressBar = QtWidgets.QProgressBar(self.groupBox)
        self.progressBar.setMinimumSize(QtCore.QSize(0, 30))
        self.progressBar.setMaximumSize(QtCore.QSize(16777215, 30))
        self.progressBar.setProperty("value", 1)
        self.progressBar.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.progressBar.setTextVisible(True)
        self.progressBar.setInvertedAppearance(False)
        self.progressBar.setTextDirection(QtWidgets.QProgressBar.TopToBottom)
        self.progressBar.setObjectName("progressBar")
        self.verticalLayout_2.addWidget(self.progressBar)
        self.label_db0 = QtWidgets.QLabel(self.groupBox)
        self.label_db0.setMinimumSize(QtCore.QSize(0, 30))
        self.label_db0.setMaximumSize(QtCore.QSize(16777215, 30))
        self.label_db0.setObjectName("label_db0")
        self.verticalLayout_2.addWidget(self.label_db0)
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.label_db1 = QtWidgets.QLabel(self.groupBox)
        self.label_db1.setMinimumSize(QtCore.QSize(0, 30))
        self.label_db1.setObjectName("label_db1")
        self.horizontalLayout.addWidget(self.label_db1)
        self.label_db2 = QtWidgets.QLabel(self.groupBox)
        self.label_db2.setMinimumSize(QtCore.QSize(0, 30))
        self.label_db2.setObjectName("label_db2")
        self.horizontalLayout.addWidget(self.label_db2)
        self.label_db3 = QtWidgets.QLabel(self.groupBox)
        self.label_db3.setMinimumSize(QtCore.QSize(0, 30))
        self.label_db3.setObjectName("label_db3")
        self.horizontalLayout.addWidget(self.label_db3)
        self.verticalLayout_2.addLayout(self.horizontalLayout)
        self.label_status = QtWidgets.QLabel(self.groupBox)
        self.label_status.setMinimumSize(QtCore.QSize(0, 20))
        self.label_status.setMaximumSize(QtCore.QSize(16777215, 20))
        self.label_status.setObjectName("label_status")
        self.verticalLayout_2.addWidget(self.label_status)
        self.verticalLayout.addWidget(self.groupBox)
        spacerItem = QtWidgets.QSpacerItem(20, 63, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout.addItem(spacerItem)
        self.formLayout = QtWidgets.QFormLayout()
        self.formLayout.setObjectName("formLayout")
        self.comboBox_user = QtWidgets.QComboBox(Dialog)
        self.comboBox_user.setMinimumSize(QtCore.QSize(0, 30))
        self.comboBox_user.setObjectName("comboBox_user")
        self.formLayout.setWidget(0, QtWidgets.QFormLayout.FieldRole, self.comboBox_user)
        self.label_user = QtWidgets.QLabel(Dialog)
        self.label_user.setMinimumSize(QtCore.QSize(0, 30))
        self.label_user.setObjectName("label_user")
        self.formLayout.setWidget(0, QtWidgets.QFormLayout.LabelRole, self.label_user)
        self.verticalLayout.addLayout(self.formLayout)
        self.pushButton = QtWidgets.QPushButton(Dialog)
        self.pushButton.setMinimumSize(QtCore.QSize(0, 50))
        self.pushButton.setMaximumSize(QtCore.QSize(16777215, 50))
        self.pushButton.setCheckable(False)
        self.pushButton.setObjectName("pushButton")
        self.verticalLayout.addWidget(self.pushButton)

        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "БД Диспетчеров (Кронштадтское РЖА)"))
        self.groupBox.setTitle(_translate("Dialog", "Проверка данных:"))
        self.label_db0.setText(_translate("Dialog", "Файл базы данных - ..."))
        self.label_db1.setText(_translate("Dialog", "Архив базы 1 - ..."))
        self.label_db2.setText(_translate("Dialog", "Архив базы 2 - ..."))
        self.label_db3.setText(_translate("Dialog", "Архив базы 3 - ..."))
        self.label_status.setText(_translate("Dialog", "..."))
        self.label_user.setText(_translate("Dialog", "Пользователь:"))
        self.pushButton.setText(_translate("Dialog", "Запуск БД"))

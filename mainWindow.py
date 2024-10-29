import sys

from PyQt6 import uic  # Импортируем uic
from PyQt6.QtCore import QStringListModel
from PyQt6.QtWidgets import QApplication, QMainWindow, QFileDialog, QListView
from PyQt6.uic.properties import QtWidgets
from PyQt6.uic.uiparser import QtCore


class MyWidget(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi('mainWindow.ui', self)  # Загружаем дизайн
        self.SelectFile.clicked.connect(self.open_file)

        self.item_list = ['1','2','3']
        self.model_1 = QStringListModel(self)
        self.model_1.setStringList(self.item_list)
        self.listView.setModel(self.model_1)
        # Обратите внимание: имя элемента такое же как в QTDesigner

    def open_file(self):
        self.file_name = QFileDialog.getOpenFileName(None, "Open", "")
        if self.file_name[0] != '':
            self.filePath.setText(self.file_name[0])

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyWidget()
    ex.show()
    sys.exit(app.exec())
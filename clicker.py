import sys
from time import sleep

import ctypes
PROCESS_PER_MONITOR_DPI_AWARE = 2
ctypes.windll.shcore.SetProcessDpiAwareness(PROCESS_PER_MONITOR_DPI_AWARE)
import pandas as pd

from pynput import mouse, keyboard

from PyQt5.QtWidgets import (QApplication, QWidget, QPushButton,
                             QDesktopWidget, QGridLayout, QLabel,
                             QSpinBox, QFileDialog,
                             QTableWidget, QTableWidgetItem)
from PyQt5.QtGui import QColor
from PyQt5.QtCore import Qt

class App(QWidget):
    def __init__(self, df=None):
        super().__init__()
        self.title = 'Clicker'
        self.width = 800
        self.height = 600

        if df is None:
            self.keyEvents = pd.DataFrame(columns=['Type', 'Button', 'Coordinates', 'WaitTime'])
        else:
            self.keyEvents = df

        self.runTimes = 1
        self.skipFirst = 0

        self.mouseWait = 300
        self.keyWait = 200

        self.mListener = mouse.Listener(on_click=self.on_click)
        self.kListener = keyboard.Listener(on_press=self.on_press, on_release=self.on_release)

        self.initUI()

    def initUI(self):
        grid = QGridLayout()
        grid.setRowStretch(12,1)
        grid.setColumnStretch(4,1)
        self.setLayout(grid)

        self.recordButton = QPushButton('Start recording')
        grid.addWidget(self.recordButton, 0, 0)
        self.recordButton.clicked.connect(self.start_record)
        self.recordLabel = QLabel('')
        grid.addWidget(self.recordLabel, 0, 1)

        self.stopButton = QPushButton('Stop/pause\nrecording')
        grid.addWidget(self.stopButton, 1, 0)
        self.stopButton.clicked.connect(self.stop_record)

        self.playButton = QPushButton('Run')
        grid.addWidget(self.playButton, 2, 0)
        self.playButton.clicked.connect(self.play)

        self.playBox = QSpinBox()
        grid.addWidget(self.playBox, 3, 1)
        self.playBox.setMinimum(1)
        self.playBox.setMaximum(1313)
        self.playBox.setValue(self.runTimes)
        self.playBox.valueChanged.connect(self.runTimes_update)
        grid.addWidget(QLabel('Run the commands .. times'), 3, 0)

        grid.addWidget(QLabel('Do not include the first ..\n'
                              'commands when running\n'
                              'multiple times'), 4, 0)
        self.skipBox = QSpinBox()
        grid.addWidget(self.skipBox, 4, 1)
        self.skipBox.setMinimum(0)
        self.skipBox.setMaximum(0)
        self.skipBox.setValue(self.skipFirst)
        self.skipBox.valueChanged.connect(self.skipFirst_update)

        grid.addWidget(QLabel('Default wait-time for\nmouseclicks'), 5, 0)
        grid.addWidget(QLabel('ms'), 5, 2)
        self.mouseBox = QSpinBox()
        grid.addWidget(self.mouseBox, 5, 1)
        self.mouseBox.setMinimum(0)
        self.mouseBox.setMaximum(100000)
        self.mouseBox.setSingleStep(50)
        self.mouseBox.setValue(self.mouseWait)
        self.mouseBox.valueChanged.connect(self.mouseWait_update)

        grid.addWidget(QLabel('Default wait-time for\nkeyboard inputs'), 6, 0)
        grid.addWidget(QLabel('ms'), 6, 2)
        self.keyBox = QSpinBox()
        grid.addWidget(self.keyBox, 6, 1)
        self.keyBox.setMinimum(0)
        self.keyBox.setMaximum(100000)
        self.keyBox.setSingleStep(50)
        self.keyBox.setValue(self.keyWait)
        self.keyBox.valueChanged.connect(self.keyWait_update)


        self.emptyButton = QPushButton('Delete all data')
        grid.addWidget(self.emptyButton, 8, 0)
        self.emptyButton.clicked.connect(self.empty_events)

        self.emptyButton2 = QPushButton('Delete row:')
        self.emptyButton2.setToolTip('Deletes this row number when pressed')
        grid.addWidget(self.emptyButton2, 9, 0)
        self.emptyButton2.clicked.connect(self.del_row)
        self.delBox = QSpinBox()
        grid.addWidget(self.delBox, 9, 1)
        self.delBox.setMinimum(1)

        self.saveButton = QPushButton('Save')
        grid.addWidget(self.saveButton, 11, 0)
        self.saveButton.clicked.connect(self.file_save)

        self.loadButton = QPushButton('Load')
        grid.addWidget(self.loadButton, 12, 0)
        self.loadButton.clicked.connect(self.file_load)

        self.table = QTableWidget(1, len(self.keyEvents.columns))
        self.table.setHorizontalHeaderLabels(self.keyEvents.columns)
        self.table.itemSelectionChanged.connect(self.change_table)
        grid.addWidget(self.table, 0, 4, 12, 1)
        self.update_table()
        grid.addWidget(QLabel('Select another cell after changing a wait-time '
                              'otherwise it will not be registered!'), 12, 4)

        self.setWindowTitle(self.title)
        self.resize(self.width, self.height)
        self.center()
        self.show()


    def keyWait_update(self):
        oldV = self.keyWait
        self.keyWait = self.keyBox.value()
        for i, row in self.keyEvents.iterrows():
            if (not type(row.Coordinates) is tuple and 
                row.WaitTime == oldV/1000 and
                row.Type == 'Press'):
                self.keyEvents.loc[i, 'WaitTime'] = self.keyWait/1000
        self.update_table()


    def mouseWait_update(self):
        oldV = self.mouseWait
        self.mouseWait = self.mouseBox.value()
        for i, row in self.keyEvents.iterrows():
            if (type(row.Coordinates) is tuple and 
                row.WaitTime == oldV/1000 and
                row.Type == 'Press'):
                self.keyEvents.loc[i, 'WaitTime'] = self.mouseWait/1000
        self.update_table()


    def file_save(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getSaveFileName(self,"QFileDialog.getSaveFileName()","","All Files (*);;CSV Files (*.csv)", options=options)
        if fileName:
            self.keyEvents.to_csv(fileName if fileName.endswith('.csv') else
                                  fileName+'.csv', index=False)


    def file_load(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        files, _ = QFileDialog.getOpenFileNames(self,"QFileDialog.getOpenFileNames()", "","All Files (*);;CSV Files (*.csv)", options=options)
        if len(files) == 1 and files[0].endswith('.csv'):
            self.keyEvents = pd.read_csv(files[0])
            Button = mouse.Button
            Key = keyboard.Key
            for i, row in self.keyEvents.iterrows():
                if type(row.Coordinates) is str:
                    row.Coordinates = tuple(eval(row.Coordinates))
                row.Button = eval(row.Button)
                self.keyEvents.iloc[i] = row
            self.update_table()


    def change_table(self):
        for i, row in self.keyEvents.iterrows():
            self.keyEvents.loc[i, 'WaitTime'] = float(self.table.item(i, 3).text())


    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())


    def start_record(self):
        self.recordLabel.setText('<font color="red"><b>Rec.</b></font>')
        self.mListener.start()
        self.kListener.start()


    def stop_record(self):
        if self.mListener.running or self.kListener.running:
            self.recordLabel.setText('')
            self.mListener.stop()
            self.kListener.stop()

            self.mListener = mouse.Listener(on_click=self.on_click)
            self.kListener = keyboard.Listener(on_press=self.on_press, on_release=self.on_release)

            self.keyEvents = self.keyEvents.iloc[:-2]
        self.update_table()


    def runTimes_update(self):
        self.runTimes = self.playBox.value()


    def skipFirst_update(self):
        self.skipFirst = self.skipBox.value()
        self.update_table()


    def play(self):
        if len(self.keyEvents) == 0:
            print('There are no logged clicks/keypresses!')
            self.runLabel.setText('')
            return

        if self.mListener.running or self.kListener.running:
            self.stop_record()

        kController = keyboard.Controller()
        mController = mouse.Controller()

        for run in range(self.runTimes): 
            rows = self.keyEvents[self.skipFirst:]
            if run == 0:
                rows = self.keyEvents
            for i, row in rows.iterrows():
                sleep(row.WaitTime)
                if type(row.Coordinates) is tuple:
                    mController.position = row.Coordinates
                    if row.Type == 'Press':
                        mController.press(row.Button)
                    elif row.Type == 'Release':
                        mController.release(row.Button)
                else:
                    if row.Type == 'Press':
                        kController.press(row.Button)
                    elif row.Type == 'Release':
                        kController.release(row.Button)


    def empty_events(self):
        if self.mListener.running or self.kListener.running:
            self.stop_record()
        self.keyEvents = self.keyEvents.iloc[0:0]
        self.update_table()


    def del_row(self):
        self.keyEvents = self.keyEvents.drop(self.delBox.value()-1)
        self.keyEvents = self.keyEvents.reset_index(drop=True)
        self.update_table()


    def on_click(self, x, y, button, pressed):
        self.keyEvents = self.keyEvents.append(
                {'Type': 'Press' if pressed else 'Release',
                 'Coordinates': (x, y),
                 'Button': button,
                 'WaitTime': self.mouseWait/1000 if pressed else 0
                 }, ignore_index=True)

    
    def on_press(self, key):
        self.keyEvents = self.keyEvents.append(
                {'Type': 'Press',
                 'Button': key,
                 'WaitTime': self.keyWait/1000
                 }, ignore_index=True)


    def on_release(self, key):
        self.keyEvents = self.keyEvents.append(
                {'Type': 'Release',
                 'Button': key,
                 'WaitTime': 0
                 }, ignore_index=True)


    def update_table(self):
        self.skipBox.setMaximum(max(0, len(self.keyEvents)-1))
        self.delBox.setMaximum(max(1, len(self.keyEvents)))

        self.table.setRowCount(len(self.keyEvents))
        for i, row in self.keyEvents.iterrows():
            for j, data in enumerate(row):
                item = QTableWidgetItem(str(data))
                if j != 3:
                    item.setFlags(Qt.ItemIsEnabled)

                color = QColor(255,255,255)
                if i < self.skipFirst:
                    color = QColor(255,0,0)

                item.setBackground(color)
                self.table.setItem(i, j, item)


    def closeEvent(self, event):
        self.stop_record()
        event.accept()



if __name__ == '__main__':
    if not QApplication.instance():
        app = QApplication(sys.argv)
    else:
        app = QApplication.instance()

    ex = App()
    app.exec_()

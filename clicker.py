import sys
from time import sleep

import pandas as pd

from pynput import mouse
from pynput import keyboard

from PyQt5.QtWidgets import (QApplication, QWidget, QPushButton,
                             QDesktopWidget, QGridLayout, QLabel,
                             QSpinBox, QFileDialog,
                             QTableWidget, QTableWidgetItem)
from PyQt5.QtGui import QColor

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

        self.initUI()

    def initUI(self):
        grid = QGridLayout()
        grid.setColumnStretch(4,1)
        grid.setRowStretch(8,1)
        self.setLayout(grid)


        self.recordButton = QPushButton('Start recording')
#        self.recordButton.setToolTip('Starts recording all your mouse and keyboard inputs')
        grid.addWidget(self.recordButton, 0, 0)
        self.recordButton.clicked.connect(self.start_record)
        self.recordLabel = QLabel('')
        grid.addWidget(self.recordLabel, 0, 1)
        
        self.stopButton = QPushButton('Stop/pause\nrecording')
        grid.addWidget(self.stopButton, 1, 0)
        self.stopButton.clicked.connect(self.stop_record)


        self.playButton = QPushButton('Run')
        grid.addWidget(self.playButton, 3, 0)
        self.playButton.clicked.connect(self.play)

        self.playBox = QSpinBox()
        grid.addWidget(self.playBox, 4, 1)
        self.playBox.setMinimum(1)
        self.playBox.setValue(self.runTimes)
        self.playBox.valueChanged.connect(self.runTimes_update)
        grid.addWidget(QLabel('Run the commands .. times'), 4, 0)


        grid.addWidget(QLabel('Do not include the first ..\n'
                              'commands when running\n'
                              'multiple times'), 5, 0)
        self.skipBox = QSpinBox()
        grid.addWidget(self.skipBox, 5, 1)
        self.skipBox.setMinimum(0)
        self.skipBox.setMaximum(0)
        self.skipBox.setValue(self.skipFirst)
        self.skipBox.valueChanged.connect(self.skipFirst_update)
    
        self.emptyButton = QPushButton('Delete all data')
        grid.addWidget(self.emptyButton, 6, 0)
        self.emptyButton.clicked.connect(self.empty_events)
        
        self.emptyButton2 = QPushButton('Delete row:')
        self.emptyButton2.setToolTip('Deletes this row number when pressed')
        grid.addWidget(self.emptyButton2, 7, 0)
        self.emptyButton2.clicked.connect(self.del_row)
        self.delBox = QSpinBox()
        grid.addWidget(self.delBox, 7, 1)
        self.delBox.setMinimum(1)

        self.saveButton = QPushButton('Save')
        grid.addWidget(self.saveButton, 9, 0)
        self.saveButton.clicked.connect(self.file_save)
        
        self.loadButton = QPushButton('Load')
        grid.addWidget(self.loadButton, 10, 0)
        self.loadButton.clicked.connect(self.file_load)


        self.table = QTableWidget(1, len(self.keyEvents.columns))
        self.table.setHorizontalHeaderLabels(self.keyEvents.columns)
        self.table.itemSelectionChanged.connect(self.change_table)
        grid.addWidget(self.table, 0, 4, 10, 1)
        self.update_table()
        grid.addWidget(QLabel('Select another cell after changing a wait-time '
                              'otherwise it will not be registered!'), 10, 4)

#        self.coordLabel = QLabel('Mouse coordinates:\n{}'.format(0))
#        grid.addWidget(self.coordLabel, 6, 0)

        self.setWindowTitle(self.title)
        self.resize(self.width, self.height)
        self.center()
        self.show()

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
        self.mListener = mouse.Listener(on_click=self.on_click)
        self.kListener = keyboard.Listener(on_press=self.on_press, on_release=self.on_release)
        
        self.mListener.start()
        self.kListener.start()
    
    def stop_record(self):
        if self.mListener.running or self.kListener.running:
            self.recordLabel.setText('')
            self.mListener.stop()
            self.kListener.stop()
            
            self.keyEvents = self.keyEvents.iloc[:-2]
#            print('\nStopped:')
#            print(self.keyEvents)
        self.update_table()


    def runTimes_update(self):
        self.runTimes = self.playBox.value()
#        print('repeating', self.playBox.value())

    def skipFirst_update(self):
        self.skipFirst = self.skipBox.value()
        self.skipBox.setMaximum(len(self.keyEvents)-1)
        for i in range(len(self.keyEvents)):
            for j in range(len(self.keyEvents.columns)):
                if i < self.skipFirst:
                    self.table.item(i, j).setBackground(QColor(255,0,0))
                else:
                    self.table.item(i, j).setBackground(QColor(255,255,255))

        print(self.skipBox.value())

    def play(self):
        if len(self.keyEvents) == 0:
            print('There are no logged clicks/keypresses!')
            return
        if self.mListener.running or self.kListener.running:
            self.stop_record()
        self.update_table()
#        print('\nRepeating you {} times!!!'.format(self.runTimes))
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
        self.skipFirst_update()
    
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
#        global keyEvents
        self.keyEvents = self.keyEvents.append(
                {'Type': 'Press' if pressed else 'Release',
                 'Coordinates': (x, y),
                 'Button': button,
                 'WaitTime': 0.2 if pressed else 0
                 }, ignore_index=True)
    
    def on_press(self, key):
#        global keyEvents
        self.keyEvents = self.keyEvents.append(
                {'Type': 'Press',
                 'Button': key,
                 'WaitTime': 0.1
                 }, ignore_index=True)
    
    def on_release(self, key):
#        global keyEvents
        self.keyEvents = self.keyEvents.append(
                {'Type': 'Release',
                 'Button': key,
                 'WaitTime': 0
                 }, ignore_index=True)


    def update_table(self):
        self.skipBox.setMaximum(len(self.keyEvents))
        self.delBox.setMaximum(max(1, len(self.keyEvents)))

        self.table.setRowCount(len(self.keyEvents))
        for i, row in self.keyEvents.iterrows():
            for j, data in enumerate(row):
                item = QTableWidgetItem(str(data))
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
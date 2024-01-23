from PyQt6.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QComboBox
import pandas as pd
from PyQt6.QtCore import Qt
from PyQt6 import QtWidgets
import json
import sys
import serial.tools.list_ports
import serial


def f_find_ports():
    ports = serial.tools.list_ports.comports()
    for port in ports:
        print(port.device)
    return ports


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("My App")
        layout = QVBoxLayout()

        self.button_to_exel = QPushButton("json to excel!")
        self.button_to_exel.clicked.connect(self.f_json_to_excel)

        self.line_ports = QComboBox()
        usb = f_find_ports()
        print(usb)
        self.line_ports.addItems([port.device for port in usb])

        self.button_download = QPushButton("Загрузить")
        self.button_download.setCheckable(False)
        self.button_download.clicked.connect(self.f_save_file)


        layout.addWidget(self.button_to_exel)
        layout.addWidget(self.button_download)
        layout.addWidget(self.line_ports)

        widget = QWidget()
        widget.setLayout(layout)
        self.setCentralWidget(widget)

    def f_json_to_excel(self):
        with open(QtWidgets.QFileDialog.getOpenFileName(
                self,
                'Open File', './',
                'Files (*.json)')[0], 'r') as file:
            print(file)
            data = json.load(file)

        df = pd.json_normalize(data)
        df.to_excel('data.xlsx', index=False)
        print('JSON data has been converted to Excel')

        ## добавить ошибку закрытия файла ##

    def f_save_file(self):
        with open(QtWidgets.QFileDialog.getOpenFileName(
                self,
                'Open File', './',
                'Files (*.bin),(*.txt)')[0], 'rb') as file:
            print(file)
            configr = file.read()
        chosen_port = self.line_ports.currentText()
        print(chosen_port)
            ## добавить ошибку закрытия файла ##
        # скорость в бодах
        speed = 4096
        try:
        # Open the COM port
            ser = serial.Serial(chosen_port, baudrate=speed)
            print("Serial connection established.")

        # Read data from the Arduino
            while True:
                print('start')

                # Send the command to the Arduino
                ser.write(configr)
                transmitted_bits = 0
                # Увеличиваем количество переданных битов
                transmitted_bits += len(configr) * 8

                # Выводим количество переданных битов
                print(f"Transmitted bits: {transmitted_bits}")

            if line:
                print("Received:", line)

        except serial.SerialException as se:
            print("Serial port error:", str(se))

        except KeyboardInterrupt:
            pass

        finally:
        # Close the serial connection
            if ser.is_open:
                ser.close()
                print("Serial connection closed.")

    #
    # def f_send_file_to(self, port, configr):
    #     try:
    #         self.realport = serial.Serial(self.list_ports.currentText(), int(self.Speed.currentText()))
    #         self.ConnectButton.setStyleSheet("background-color: green")
    #         self.ConnectButton.setText('Подключено')
    #     except Exception as e:
    #         print(e)
    #
    # def send(self):
    #     if self.realport:
    #         self.realport.write(b'b')


app = QApplication(sys.argv)

window = MainWindow()
window.show()

app.exec()

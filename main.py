import sys
import cv2
from PyQt5 import uic
from PyQt5.QtWidgets import QApplication
from excelFunctions import *
from Image2str import img2str

def ConvertImage(image):
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
    return thresh

def Start():
    capture_one = cv2.VideoCapture(0)
    capture_two = cv2.VideoCapture(1)
    while True:

        ret1, img1 = capture_one.read()
        cv2.imshow("Camera-1", img1)
        first_result = img2str(ConvertImage(img1))
        if first_result != None:
            name = CheckPerson("Personal.xlsx", first_result)
            if name == None:
                print('Такого номера нет в базе!')
                form.Status.setText("Доступ запрещён")
            else:
                MakeFirstEntry("StatusTable.xlsx", first_result, CheckPerson("Personal.xlsx", first_result))
                print('entry completed')
                form.PersonalName.setText("Сотрудник " + name)
                form.Status.setText("Вошел на территорию")
                time.sleep(5)

        ret2, img2 = capture_two.read()
        cv2.imshow("Camera-2", img2)
        second_result = img2str(ConvertImage(img2))
        if second_result != None:
            name = CheckPerson("Personal.xlsx", second_result)
            if name == None:
                print('Такого номера нет в базе!')
            else:
                MakeSecondEntry("StatusTable.xlsx", second_result)
                print('second entry completed')
                form.PersonalName.setText("Сотрудник " + name)
                form.Status.setText("Покинул территорию")
                time.sleep(3)


        k = cv2.waitKey(30) & 0xFF
        if k == 27:
            break
    capture_one.release()
    capture_two.release()
    cv2.destroyAllWindows()

# вызов GUI
Form, Window = uic.loadUiType('pyqt.ui')
app = QApplication([])
window = Window()
form = Form()
form.setupUi(window)
window.show()

#Кнопки
form.Start.clicked.connect(Start)
form.Excel_1.clicked.connect(OpenPersonalDataBase)
form.Excel_2.clicked.connect(OpenMonitoringWindow)
form.Exit.clicked.connect(sys.exit)

app.exec()

from PyQt5.uic import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import sys
import Bentley_algorithm as Ba
import numpy as np
import matplotlib.pyplot as plt
import xlsxwriter


class MainWindow(QMainWindow):
    
    

    def __init__(self, parent=None):
        super().__init__(parent)
        loadUi('Interface.ui',self)
        self.setWindowTitle("Bentley")
        self.ConnectSignals()
        self.coord_tuples = []
        self.scale_index = 2
        self.draw()
    
    
    @pyqtSlot()
    def AddPushButtonPressed(self):
        Ax = self.doubleSpinBox_Ax.value()
        Ay = self.doubleSpinBox_Ay.value()
        Bx = self.doubleSpinBox_Bx.value()
        By = self.doubleSpinBox_By.value()
        length = np.sqrt(pow(Bx - Ax,2) + pow(By - Ay,2))
        self.coord_tuples.append(((Ax, Ay),(Bx, By)))
        text = "A = [" + str(Ax) + ", " + str(Ay) + "] B = [" +  str(Bx) + ", " + str(By) + "] length = " + str(length)
        self.Coords_listWidget.addItem(text)
        self.draw()
    
    
    @pyqtSlot()
    def FindPushButtonPressed(self):
        isect = Ba.isect_segments(self.coord_tuples)
        
        absciss_coords = [[-100, 100], [0, 0]]
        ordinat_coords = [[0, 0], [-100, 100]]
        plt.plot(absciss_coords[0], absciss_coords[1], color = 'gray', linewidth = 1)
        plt.plot(ordinat_coords[0], ordinat_coords[1], color = 'grey',linewidth = 1)
        
        for lines_index in range(len(self.coord_tuples)):
            plt.plot([self.coord_tuples[lines_index][0][0],self.coord_tuples[lines_index][1][0]],
                     [self.coord_tuples[lines_index][0][1],self.coord_tuples[lines_index][1][1]],
                     color = 'black')
        
        workbook = xlsxwriter.Workbook('Output.xlsx')
        worksheet = workbook.add_worksheet()
        worksheet.write('A1', 'X')
        worksheet.write('B1', 'Y')
        for circle_index in range(len(isect)):
            circle = plt.Circle(isect[circle_index], 0.3, color = 'r')
            plt.gca().add_artist(circle)
            plt.text(isect[circle_index][0]-1.5,isect[circle_index][1]+0.5,str(isect[circle_index]))
            worksheet.write('A' + str(circle_index + 2), str(isect[circle_index][0]))
            worksheet.write('B' + str(circle_index + 2), str(isect[circle_index][1]))
        workbook.close()
        plt.show()   
   
        
    @pyqtSlot()
    def ZoomInPushButtonPressed(self):
        self.scale_index *= 1.25
        self.draw()
        
   
    @pyqtSlot()
    def ZoomOutPushButtonPressed(self):
        self.scale_index /= 1.25
        self.draw()
     
       
    @pyqtSlot()
    def DeleteAllPushButtonPressed(self):
        msgBox = QMessageBox()
        msgBox.setIcon(QMessageBox.Question)
        msgBox.setText("Delete all lines?")
        msgBox.setWindowTitle("Confirmation")
        msgBox.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        returnValue = msgBox.exec()
        if returnValue == QMessageBox.Yes:
            self.doubleSpinBox_Ax.clear()
            self.doubleSpinBox_Ay.clear()
            self.doubleSpinBox_Bx.clear()
            self.doubleSpinBox_By.clear()
            self.Coords_listWidget.clear()
            self.coord_tuples = []
            self.scale_index = 2
            self.draw()
    
        
    @pyqtSlot()
    def DeletePushButtonPressed(self):
        check_row_index = self.Coords_listWidget.selectedItems()
        if  not check_row_index:
            QMessageBox.about(None, "Error", "Mark any row")
        else:
            list_row_index = self.Coords_listWidget.currentRow()
            self.Coords_listWidget.takeItem(list_row_index)
            del self.coord_tuples[list_row_index]
            self.draw()
     
                    
    @pyqtSlot()
    def UsageActionTriggered(self):
        QMessageBox.about(None, "Применение", "Программу можно использовать для нахождения точек пересечения заданных прямых")

    @pyqtSlot()
    def HowToEnterDataActionTriggered(self):
        QMessageBox.about(None, "Использование", "Для ввода данных предоставлены 4 спинбокса.\n \
            Числа можно вводить при помощи кнопочек вверх и вних \
            Заданные ограничения ввода от -100 до 100, шаг прибавления 0.01")

    @pyqtSlot()
    def ElementsActionTriggered(self):
        QMessageBox.about(None, "Элементы", "Для добавления прямой на график предусмотрена кнопка \"Add\".\
        Для приближения элементов на графике используйте кнопки \"+\" и \"-\".\n\
        \"Delete all\" позволяют удалить все построенные прямые из списка и из графика.\n\
        \"Dlelete\" позволяет удалить выделенную построенную прямую из списка и из графика\n\
        \"Find dots\" находит все точки пересечения прямых, строит график с полученными точками и добавляет координаты в Excel файл")
   
   
    def draw(self):
        scene = QGraphicsScene()
        self.graphicsView.setScene(scene)
        pen = QPen()
        pen.setColor(Qt.black)
        pen.setWidth(1)
        axess_pen = QPen()
        axess_pen.setColor(Qt.black)
        axess_pen.setWidth(0.5)
        dot_pen = QPen()
        dot_pen.setColor(Qt.yellow)
        absciss = QLineF(-100,0,100,0)
        ordinat = QLineF(0,-100,0,100)
        item_abs = scene.addLine(absciss,axess_pen)
        item_ord = scene.addLine(ordinat,axess_pen)
        for line_index in range(len(self.coord_tuples)): 
            item = scene.addLine(self.coord_tuples[line_index][0][0],
                                 self.coord_tuples[line_index][0][1] *-1,
                                 self.coord_tuples[line_index][1][0], 
                                 self.coord_tuples[line_index][1][1] *-1,
                                 pen)
            item_dotA = scene.addEllipse(self.coord_tuples[line_index][0][0] - 0.375,
                                        (-1*self.coord_tuples[line_index][0][1]) - 0.375,
                                        0.75,0.75,dot_pen)
            item_dotB = scene.addEllipse(self.coord_tuples[line_index][1][0] - 0.375,
                                        (-1*self.coord_tuples[line_index][1][1]) - 0.375,
                                        0.75,0.75,dot_pen)
            item_dotA.setScale(self.scale_index)
            item_dotB.setScale(self.scale_index)
            item.setScale(self.scale_index)
        item_abs.setScale(self.scale_index)
        item_ord.setScale(self.scale_index)
        
           
    def ConnectSignals(self):
       self.usage_action.triggered.connect(self.UsageActionTriggered)
       self.how_to_enter_data_action.triggered.connect(self.HowToEnterDataActionTriggered)
       self.elements_action.triggered.connect(self.ElementsActionTriggered)
       self.add_pushButton.pressed.connect(self.AddPushButtonPressed)
       self.find_pushButton.pressed.connect(self.FindPushButtonPressed)
       self.zoom_in_pushButton.pressed.connect(self.ZoomInPushButtonPressed)
       self.zoom_out_pushButton.pressed.connect(self.ZoomOutPushButtonPressed)
       self.delete_all_pushButton.pressed.connect(self.DeleteAllPushButtonPressed)
       self.delete_pushButton.pressed.connect(self.DeletePushButtonPressed)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())    
    

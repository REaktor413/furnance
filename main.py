# -*- coding: utf8 -*- 
import os, sys, math
from datetime import *
from datetime import timedelta
from PyQt4 import QtGui 
from PyQt4 import QtCore 
from PyQt4.QtCore import SIGNAL,SLOT
import win32com.client


class MainWindow(QtGui.QMainWindow):
	def __init__(self):
		QtGui.QMainWindow.__init__(self)

		self.resize(1400, 800)
		self.setWindowTitle(("Расчет работы туннельной печи №2, цеха №3 ООО 'ДОЗ' "))
		self.setWindowIcon(QtGui.QIcon("logo.png"))
		
		#Разметка рабочей области интерфейса 
		group1 = QtGui.QGroupBox(u"Исходные данные:",self ) 
		group1.resize(1100,200)
		group1.move(20,10)

		group2 = QtGui.QGroupBox(u"Сегодня",self ) 
		group2.resize(240,200)
		group2.move(1140,10)

		group3 = QtGui.QGroupBox(u"Исходные данные о планируемом составе (количество, время прогонок, разрывы):",self ) 
		group3.resize(1360,450)
		group3.move(20,230)

		#Прорисовка надписей и списков выбора группы №1
		lbl_fist_vagon = QtGui.QLabel("Дата и время загонки первого вагона из планируемого состава: ", group1)
		lbl_fist_vagon.move(20,30)

		lbl_fist_day = QtGui.QLabel("День:", group1)
		lbl_fist_day.move(70,60)
		ent_fist_day = QtGui.QLineEdit(group1) 
		ent_fist_day.setMaxLength(2)
		ent_fist_day.setAlignment(QtCore.Qt.AlignCenter)
		ent_fist_day.setPlaceholderText("Число")
		ent_fist_day.move(20,80)

		lbl_fist_mounth = QtGui.QLabel("Месяц:", group1)
		lbl_fist_mounth.move(270,60)
		ent_fist_mounth = QtGui.QLineEdit(group1) 
		ent_fist_mounth.setMaxLength(2)
		ent_fist_mounth.setAlignment(QtCore.Qt.AlignCenter)
		ent_fist_mounth.setPlaceholderText("Месяц")
		ent_fist_mounth.move(220,80)
		
		lbl_fist_year = QtGui.QLabel("Год:", group1)
		lbl_fist_year.move(470,60)
		ent_fist_year = QtGui.QLineEdit(group1) 
		ent_fist_year.setMaxLength(4)
		ent_fist_year.setAlignment(QtCore.Qt.AlignCenter)
		ent_fist_year.setPlaceholderText("Год")
		ent_fist_year.move(420,80)

		lbl_fist_hour = QtGui.QLabel("Часов:", group1)
		lbl_fist_hour.move(670,60)
		ent_fist_hour = QtGui.QLineEdit(group1) 
		ent_fist_hour.setMaxLength(2)
		ent_fist_hour.setAlignment(QtCore.Qt.AlignCenter)
		ent_fist_hour.setPlaceholderText("Час первого вагона")
		ent_fist_hour.move(620,80)

		lbl_fist_minutes = QtGui.QLabel("Минут:", group1)
		lbl_fist_minutes.move(870,60)
		ent_fist_minutes = QtGui.QLineEdit(group1) 
		ent_fist_minutes.setMaxLength(2)
		ent_fist_minutes.setAlignment(QtCore.Qt.AlignCenter)
		ent_fist_minutes.setPlaceholderText("Минуты первого вагона")
		ent_fist_minutes.move(820,80)

		lbl_fist_28 = QtGui.QLabel("Движение состава до позиции №                                                                 по:", group1)
		lbl_fist_28.move(20,140)
		lbl1_fist_28 = QtGui.QLabel("минут", group1)
		lbl1_fist_28.move(570,140)
		
		ent_fist_28 = QtGui.QLineEdit(group1) 
		ent_fist_28.setMaxLength(2)
		ent_fist_28.setAlignment(QtCore.Qt.AlignCenter)
		ent_fist_28.setPlaceholderText("№ позиции")
		ent_fist_28.move(220,138)

		ent_fist_28_min = QtGui.QLineEdit(group1) 
		ent_fist_28_min.setMaxLength(3)
		ent_fist_28_min.setAlignment(QtCore.Qt.AlignCenter)
		ent_fist_28_min.setPlaceholderText("минут")
		ent_fist_28_min.move(420,138)

		#Прорисовка надписей и списков выбора группы №2
		timenow = datetime.now()
		timenow = timenow.strftime("%d - %m - %y")
		lbl_time = QtGui.QLabel(timenow, group2)
		lbl_time.setFont(QtGui.QFont("Serif", 24, QtGui.QFont.Bold))
		lbl_time.move(30,80)

		#Прорисовка надписей и списков выбора группы №3
		lbl_number = QtGui.QLabel("№ п/п", group3)
		lbl_number.move(20,40)
		lbl_number1 = QtGui.QLabel("1", group3)
		lbl_number1.move(30,80)
		lbl_number2 = QtGui.QLabel("2", group3)
		lbl_number2.move(30,120)
		lbl_number3 = QtGui.QLabel("3", group3)
		lbl_number3.move(30,160)
		lbl_number4 = QtGui.QLabel("4", group3)
		lbl_number4.move(30,200)
		lbl_number5 = QtGui.QLabel("5", group3)
		lbl_number5.move(30,240)
		lbl_number6 = QtGui.QLabel("6", group3)
		lbl_number6.move(30,280)
		lbl_number7 = QtGui.QLabel("7", group3)
		lbl_number7.move(30,320)
		
		lbl_brand = QtGui.QLabel("Марка", group3)
		lbl_brand.move(120,40)
		lbl_tonnage = QtGui.QLabel("Тоннаж, т", group3)
		lbl_tonnage.move(280,40)
		lbl_ton_vagon = QtGui.QLabel("На вагоне, т", group3)
		lbl_ton_vagon.move(480,40)
		lbl_period = QtGui.QLabel("Прогонки, мин", group3)
		lbl_period.move(680,40)
		lbl_temp = QtGui.QLabel("Температура, 0С", group3)
		lbl_temp.move(880,40)
		lbl_gas = QtGui.QLabel("Расход газа, М3/час", group3)
		lbl_gas.move(1080,40)
		
		#Прорисовка полей ввода по марке №1
		ent_brand1 = QtGui.QLineEdit(group3) 
		ent_brand1.setAlignment(QtCore.Qt.AlignCenter)
		ent_brand1.setPlaceholderText("Введите первую марку")
		ent_brand1.move(80,78)

		ent_tonnage1 = QtGui.QLineEdit(group3) 
		ent_tonnage1.setAlignment(QtCore.Qt.AlignCenter)
		ent_tonnage1.setPlaceholderText("Тоннаж первой марки")
		ent_tonnage1.move(250,78)

		ent_ton_vagon1 = QtGui.QLineEdit(group3) 
		ent_ton_vagon1.setAlignment(QtCore.Qt.AlignCenter)
		ent_ton_vagon1.setPlaceholderText("Тонн на вагоне")
		ent_ton_vagon1.move(450,78)

		ent_period1 = QtGui.QLineEdit(group3) 
		ent_period1.setAlignment(QtCore.Qt.AlignCenter)
		ent_period1.setPlaceholderText("Минут")
		ent_period1.move(660,78)

		ent_temp1 = QtGui.QLineEdit(group3) 
		ent_temp1.setAlignment(QtCore.Qt.AlignCenter)
		ent_temp1.setPlaceholderText("Температура")
		ent_temp1.move(860,78)

		ent_gas1 = QtGui.QLineEdit(group3) 
		ent_gas1.setAlignment(QtCore.Qt.AlignCenter)
		ent_gas1.setPlaceholderText("м3 в час")
		ent_gas1.move(1060,78)

		#Прорисовка полей ввода по марке №2
		ent_brand2 = QtGui.QLineEdit(group3) 
		ent_brand2.setAlignment(QtCore.Qt.AlignCenter)
		ent_brand2.setPlaceholderText("Введите вторую марку")
		ent_brand2.move(80,118)

		ent_tonnage2 = QtGui.QLineEdit(group3) 
		ent_tonnage2.setAlignment(QtCore.Qt.AlignCenter)
		ent_tonnage2.setPlaceholderText("Тоннаж второй марки")
		ent_tonnage2.move(250,118)

		ent_ton_vagon2 = QtGui.QLineEdit(group3) 
		ent_ton_vagon2.setAlignment(QtCore.Qt.AlignCenter)
		ent_ton_vagon2.setPlaceholderText("Тонн на вагоне")
		ent_ton_vagon2.move(450,118)

		ent_period2 = QtGui.QLineEdit(group3) 
		ent_period2.setAlignment(QtCore.Qt.AlignCenter)
		ent_period2.setPlaceholderText("Минут")
		ent_period2.move(660,118)

		ent_temp2 = QtGui.QLineEdit(group3) 
		ent_temp2.setAlignment(QtCore.Qt.AlignCenter)
		ent_temp2.setPlaceholderText("Температура")
		ent_temp2.move(860,118)

		ent_gas2 = QtGui.QLineEdit(group3) 
		ent_gas2.setAlignment(QtCore.Qt.AlignCenter)
		ent_gas2.setPlaceholderText("м3 в час")
		ent_gas2.move(1060,118)

		#Прорисовка полей ввода по марке №3
		ent_brand3 = QtGui.QLineEdit(group3) 
		ent_brand3.setAlignment(QtCore.Qt.AlignCenter)
		ent_brand3.setPlaceholderText("Введите третью марку")
		ent_brand3.move(80,158)

		ent_tonnage3 = QtGui.QLineEdit(group3) 
		ent_tonnage3.setAlignment(QtCore.Qt.AlignCenter)
		ent_tonnage3.setPlaceholderText("Тоннаж третьей марки")
		ent_tonnage3.move(250,158)

		ent_ton_vagon3 = QtGui.QLineEdit(group3) 
		ent_ton_vagon3.setAlignment(QtCore.Qt.AlignCenter)
		ent_ton_vagon3.setPlaceholderText("Тонн на вагоне")
		ent_ton_vagon3.move(450,158)

		ent_period3 = QtGui.QLineEdit(group3) 
		ent_period3.setAlignment(QtCore.Qt.AlignCenter)
		ent_period3.setPlaceholderText("Минут")
		ent_period3.move(660,158)

		ent_temp3 = QtGui.QLineEdit(group3) 
		ent_temp3.setAlignment(QtCore.Qt.AlignCenter)
		ent_temp3.setPlaceholderText("Температура")
		ent_temp3.move(860,158)

		ent_gas3 = QtGui.QLineEdit(group3) 
		ent_gas3.setAlignment(QtCore.Qt.AlignCenter)
		ent_gas3.setPlaceholderText("м3 в час")
		ent_gas3.move(1060,158)

		#Прорисовка полей ввода по марке №4
		ent_brand4 = QtGui.QLineEdit(group3) 
		ent_brand4.setAlignment(QtCore.Qt.AlignCenter)
		ent_brand4.setPlaceholderText("Введите 4-ю марку")
		ent_brand4.move(80,198)

		ent_tonnage4 = QtGui.QLineEdit(group3) 
		ent_tonnage4.setAlignment(QtCore.Qt.AlignCenter)
		ent_tonnage4.setPlaceholderText("Тоннаж 4-й марки")
		ent_tonnage4.move(250,198)

		ent_ton_vagon4 = QtGui.QLineEdit(group3) 
		ent_ton_vagon4.setAlignment(QtCore.Qt.AlignCenter)
		ent_ton_vagon4.setPlaceholderText("Тонн на вагоне")
		ent_ton_vagon4.move(450,198)

		ent_period4 = QtGui.QLineEdit(group3) 
		ent_period4.setAlignment(QtCore.Qt.AlignCenter)
		ent_period4.setPlaceholderText("Минут")
		ent_period4.move(660,198)

		ent_temp4 = QtGui.QLineEdit(group3) 
		ent_temp4.setAlignment(QtCore.Qt.AlignCenter)
		ent_temp4.setPlaceholderText("Температура")
		ent_temp4.move(860,198)

		ent_gas4 = QtGui.QLineEdit(group3) 
		ent_gas4.setAlignment(QtCore.Qt.AlignCenter)
		ent_gas4.setPlaceholderText("м3 в час")
		ent_gas4.move(1060,198)
		
		#Прорисовка полей ввода по марке №5
		ent_brand5 = QtGui.QLineEdit(group3) 
		ent_brand5.setAlignment(QtCore.Qt.AlignCenter)
		ent_brand5.setPlaceholderText("Введите 5-ю марку")
		ent_brand5.move(80,238)

		ent_tonnage5 = QtGui.QLineEdit(group3) 
		ent_tonnage5.setAlignment(QtCore.Qt.AlignCenter)
		ent_tonnage5.setPlaceholderText("Тоннаж 5-й марки")
		ent_tonnage5.move(250,238)

		ent_ton_vagon5 = QtGui.QLineEdit(group3) 
		ent_ton_vagon5.setAlignment(QtCore.Qt.AlignCenter)
		ent_ton_vagon5.setPlaceholderText("Тонн на вагоне")
		ent_ton_vagon5.move(450,238)

		ent_period5 = QtGui.QLineEdit(group3) 
		ent_period5.setAlignment(QtCore.Qt.AlignCenter)
		ent_period5.setPlaceholderText("Минут")
		ent_period5.move(660,238)

		ent_temp5 = QtGui.QLineEdit(group3) 
		ent_temp5.setAlignment(QtCore.Qt.AlignCenter)
		ent_temp5.setPlaceholderText("Температура")
		ent_temp5.move(860,238)

		ent_gas5 = QtGui.QLineEdit(group3) 
		ent_gas5.setAlignment(QtCore.Qt.AlignCenter)
		ent_gas5.setPlaceholderText("м3 в час")
		ent_gas5.move(1060,238)
		
		#Прорисовка полей ввода по марке №6
		ent_brand6 = QtGui.QLineEdit(group3) 
		ent_brand6.setAlignment(QtCore.Qt.AlignCenter)
		ent_brand6.setPlaceholderText("Введите 6-ю марку")
		ent_brand6.move(80,278)

		ent_tonnage6 = QtGui.QLineEdit(group3) 
		ent_tonnage6.setAlignment(QtCore.Qt.AlignCenter)
		ent_tonnage6.setPlaceholderText("Тоннаж 6-й марки")
		ent_tonnage6.move(250,278)

		ent_ton_vagon6 = QtGui.QLineEdit(group3) 
		ent_ton_vagon6.setAlignment(QtCore.Qt.AlignCenter)
		ent_ton_vagon6.setPlaceholderText("Тонн на вагоне")
		ent_ton_vagon6.move(450,278)

		ent_period6 = QtGui.QLineEdit(group3) 
		ent_period6.setAlignment(QtCore.Qt.AlignCenter)
		ent_period6.setPlaceholderText("Минут")
		ent_period6.move(660,278)

		ent_temp6 = QtGui.QLineEdit(group3) 
		ent_temp6.setAlignment(QtCore.Qt.AlignCenter)
		ent_temp6.setPlaceholderText("Температура")
		ent_temp6.move(860,278)

		ent_gas6 = QtGui.QLineEdit(group3) 
		ent_gas6.setAlignment(QtCore.Qt.AlignCenter)
		ent_gas6.setPlaceholderText("м3 в час")
		ent_gas6.move(1060,278)
		
		#Прорисовка полей ввода по марке №7		
		ent_brand7 = QtGui.QLineEdit(group3) 
		ent_brand7.setAlignment(QtCore.Qt.AlignCenter)
		ent_brand7.setPlaceholderText("Введите 7-ю марку")
		ent_brand7.move(80,318)

		ent_tonnage7 = QtGui.QLineEdit(group3) 
		ent_tonnage7.setAlignment(QtCore.Qt.AlignCenter)
		ent_tonnage7.setPlaceholderText("Тоннаж 7-й марки")
		ent_tonnage7.move(250,318)

		ent_ton_vagon7 = QtGui.QLineEdit(group3) 
		ent_ton_vagon7.setAlignment(QtCore.Qt.AlignCenter)
		ent_ton_vagon7.setPlaceholderText("Тонн на вагоне")
		ent_ton_vagon7.move(450,318)

		ent_period7 = QtGui.QLineEdit(group3) 
		ent_period7.setAlignment(QtCore.Qt.AlignCenter)
		ent_period7.setPlaceholderText("Минут")
		ent_period7.move(660,318)

		ent_temp7 = QtGui.QLineEdit(group3) 
		ent_temp7.setAlignment(QtCore.Qt.AlignCenter)
		ent_temp7.setPlaceholderText("Температура")
		ent_temp7.move(860,318)

		ent_gas7 = QtGui.QLineEdit(group3) 
		ent_gas7.setAlignment(QtCore.Qt.AlignCenter)
		ent_gas7.setPlaceholderText("м3 в час")
		ent_gas7.move(1060,318)
		
		# Функция ошибки
		def err(err):
			def close_err():
				modal_err.close()

			# Прорисовка окна ошибки
			global modal_err
			modal_err = QtGui.QWidget()
			modal_err.resize(300, 150)
			modal_err.setWindowIcon(QtGui.QIcon("logo.png"))
			modal_err.setWindowTitle("ОШИБКА !!!")
			
			lbl_err = QtGui.QLabel("Не верно заполнено или не заполнено поле: ", modal_err) 
			lbl_err1 = QtGui.QLabel(err, modal_err) 
			lbl_err.move(40,40)
			lbl_err1.move(60,60)

			btn_err =  QtGui.QPushButton("Исправить", modal_err)
			btn_err.move(200,100)
			btn_err.connect(btn_err, QtCore.SIGNAL("clicked()"), close_err)

			modal_err.show()

		# Функция расчета состава с проверкой на ошибки.
		def count_furnance():
			try:
				start_day = int(ent_fist_day.text())
				if start_day <= 0 or start_day > 31:
					err(lbl_fist_day.text())
			except:
				err(lbl_fist_day.text())
			try:
				start_mounth = int(ent_fist_mounth.text())
				if start_mounth <= 0 or start_mounth > 12:
					err(lbl_fist_mounth.text())
			except:
				err(lbl_fist_mounth.text())
			try:
				start_year = int(ent_fist_year.text())
				if start_year < 2019 or start_year > 2050:
					err(lbl_fist_year.text())
			except:
				err(lbl_fist_year.text())
			try:
				start_hour = int(ent_fist_hour.text())
				if start_hour < 0 or start_hour > 24:
					err(lbl_fist_hour.text())
			except:
				err(lbl_fist_hour.text())
			try:
				start_minutes = int(ent_fist_minutes.text())
				if start_minutes < 0 or start_minutes >= 60:
					err(lbl_fist_minutes.text())
			except:
				err(lbl_fist_minutes.text())
			try:
				start_fist_28 = int(ent_fist_28.text())
				if start_fist_28 < 0 or start_fist_28 >= 40:
					err(lbl_fist_28.text())
			except:
				err(lbl_fist_28.text())
			try:
				start_fist_28_min = int(ent_fist_28_min.text())
				if start_fist_28_min < 20 or start_fist_28_min >= 121:
					err("Прогонки до позиции..(Минут)")
			except:
				err("Прогонки до позиции..(Минут)")
			
			# Перевод введенных данных в формат datetime для определения начала первой прогонки и расчета delta
			start_time = datetime(start_year, start_mounth, start_day, start_hour, start_minutes)
			time_list =[]
			brand_list = []
			# Подсчет количества вагонов 1-й марки
			if ent_brand1.text() != '':
				try:
					count_vagon1 = math.floor(float(ent_tonnage1.text())/float(ent_ton_vagon1.text()))
				except:
					err(" Ошибка, для десятых использовать точку")
			
				vagon_list1 = []

				for i in range(count_vagon1):
					vagon_list = [ent_brand1.text(), i,int(ent_period1.text()) ]
					vagon_list1.append(vagon_list)
				brand_list.append(ent_brand1.text())
				brand_list.append(count_vagon1)
				brand_list.append(int(ent_period1.text()))

			# Подсчет количества вагонов 2-й марки
			if ent_brand2.text() != '':
				try:
					count_vagon2 = math.floor(float(ent_tonnage2.text())/float(ent_ton_vagon2.text()))
				except:
					err(" Ошибка, для десятых использовать точку")
			
				vagon_list2 = []

				for i in range(count_vagon2):
					vagon_list = [ent_brand2.text(), i,int(ent_period2.text()) ]
					vagon_list2.append(vagon_list)
				brand_list.append(ent_brand2.text())
				brand_list.append(count_vagon2)
				brand_list.append(int(ent_period2.text()))
					
			
			# Подсчет количества вагонов 3-й марки
			if ent_brand3.text() != '':
				try:
					count_vagon3 = math.floor(float(ent_tonnage3.text())/float(ent_ton_vagon3.text()))
				except:
					err(" Ошибка, для десятых использовать точку")
			
				vagon_list3 = []

				for i in range(count_vagon3):
					vagon_list = [ent_brand3.text(), i,int(ent_period3.text()) ]
					vagon_list3.append(vagon_list)
				brand_list.append(ent_brand3.text())
				brand_list.append(count_vagon3)
				brand_list.append(int(ent_period3.text()))

			# Подсчет количества вагонов 4-й марки
			if ent_brand4.text() != '':
				try:
					count_vagon4 = math.floor(float(ent_tonnage4.text())/float(ent_ton_vagon4.text()))
				except:
					err(" Ошибка, для десятых использовать точку")
			
				vagon_list4 = []

				for i in range(count_vagon4):
					vagon_list = [ent_brand4.text(), i,int(ent_period4.text()) ]
					vagon_list4.append(vagon_list)
				brand_list.append(ent_brand4.text())
				brand_list.append(count_vagon4)
				brand_list.append(int(ent_period4.text()))

			# Подсчет количества вагонов 5-й марки
			if ent_brand5.text() != '':
				try:
					count_vagon5 = math.floor(float(ent_tonnage5.text())/float(ent_ton_vagon5.text()))
				except:
					err(" Ошибка, для десятых использовать точку")
			
				vagon_list5 = []

				for i in range(count_vagon5):
					vagon_list = [ent_brand5.text(), i,int(ent_period5.text()) ]
					vagon_list5.append(vagon_list)
				brand_list.append(ent_brand5.text())
				brand_list.append(count_vagon5)
				brand_list.append(int(ent_period5.text()))

			# Подсчет количества вагонов 6-й марки
			if ent_brand6.text() != '':
				try:
					count_vagon6 = math.floor(float(ent_tonnage6.text())/float(ent_ton_vagon6.text()))
				except:
					err(" Ошибка, для десятых использовать точку")
			
				vagon_list6 = []

				for i in range(count_vagon6):
					vagon_list = [ent_brand6.text(), i,int(ent_period6.text()) ]
					vagon_list6.append(vagon_list)
				brand_list.append(ent_brand6.text())
				brand_list.append(count_vagon6)
				brand_list.append(int(ent_period6.text()))

			# Подсчет количества вагонов 7-й марки
			if ent_brand7.text() != '':
				try:
					count_vagon7 = math.floor(float(ent_tonnage7.text())/float(ent_ton_vagon7.text()))
				except:
					err(" Ошибка, для десятых использовать точку")
			
				vagon_list7 = []

				for i in range(count_vagon7):
					vagon_list = [ent_brand7.text(), i,int(ent_period7.text()) ]
					vagon_list7.append(vagon_list)
				brand_list.append(ent_brand7.text())
				brand_list.append(count_vagon7)
				brand_list.append(int(ent_period7.text()))

			# Конкатениция списков всех марок
			vagon_list_all  =[]
			if ent_brand1.text() != '':
				vagon_list_all += vagon_list1
			if ent_brand2.text() != '':
				vagon_list_all += vagon_list2
			if ent_brand3.text() != '':
				vagon_list_all += vagon_list3
			if ent_brand4.text() != '':
				vagon_list_all += vagon_list4
			if ent_brand5.text() != '':
				vagon_list_all += vagon_list5
			if ent_brand6.text() != '':
				vagon_list_all += vagon_list6
			if ent_brand7.text() != '':
				vagon_list_all += vagon_list7

			# Расчет общего количества вагонов
			count_vagon_all = len(vagon_list_all)

			# Замена времени прогонок для первых позиций
			end_fist_28 = int(ent_fist_28.text())

			if count_vagon_all < end_fist_28:
				for i in range(count_vagon_all):
					vagon_list_all[i][2] = int(ent_fist_28_min.text())
			else:
				for i in range(end_fist_28):
					vagon_list_all[i][2] = int(ent_fist_28_min.text())
			

			# Ядро расчета туннельной печи
			def count_core(start_brand, start_vagon, start_brand_period, end_brand, end_vagon, end_brand_period):
				
				#Два состава белого
				if "М" in start_brand and "М" in end_brand:
					if start_brand_period > 40 and end_brand_period <= 40 or start_brand_period == 40 and  end_brand_period < 40:
						for i in range(start_vagon-1,start_vagon+40,1):
							vagon_list_all[i][2] = start_brand_period

					if start_brand_period <= 40 and end_brand_period > 40 or start_brand_period < 40 and  end_brand_period >= 40:
						for i in range(start_vagon-1,start_vagon+26,1):
							vagon_list_all[i][2] = start_brand_period
					if start_brand_period == end_brand_period:
						pass
					
				# Два состава черного
				if "П" in start_brand and "П" in end_brand:
					if start_brand_period > 40 and end_brand_period <= 40 or start_brand_period == 40 and  end_brand_period < 40:
						for i in range(start_vagon-1,start_vagon+40,1):
							vagon_list_all[i][2] = start_brand_period

					if start_brand_period <= 40 and end_brand_period > 40 or start_brand_period < 40 and  end_brand_period >= 40:
						for i in range(start_vagon-1,start_vagon+29,1):
							vagon_list_all[i][2] = start_brand_period
					if start_brand_period == end_brand_period:
						pass
					
				# Переход с белого на черное
				if "М" in start_brand and "П" in end_brand:
					for i in range(start_vagon,start_vagon+9,1):
							vagon_list_all.insert(i,["разрыв",i,start_brand_period])
					
					flag1 = brand_list.index(start_brand)
					brand_list[flag1+1] = start_vagon+9

					for i in range(start_vagon+9,start_vagon+38,1):
							vagon_list_all[i][2] = start_brand_period
								
				# Переход с черного на белое
				if "П" in start_brand and "М" in end_brand:
					for i in range(start_vagon,start_vagon+19,1):
							vagon_list_all.insert(i,["разрыв",i,start_brand_period])
					
					flag1 = brand_list.index(start_brand)
					brand_list[flag1+1] = start_vagon+19

					for i in range(start_vagon+19,start_vagon+38,1):
							vagon_list_all[i][2] = start_brand_period

				# Подсчет времени для каждого состава
				delta_minutes1 = 0
				for i in range(brand_list[1]):
					delta_minutes1 += vagon_list_all[i][2]
								
				delta_minutes2 = 0
				for i in range(brand_list[1]+brand_list[4]):
					delta_minutes2 += vagon_list_all[i][2]





			def single_core(start_brand, start_vagon, start_brand_period):
				# Первый вагон в печь при просчете одиночного состава
				fist_vagon_in = start_time.strftime("%H:%M %d.%m.%Y")
				
				#Последний вагон в печь при просчете одиночного состава
				delta_minutes = 0
				for i in range(len(vagon_list_all)):
					delta_minutes += vagon_list_all[i][2]
				last_vagon_in = start_time + timedelta(minutes=delta_minutes)
				last_vagon_in = last_vagon_in.strftime("%H:%M %d.%m.%Y")

				# Первый вагон из печи при просчете одиночного состава
				if len(vagon_list_all) > 65:
					delta_minutes = 0
					for i in range(65):
						delta_minutes += vagon_list_all[i][2]
					fist_vagon_out = start_time + timedelta(minutes=delta_minutes)
					fist_vagon_out = fist_vagon_out.strftime("%H:%M %d.%m.%Y")
				else:
					fist_vagon_out = "В печи"
				
				# Послединй вагон из печи всегда будет в печи при одиночном составе
				last_vagon_out = "В печи"

				time_list.append(fist_vagon_in)
				time_list.append(last_vagon_in)
				time_list.append(fist_vagon_out)
				time_list.append(last_vagon_out)


												
			# Функции для каждых марок (1-2.2-3.3-4.4-5 и т.д)
			if len(brand_list) <= 3:
				single_core(*brand_list)
			else:
				if len(brand_list) == 6:
					count_core(brand_list[0],brand_list[1],brand_list[2],brand_list[3],brand_list[4],brand_list[5])
				if len(brand_list) == 9:
					count_core(brand_list[0],brand_list[1],brand_list[2],brand_list[3],brand_list[4],brand_list[5])
				
			print(vagon_list_all)
			print(brand_list)
		

			
			# ************************************************** Расчет состава с датами ***********************************************
			# Интерфейс окна для вывода результата

			global modal_res
			modal_res = QtGui.QWidget()
			modal_res.resize(1300, 800)
			modal_res.setWindowIcon(QtGui.QIcon("logo.png"))
			modal_res.setWindowTitle("Результаты расчета режима работы печи №2 цеха №3")
			
			lbl_res = QtGui.QLabel("Результаты расчета по маркам ", modal_res) 
			lbl_res.move(20,40)

			lbl_num1 = QtGui.QLabel("№1", modal_res) 
			lbl_num1.move(270,40)
			lbl_num2 = QtGui.QLabel("№2", modal_res) 
			lbl_num2.move(420,40)
			lbl_num3 = QtGui.QLabel("№3", modal_res) 
			lbl_num3.move(550,40)
			lbl_num4 = QtGui.QLabel("№4", modal_res) 
			lbl_num4.move(680,40)
			lbl_num5 = QtGui.QLabel("№5", modal_res) 
			lbl_num5.move(810,40)
			lbl_num6 = QtGui.QLabel("№6", modal_res) 
			lbl_num6.move(940,40)
			lbl_num7 = QtGui.QLabel("№7", modal_res) 
			lbl_num7.move(1070,40)

			lbl_brand_res = QtGui.QLabel("Марка:", modal_res) 
			lbl_brand_res.move(20,80)

			lbl_brand_res1 = QtGui.QLabel(ent_brand1.text(), modal_res) 
			lbl_brand_res1.move(270,80)
			lbl_brand_res2 = QtGui.QLabel(ent_brand2.text(), modal_res) 
			lbl_brand_res2.move(370,80)
			lbl_brand_res3 = QtGui.QLabel(ent_brand3.text(), modal_res) 
			lbl_brand_res3.move(470,80)
			lbl_brand_res4 = QtGui.QLabel(ent_brand4.text(), modal_res) 
			lbl_brand_res4.move(570,80)
			lbl_brand_res5 = QtGui.QLabel(ent_brand5.text(), modal_res) 
			lbl_brand_res5.move(670,80)
			lbl_brand_res6 = QtGui.QLabel(ent_brand6.text(), modal_res) 
			lbl_brand_res6.move(770,80)
			lbl_brand_res7 = QtGui.QLabel(ent_brand7.text(), modal_res) 
			lbl_brand_res7.move(870,80)

			lbl_tonnage_res = QtGui.QLabel("Всего тонн, т", modal_res) 
			lbl_tonnage_res.move(20,120)
		
			lbl_tonnage_res1 = QtGui.QLabel(ent_tonnage1.text(), modal_res) 
			lbl_tonnage_res1.move(270,120)
			lbl_tonnage_res2 = QtGui.QLabel(ent_tonnage2.text(), modal_res) 
			lbl_tonnage_res2.move(370,120)
			lbl_tonnage_res3 = QtGui.QLabel(ent_tonnage3.text(), modal_res) 
			lbl_tonnage_res3.move(470,120)
			lbl_tonnage_res4 = QtGui.QLabel(ent_tonnage4.text(), modal_res) 
			lbl_tonnage_res4.move(570,120)
			lbl_tonnage_res5 = QtGui.QLabel(ent_tonnage5.text(), modal_res) 
			lbl_tonnage_res5.move(670,120)
			lbl_tonnage_res6 = QtGui.QLabel(ent_tonnage6.text(), modal_res) 
			lbl_tonnage_res6.move(770,120)
			lbl_tonnage_res7 = QtGui.QLabel(ent_tonnage7.text(), modal_res) 
			lbl_tonnage_res7.move(870,120)

			lbl_vagon_res = QtGui.QLabel("Всего вагонов, шт.", modal_res) 
			lbl_vagon_res.move(20,160)
		
			lbl_vagon_res1 = QtGui.QLabel(str(count_vagon1), modal_res) 
			lbl_vagon_res1.move(270,160) 	
			if ent_brand2.text() != '':
				lbl_vagon_res2 = QtGui.QLabel(str(count_vagon2), modal_res) 
				lbl_vagon_res2.move(370,160)
			if ent_brand3.text() != '':
				lbl_vagon_res3 = QtGui.QLabel(str(count_vagon3), modal_res) 
				lbl_vagon_res3.move(470,160)
			if ent_brand4.text() != '':
				lbl_vagon_res4 = QtGui.QLabel(str(count_vagon4), modal_res) 
				lbl_vagon_res4.move(570,160)
			if ent_brand5.text() != '':
				lbl_vagon_res5 = QtGui.QLabel(str(count_vagon5), modal_res) 
				lbl_vagon_res5.move(670,160)
			if ent_brand6.text() != '':
				lbl_vagon_res6 = QtGui.QLabel(str(count_vagon6), modal_res) 
				lbl_vagon_res6.move(770,160)
			if ent_brand7.text() != '':				
				lbl_vagon_res7 = QtGui.QLabel(str(count_vagon7), modal_res) 
				lbl_vagon_res7.move(870,160)

			lbl_vagon_fist_in = QtGui.QLabel("Первый вагон в печь", modal_res) 
			lbl_vagon_fist_in.move(20,200)
			lbl_vagon_last_in = QtGui.QLabel("Последний вагон в печь", modal_res) 
			lbl_vagon_last_in.move(20,240)
			lbl_vagon_fist_out = QtGui.QLabel("Первый вагон из печи", modal_res) 
			lbl_vagon_fist_out.move(20,280)
			lbl_vagon_last_out = QtGui.QLabel("Последний вагон из печи", modal_res) 
			lbl_vagon_last_out.move(20,320)

			# Сетка для печи (ярлыки)
	
			lbl_furnance = QtGui.QLabel("Туннельная печь №2", modal_res) 
			lbl_furnance.move(20,380)
			
			k = 40
			for i in range(16):
				lbl_pos_furnance = QtGui.QLabel(str("№" + str(i+1)), modal_res) 
				lbl_pos_furnance.move(k,420)
				lbl_pos_furnance = QtGui.QLabel(str(vagon_list_all[i-65][0]), modal_res) 
				lbl_pos_furnance.move(k,460)
				k += 75
			k = 40
			for i in range(16):
				lbl_pos_furnance = QtGui.QLabel(str("№" + str(i+17)), modal_res)
				if i+17 >= 32 and i+17 <=40:
					lbl_pos_furnance.setStyleSheet('color: red') 
				lbl_pos_furnance.move(k,500)
				lbl_pos_furnance = QtGui.QLabel(str(vagon_list_all[i-49][0]), modal_res)
				if i+17 >= 32 and i+17 <=40:
					lbl_pos_furnance.setStyleSheet('color: red')  
				lbl_pos_furnance.move(k,540)
				k += 75
			k = 40
			for i in range(16):
				lbl_pos_furnance = QtGui.QLabel(str("№" + str(i+33)), modal_res) 
				if i+33 >= 32 and i+33 <=40:
					lbl_pos_furnance.setStyleSheet('color: red')
				lbl_pos_furnance.move(k,580)
				lbl_pos_furnance = QtGui.QLabel(str(vagon_list_all[i-33][0]), modal_res)
				if i+33 >= 32 and i+33 <=40:
					lbl_pos_furnance.setStyleSheet('color: red') 
				lbl_pos_furnance.move(k,620)
				k += 75
			k = 40
			for i in range(17):
				lbl_pos_furnance = QtGui.QLabel(str("№" + str(i+49)), modal_res) 
				lbl_pos_furnance.move(k,660)
				lbl_pos_furnance = QtGui.QLabel(str(vagon_list_all[i-17][0]), modal_res) 
				lbl_pos_furnance.move(k,700)
				k += 75
			
		




			if len(time_list) == 4:
				lbl_vagon_fist_in1 = QtGui.QLabel(str(time_list[0]), modal_res) 
				lbl_vagon_fist_in1.move(270,200)
				lbl_vagon_last_in1 = QtGui.QLabel(str(time_list[1]), modal_res) 
				lbl_vagon_last_in1.move(270,240)
				lbl_vagon_fist_out1 = QtGui.QLabel(str(time_list[2]), modal_res) 
				lbl_vagon_fist_out1.move(270,280)
				lbl_vagon_last_out1 = QtGui.QLabel(str(time_list[3]), modal_res) 
				lbl_vagon_last_out1.move(270,320)
			else:
				pass
			


			modal_res.show()








		# Кнопка запуска рассчета режима печи
		btn_count = QtGui.QPushButton("Рассчитать", self)
		btn_count.move(1250,720)
		btn_count.connect(btn_count, QtCore.SIGNAL("clicked()"), count_furnance)
		

app = QtGui.QApplication(sys.argv)
main = MainWindow()	
main.show()
sys.exit(app.exec_())
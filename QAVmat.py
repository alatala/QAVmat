"""QAVmat helps to analyse RA QA tests fast and effectively
Copyright (C) 2015 Agata Latala <a.latala@zfm.coi.pl>

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <http://www.gnu.org/licenses/>.
"""


import sys
from PyQt4 import QtCore, QtGui
from PIL import Image, ImageQt
import dicom
import numpy
from MyDicom1 import MDicom
from MyAnalysisSet import MyAnalysisSet
import xlwt
import pyqtgraph

#Main window class
class MWindow(QtGui.QMainWindow):

#Main window constructor
    def __init__(self):
        super(MWindow, self).__init__()
        self.initUI()

#Addition function for constructor
    def initUI(self):

        # Time interval for timers (in milisec.). Timers are used in sliders for pictures and graph
        self.SliderTimeInterval=10

        # Slider parameters
        self.SliderMinimumValue=1
        self.SliderMaximumValue=300
        self.SliderInterval=1

        # Setting timers for smooth refreshing both pictures and graph. For each timer 10ms delay time is set.
        self.TimerDicom1=QtCore.QTimer()
        self.TimerDicom1.setSingleShot(True)
        self.TimerDicom1.setInterval(self.SliderTimeInterval)
        self.TimerDicom1.timeout.connect(self.RefreshDicom1)
        self.TimerDicom2=QtCore.QTimer()
        self.TimerDicom2.setSingleShot(True)
        self.TimerDicom2.setInterval(self.SliderTimeInterval)
        self.TimerDicom2.timeout.connect(self.RefreshDicom2)
        self.TimerNormalization=QtCore.QTimer()
        self.TimerNormalization.setSingleShot(True)
        self.TimerNormalization.setInterval(self.SliderInterval)
        self.TimerNormalization.timeout.connect(self.RefreshGraph)
        self.TimerSegmentWidth=QtCore.QTimer()
        self.TimerSegmentWidth.setSingleShot(True)
        self.TimerSegmentWidth.setInterval(self.SliderInterval)
        self.TimerSegmentWidth.timeout.connect(self.RefreshSegmentsWidth)

        # Class in which all data are stored
        self.AnalysisSet=MyAnalysisSet()

        # Main widget in which all widgets are placed
        mainWidget=QtGui.QWidget()

        # Main window parameters setting
        self.statusBar()
        self.setWindowTitle('QAVmat')
        self.resize(1100,700)

        # Pictures for displaying image files
        self.picture1=QtGui.QLabel()
        self.picture2=QtGui.QLabel()
        self.picture1.resize(512,384)
        self.picture2.resize(512,384)
        self.picture1.setPixmap(self.AnalysisSet.Dicom1.GetPixmap())
        self.picture2.setPixmap(self.AnalysisSet.Dicom2.GetPixmap())

        # Sliders for setting picture windows
        self.Slider1W=QtGui.QSlider(0x1)
        self.Slider1C=QtGui.QSlider(0x1)
        self.Slider2W=QtGui.QSlider(0x1)
        self.Slider2C=QtGui.QSlider(0x1)

        # Tables in which results of T3, T2, Dosimetry and T3 profile tests are displayed
        self.TableT3H=QtGui.QTableWidget(5,5)
        self.TableT2H=QtGui.QTableWidget(5,8)
        self.TableDosimetry=QtGui.QTableWidget(1,2)
        self.TableProfileT3H=QtGui.QTableWidget(41,4)
        self.TableProfileT3H.setSelectionBehavior(QtGui.QAbstractItemView.SelectRows)
        self.TableProfileT3H.setSelectionMode(QtGui.QAbstractItemView.SingleSelection)
        self.TableProfileT2H=QtGui.QTableWidget(41,7)
        self.TableProfileT2H.setSelectionBehavior(QtGui.QAbstractItemView.SelectRows)
        self.TableProfileT2H.setSelectionMode(QtGui.QAbstractItemView.SingleSelection)

        self.TableT2H.hide()
        self.TableDosimetry.hide()
        self.TableProfileT3H.hide()
        self.TableProfileT2H.hide()

        # Labels with name, date and time of test picture and reference picture
        # Date and time are set after reading dicom files
        self.LabelTestPicture=QtGui.QLabel()
        self.LabelTestPicture.setText("Test picture")
        self.LabelTestPictureDate=QtGui.QLabel()
        self.LabelTestPictureTime=QtGui.QLabel()

        self.LabelReferencePicture=QtGui.QLabel()
        self.LabelReferencePicture.setText("Reference (open) picture")
        self.LabelReferencePictureDate=QtGui.QLabel()
        self.LabelReferencePictureTime=QtGui.QLabel()

        # Labels with values set by sliders
        self.LabelWindowWidthTestPicture=QtGui.QLabel()
        self.LabelWindowWidthTestPicture.setText("Window width = "+str(self.AnalysisSet.Dicom1.WindowWidth)+" CU")
        self.LabelWindowWidthReferencePicture=QtGui.QLabel()
        self.LabelWindowWidthReferencePicture.setText("Window width = "+str(self.AnalysisSet.Dicom2.WindowWidth)+" CU")
        self.LabelWindowCenterTestPicture=QtGui.QLabel()
        self.LabelWindowCenterTestPicture.setText("Window center = "+str(self.AnalysisSet.Dicom1.WindowCenter)+" CU")
        self.LabelWindowCenterReferencePicture=QtGui.QLabel()
        self.LabelWindowCenterReferencePicture.setText(str("Window center = ")+str(self.AnalysisSet.Dicom2.WindowCenter)+" CU")

        # Main vertical box
        BoxMainV=QtGui.QVBoxLayout()

        # Horizontal box for pictures' labels (pictures description)
        BoxPictureLabels=QtGui.QHBoxLayout()
        BoxPictureLabels.addWidget(self.LabelTestPicture)
        BoxPictureLabels.addWidget(self.LabelTestPictureDate)
        BoxPictureLabels.addWidget(self.LabelTestPictureTime)
        BoxPictureLabels.addWidget(self.LabelReferencePicture)
        BoxPictureLabels.addWidget(self.LabelReferencePictureDate)
        BoxPictureLabels.addWidget(self.LabelReferencePictureTime)

        # Horizontal box for pictures
        BoxPicturesH=QtGui.QHBoxLayout()
        BoxPicturesH.addWidget(self.picture1)
        BoxPicturesH.addWidget(self.picture2)

        # Graphics window for displaying graph
        self.MyGraphicsWindow=pyqtgraph.GraphicsWindow()
        self.Plot=pyqtgraph.PlotItem()
        self.MyGraphicsWindow.addItem(self.Plot)
        self.MyGraphicsWindow.setVisible(False)

        self.ReferenceGraph=None
        self.TestNormalisedGraph=None
        self.Maximum=None
        self.Minimum=None

        self.T2Range1=None
        self.T2Range2=None
        self.T2Range3=None
        self.T2Range4=None
        self.T2Range5=None
        self.T2Range6=None
        self.T2Range7=None

        self.T3Range1=None
        self.T3Range2=None
        self.T3Range3=None
        self.T3Range4=None

        self.rangeT3=[self.T3Range1, self.T3Range2, self.T3Range3,  self.T3Range4]
        self.rangeT2=[self.T2Range1, self.T2Range2, self.T2Range3,  self.T2Range4, self.T2Range5, self.T2Range6,  self.T2Range7]

        # Label with normalization factor
        self.NormalizationLabel=QtGui.QLabel()
        self.NormalizationLabel.setText("Normalization factor = "+str(round(self.AnalysisSet.NormalizationFactor,4)))
        self.NormalizationLabel.hide()

        # Slider for changing normalization factor
        self.NormalizationSlider=QtGui.QSlider(0x1)
        self.NormalizationSlider.setMinimum(1)
        self.NormalizationSlider.setMaximum(10000)
        self.NormalizationSlider.setValue(self.AnalysisSet.NormalizationFactor)
        self.NormalizationSlider.setTickInterval(1)
        self.NormalizationSlider.sliderMoved.connect(self.ChangeNormalization)
        self.NormalizationSlider.hide()

        # Label with segment width
        self.SegmentWidthLabel=QtGui.QLabel()
        self.SegmentWidthLabel.setText("Segment width = " +str(self.AnalysisSet.PATwidth)+" pixels (~"+str(round(self.AnalysisSet.PATwidth*0.0392,2))+"cm)")
        self.SegmentWidthLabel.hide()

        #Slider for changing segment width
        self.SegmentWidthSlider=QtGui.QSlider(0x1)
        self.SegmentWidthSlider.setMinimum(1)
        self.SegmentWidthSlider.setMaximum(100)
        self.SegmentWidthSlider.setValue(self.AnalysisSet.PATwidth)
        self.SegmentWidthSlider.setTickInterval(1)
        self.SegmentWidthSlider.sliderMoved.connect(self.ChangeSegmentWidth)
        self.SegmentWidthSlider.hide()

        # Vertical box for graph, labels and sliders
        BoxGraphV=QtGui.QVBoxLayout()
        BoxGraphV.addWidget(self.MyGraphicsWindow)
        BoxGraphV.addWidget(self.NormalizationLabel)
        BoxGraphV.addWidget(self.NormalizationSlider)
        BoxGraphV.addWidget(self.SegmentWidthLabel)
        BoxGraphV.addWidget(self.SegmentWidthSlider)

        # Horizontal box for labels connected with sliders which set windows width
        BoxSlidersWidthLabels=QtGui.QHBoxLayout()
        BoxSlidersWidthLabels.addWidget(self.LabelWindowWidthTestPicture)
        BoxSlidersWidthLabels.addWidget(self.LabelWindowWidthReferencePicture)

        # Horizontal box for sliders which set windows width
        BoxSlidersWidthH=QtGui.QHBoxLayout()
        BoxSlidersWidthH.addWidget(self.Slider1W)
        BoxSlidersWidthH.addWidget(self.Slider2W)

        # Horizontal box for labels connected with sliders which set windows centre
        BoxSlidersCenterLabels=QtGui.QHBoxLayout()
        BoxSlidersCenterLabels.addWidget(self.LabelWindowCenterTestPicture)
        BoxSlidersCenterLabels.addWidget(self.LabelWindowCenterReferencePicture)

        # Horizontal box for sliders which set windows centre
        BoxSlidersCenterH=QtGui.QHBoxLayout()
        BoxSlidersCenterH.addWidget(self.Slider1C)
        BoxSlidersCenterH.addWidget(self.Slider2C)

        # Horizontal box for table with T3 test results
        BoxTableT3H=QtGui.QHBoxLayout()
        BoxTableT3H.addWidget(self.TableT3H)

        # Horizontal box for table with T2 test results
        BoxTableT2H=QtGui.QHBoxLayout()
        BoxTableT2H.addWidget(self.TableT2H)

        # Horizontal box for table with dosimetry test results
        BoxTableDosimetry=QtGui.QHBoxLayout()
        BoxTableDosimetry.addWidget(self.TableDosimetry)

        # Horizontal box for table with profile T3 test results
        BoxTableProfileT3H=QtGui.QHBoxLayout()
        BoxTableProfileT3H.addWidget(self.TableProfileT3H)

        #Horizontal box for table with profile T3 test results
        BoxTableProfileT2H=QtGui.QHBoxLayout()
        BoxTableProfileT2H.addWidget(self.TableProfileT2H)

        # Adding all boxes to the main box
        BoxMainV.addLayout(BoxPictureLabels)
        BoxMainV.addLayout(BoxPicturesH)
        BoxMainV.addLayout(BoxGraphV)
        BoxMainV.addLayout(BoxSlidersWidthLabels)
        BoxMainV.addLayout(BoxSlidersWidthH)
        BoxMainV.addLayout(BoxSlidersCenterLabels)
        BoxMainV.addLayout(BoxSlidersCenterH)
        BoxMainV.addLayout(BoxTableT3H)
        BoxMainV.addLayout(BoxTableT2H)
        BoxMainV.addLayout(BoxTableDosimetry)
        BoxMainV.addLayout(BoxTableProfileT3H)
        BoxMainV.addLayout(BoxTableProfileT2H)

        mainWidget.setLayout(BoxMainV)
        self.setCentralWidget(mainWidget)
        self.RefreshLabels()

        self.show()

        # Actions from file menu
        openFile1 = QtGui.QAction('with a test picture', self)
        openFile1.setStatusTip('Open a new file with a test picture')
        openFile1.triggered.connect(self.showDialog1)

        openFile2 = QtGui.QAction('with a reference (open) picture', self)
        openFile2.setStatusTip('Open a new file with a reference (open) picture (for T2 and T3 tests)')
        openFile2.triggered.connect(self.showDialog2)

        # Actions from calculate menu
        calculate0=QtGui.QAction('Dosimetry test', self)
        calculate0.setStatusTip('Calculate dosimetry test (on the basis of the test picture only)')
        calculate0.triggered.connect(self.CalculateDosimetry)

        calculate1 = QtGui.QAction('T2 test', self)
        calculate1.setStatusTip('Calculate values for T2 test')
        calculate1.triggered.connect(self.CalculateT2)

        calculate2 = QtGui.QAction('T3 test', self)
        calculate2.setStatusTip('Calculate values for T3 test')
        calculate2.triggered.connect(self.CalculateT3)

        calculate3= QtGui.QAction('Profile analysis for T2 test',self)
        calculate3.setStatusTip('Perform profile analysis for T2 test')
        calculate3.triggered.connect(self.ProfileAnalysisT2)

        calculate4= QtGui.QAction('Profile analysis for T3 test',self)
        calculate4.setStatusTip('Perform profile analysis for T3 test')
        calculate4.triggered.connect(self.ProfileAnalysisT3)

        # Actions from save menu
        saveReport1=QtGui.QAction('as a new .xls file', self)
        saveReport1.setStatusTip('Save report as a new .xls file')
        saveReport1.triggered.connect(self.saveReport1)

        # Adding actions to open file menu
        menubar = self.menuBar()
        fileMenu = menubar.addMenu('&Open file')
        fileMenu.addAction(openFile1)
        fileMenu.addAction(openFile2)

        # Adding actions to calculate menu
        calculateMenu = menubar.addMenu('&Calculate')
        calculateMenu.addAction(calculate0)
        calculateMenu.addAction(calculate1)
        calculateMenu.addAction(calculate2)
        calculateMenu.addAction(calculate3)
        calculateMenu.addAction(calculate4)

        # Adding actions to report menu
        reportMenu=menubar.addMenu('&Save report')
        reportMenu.addAction(saveReport1)

        # Setting sliders' parameters and connections
        self.Slider1W.setMinimum(self.SliderMinimumValue)
        self.Slider1W.setMaximum(self.SliderMaximumValue)
        self.Slider1W.setTickInterval(self.SliderInterval)
        self.Slider1W.setValue(self.AnalysisSet.Dicom1.WindowWidth)

        self.Slider1C.setMinimum(self.SliderMinimumValue)
        self.Slider1C.setMaximum(self.SliderMaximumValue)
        self.Slider1C.setTickInterval(self.SliderInterval)
        self.Slider1C.setValue(self.AnalysisSet.Dicom1.WindowCenter)

        self.Slider2W.setMinimum(self.SliderMinimumValue)
        self.Slider2W.setMaximum(self.SliderMaximumValue)
        self.Slider2W.setTickInterval(self.SliderInterval)
        self.Slider2W.setValue(self.AnalysisSet.Dicom2.WindowWidth)

        self.Slider2C.setMinimum(self.SliderMinimumValue)
        self.Slider2C.setMaximum(self.SliderMaximumValue)
        self.Slider2C.setTickInterval(self.SliderInterval)
        self.Slider2C.setValue(self.AnalysisSet.Dicom2.WindowCenter)

        self.Slider1W.valueChanged.connect(self.ChangeWindowDicom1)
        self.Slider1C.valueChanged.connect(self.ChangeWindowDicom1)
        self.Slider2W.valueChanged.connect(self.ChangeWindowDicom2)
        self.Slider2C.valueChanged.connect(self.ChangeWindowDicom2)

        #Action for selection row in table for T2 test profile analysis
        self.TableProfileT3H.selectionModel().selectionChanged.connect(self.PlotProfilesSet)
        self.TableProfileT2H.selectionModel().selectionChanged.connect(self.PlotProfilesSet)

    def closeEvent(self, event):
        reply = QtGui.QMessageBox.question(self, 'Message', "Are you sure to quit?", QtGui.QMessageBox.Yes | QtGui.QMessageBox.No, QtGui.QMessageBox.No)
        if reply == QtGui.QMessageBox.Yes:
                event.accept()
        else:
                event.ignore()

# This function is called when you choose: open file -> with a test picture
# The function reads dicom file, scales matrix according to sliders settings, shows rescaled picture in test picture window
# and zeroes all calculations connected with that file.
    def showDialog1(self):
        fname = QtGui.QFileDialog.getOpenFileName(self, 'Open file with a test picture','','Dicom Files (*.dcm)')
        if len(fname)==0:
            return
        else:
            self.AnalysisSet.Dicom1.ReadDicom(fname)
            self.AnalysisSet.Dicom1.ScaleMatrix()
            self.picture=self.AnalysisSet.Dicom1.GetPixmap()
            self.picture1.setPixmap(self.picture)
            self.AnalysisSet.Dicom1.FindLeavesCentre()
            self.show()
            self.AnalysisSet.PercentT2=numpy.zeros(7,float)
            self.AnalysisSet.Percent2MeanT2=numpy.zeros(7,float)
            self.AnalysisSet.PercentT3=numpy.zeros(4,float)
            self.AnalysisSet.Percent2MeanT3=numpy.zeros(4,float)
            self.AnalysisSet.Dicom1.MeanT2=numpy.zeros(7,float)
            self.AnalysisSet.Dicom1.MeanT3=numpy.zeros(4,float)
            self.AnalysisSet.MADtestT2[0]=0.0
            self.AnalysisSet.MADtestT2[2]=0.0
            self.AnalysisSet.MADtestT2[3]=0.0
            self.AnalysisSet.MADtestT3[0]=0.0
            self.AnalysisSet.MADtestT3[2]=0.0
            self.AnalysisSet.MADtestT3[3]=0.0
            self.AnalysisSet.DosimetryTestValue=0.0
            self.AnalysisSet.MaximumDifferenceForProfileAnalysisT3=numpy.zeros((40,4))
            self.AnalysisSet.MaximumDifferenceForProfileAnalysisT3ForAllLeaves=numpy.zeros(4,float)
            self.AnalysisSet.MaximumDifferenceForProfileAnalysisT2=numpy.zeros((40,7))
            self.AnalysisSet.MaximumDifferenceForProfileAnalysisT2ForAllLeaves=numpy.zeros(7,float)
            self.ShowSliders()
            self.RefreshLabels()

# This function is called when you choose: open file -> with a reference picture
# The function reads dicom file, scales matrix according to sliders settings, shows rescaled picture in reference picture window
# and zeroes all calculations connected with that file.
    def showDialog2(self):
        fname = QtGui.QFileDialog.getOpenFileName(self, 'Open file with a reference picture','','Dicom Files (*.dcm)')
        if len(fname)==0:
            return
        else:
            self.AnalysisSet.Dicom2.ReadDicom(fname)
            self.AnalysisSet.Dicom2.ScaleMatrix()
            self.picture2.setPixmap(self.AnalysisSet.Dicom2.GetPixmap())
            self.show()
            self.AnalysisSet.PercentT2=numpy.zeros(7,float)
            self.AnalysisSet.Percent2MeanT2=numpy.zeros(7,float)
            self.AnalysisSet.PercentT3=numpy.zeros(4,float)
            self.AnalysisSet.Percent2MeanT3=numpy.zeros(4,float)
            self.AnalysisSet.Dicom2.MeanT2=numpy.zeros(7,float)
            self.AnalysisSet.Dicom2.MeanT3=numpy.zeros(4,float)
            self.AnalysisSet.MADtestT2[1]=0.0
            self.AnalysisSet.MADtestT2[2]=0.0
            self.AnalysisSet.MADtestT2[3]=0.0
            self.AnalysisSet.MADtestT3[1]=0.0
            self.AnalysisSet.MADtestT3[2]=0.0
            self.AnalysisSet.MADtestT3[3]=0.0
            self.AnalysisSet.MaximumDifferenceForProfileAnalysisT3=numpy.zeros((40,4))
            self.AnalysisSet.MaximumDifferenceForProfileAnalysisT3ForAllLeaves=numpy.zeros(4,float)
            self.AnalysisSet.MaximumDifferenceForProfileAnalysisT2=numpy.zeros((40,7))
            self.AnalysisSet.MaximumDifferenceForProfileAnalysisT2ForAllLeaves=numpy.zeros(7,float)
            self.ShowSliders()
            self.RefreshLabels()

# This function is called when you choose: save file -> as a new .xls file
# The function creates a simple report for T2, T3, dosimetry and profile T3 test
    def saveReport1(self):
        column1=40        #width of column 0
		
        column2=15        #width of column with data
        fname = QtGui.QFileDialog.getSaveFileName(self, 'Save file','','(*.xls)')
        if len(fname)==0:
            return
        else:
            workbook=xlwt.Workbook()

            if(self.AnalysisSet.testT2==True):
                worksheet=workbook.add_sheet('T2 test')
            if(self.AnalysisSet.testT3==True):
                worksheet=workbook.add_sheet('T3 test')
            if(self.AnalysisSet.DosimetryTest==True):
                worksheet=workbook.add_sheet('Dosimetry test')
            if(self.AnalysisSet.ProfileAnalysisT3==True):
                worksheet=workbook.add_sheet('Profile analysis T3')
            if(self.AnalysisSet.ProfileAnalysisT2==True):
                worksheet=workbook.add_sheet('Profile analysis T2')

            worksheet.col(9).width=256*column1
            worksheet.write(7,9,"Test picture date")
            worksheet.write(7,10,self.AnalysisSet.Dicom1.Date)
            worksheet.write(8,9,"Test picture time")
            worksheet.write(8,10,self.AnalysisSet.Dicom1.Time)
            worksheet.write(9,9,"Reference picture date")
            worksheet.write(10,9,"Reference picture time")

            if((self.AnalysisSet.testT2==True)or(self.AnalysisSet.testT3==True)or(self.AnalysisSet.ProfileAnalysisT3==True)or(self.AnalysisSet.ProfileAnalysisT2==True)):
                worksheet.write(9,10,self.AnalysisSet.Dicom2.Date)
                worksheet.write(10,10,self.AnalysisSet.Dicom2.Time)
            if(self.AnalysisSet.DosimetryTest==True):
                worksheet.write(9,10,"-")
                worksheet.write(10,10,"-")

            if self.AnalysisSet.testT3==True:

                worksheet.col(0).width=256*column1
                for i in range(1, 6):
                    worksheet.col(i).width=256*column2

                RowTitleT3=["segment 1","segment 2","segment 3","segment 4","MAD"]
                ColumnTitleT3=["mean value for test picture [CU]","mean value for reference (open) picture [CU]","test to open ratio [%]","ratio to mean [%]","tolerance [%]"]
                RowToleranceT3=["< +/-3","< +/-3","< +/-3","< +/-3","< +/-1.5"]

                for i, item in enumerate(RowTitleT3):
                    worksheet.write(0,i+1,item)
                for i, item in enumerate(ColumnTitleT3):
                    worksheet.write(i+1,0,item)
                for i, item in enumerate(self.AnalysisSet.Dicom1.MeanT3):
                    worksheet.write(1,i+1,item)
                for i, item in enumerate(self.AnalysisSet.Dicom2.MeanT3):
                    worksheet.write(2,i+1,item)
                for i, item in enumerate(self.AnalysisSet.PercentT3):
                    worksheet.write(3,i+1,item)
                for i, item in enumerate(self.AnalysisSet.Percent2MeanT3):
                    worksheet.write(4,i+1,item)
                for i, item in enumerate(self.AnalysisSet.MADtestT3):
                    worksheet.write(1+i,5,item)
                for i, item in enumerate(RowToleranceT3):
                    worksheet.write(5,i+1,item)

            if self.AnalysisSet.testT2==True:
                worksheet.col(0).width=256*column1
                for i in range(1, 9):
                    worksheet.col(i).width=256*column2
                RowTitleT2=["segment 1","segment 2","segment 3","segment 4","segment 5","segment 6","segment 7","MAD"]
                ColumnTitleT2=["mean value for test picture [CU]","mean value for reference (open) picture [CU]","test to open ratio [%]","ratio to mean [%]", "tolerance [%]"]
                RowToleranceT2=["< +/-3","< +/-3","< +/-3","< +/-3","< +/-3","< +/-3","< +/-3","< +/-1.5"]
                for i, item in enumerate(RowTitleT2):
                    worksheet.write(0,i+1,item)
                for i, item in enumerate(ColumnTitleT2):
                    worksheet.write(i+1,0,item)
                for i, item in enumerate(self.AnalysisSet.Dicom1.MeanT2):
                    worksheet.write(1,i+1,item)
                for i, item in enumerate(self.AnalysisSet.Dicom2.MeanT2):
                    worksheet.write(2,i+1,item)
                for i, item in enumerate(self.AnalysisSet.PercentT2):
                    worksheet.write(3,i+1,item)
                for i, item in enumerate(self.AnalysisSet.Percent2MeanT2):
                    worksheet.write(4,i+1,item)
                for i, item in enumerate(self.AnalysisSet.MADtestT2):
                    worksheet.write(1+i,8,item)
                for i, item in enumerate(RowToleranceT2):
                    worksheet.write(5,i+1,item)

            if self.AnalysisSet.DosimetryTest==True:
                worksheet.col(0).width=256*column1
                worksheet.col(1).width=256*column2
                worksheet.write(0,0,'mean value [CU]')
                worksheet.write(0,1,self.AnalysisSet.DosimetryTestValue)

            if self.AnalysisSet.ProfileAnalysisT3==True:
                worksheet.col(0).width=256*column1
                for i in range(1, 5):
                    worksheet.col(i).width=256*column2
                RowTitleProfileT3=["segment 1","segment 2","segment 3","segment 4"]
                for i, item in enumerate(RowTitleProfileT3):
                    worksheet.write(0,i+1,item)
                ColumnTitleProfileT3=(["maximum dose difference for leaf 1 [%]","maximum dose difference for leaf 2 [%]","maximum dose difference for leaf 3 [%]","maximum dose difference for leaf 4 [%]","maximum dose difference for leaf 5 [%]","maximum dose difference for leaf 6 [%]","maximum dose difference for leaf 7 [%]","maximum dose difference for leaf 8 [%]","maximum dose difference for leaf 9 [%]","maximum dose difference for leaf 10 [%]",
                                       "maximum dose difference for leaf 11 [%]","maximum dose difference for leaf 12 [%]","maximum dose difference for leaf 13 [%]","maximum dose difference for leaf 14 [%]","maximum dose difference for leaf 15 [%]","maximum dose difference for leaf 16 [%]","maximum dose difference for leaf 17 [%]","maximum dose difference for leaf 18 [%]","maximum dose difference for leaf 19 [%]","maximum dose difference for leaf 20 [%]",
                                       "maximum dose difference for leaf 21 [%]","maximum dose difference for leaf 22 [%]","maximum dose difference for leaf 23 [%]","maximum dose difference for leaf 24 [%]","maximum dose difference for leaf 25 [%]","maximum dose difference for leaf 26 [%]","maximum dose difference for leaf 27 [%]","maximum dose difference for leaf 28 [%]","maximum dose difference for leaf 29 [%]","maximum dose difference for leaf 30 [%]",
                                       "maximum dose difference for leaf 31 [%]","maximum dose difference for leaf 32 [%]","maximum dose difference for leaf 33 [%]","maximum dose difference for leaf 34 [%]","maximum dose difference for leaf 35 [%]","maximum dose difference for leaf 36 [%]","maximum dose difference for leaf 37 [%]","maximum dose difference for leaf 38 [%]","maximum dose difference for leaf 39 [%]","maximum dose difference for leaf 40 [%]",
                                       "maximum dose difference for all leaves [%]",])
                for i, item in enumerate(ColumnTitleProfileT3):
                    worksheet.write(i+1,0,item)

                for i in range(0,40):
                    for j in range(0,4):
                        if(self.AnalysisSet.MaximumDifferenceForProfileAnalysisT3[i,j]==0):
                            worksheet.write(i+1,j+1,'-')
                        else:
                            worksheet.write(i+1,j+1,self.AnalysisSet.MaximumDifferenceForProfileAnalysisT3[i,j])
                for i in range(0,4):
                    worksheet.write(41,i+1,self.AnalysisSet.MaximumDifferenceForProfileAnalysisT3ForAllLeaves[i])

            if self.AnalysisSet.ProfileAnalysisT2==True:
                worksheet.col(0).width=256*column1
                for i in range(1, 8):
                    worksheet.col(i).width=256*column2
                RowTitleProfileT2=["segment 1","segment 2","segment 3","segment 4","segment 5","segment 6","segment 7"]
                for i, item in enumerate(RowTitleProfileT2):
                    worksheet.write(0,i+1,item)
                ColumnTitleProfileT2=(["maximum dose difference for leaf 1 [%]","maximum dose difference for leaf 2 [%]","maximum dose difference for leaf 3 [%]","maximum dose difference for leaf 4 [%]","maximum dose difference for leaf 5 [%]","maximum dose difference for leaf 6 [%]","maximum dose difference for leaf 7 [%]","maximum dose difference for leaf 8 [%]","maximum dose difference for leaf 9 [%]","maximum dose difference for leaf 10 [%]",
                                           "maximum dose difference for leaf 11 [%]","maximum dose difference for leaf 12 [%]","maximum dose difference for leaf 13 [%]","maximum dose difference for leaf 14 [%]","maximum dose difference for leaf 15 [%]","maximum dose difference for leaf 16 [%]","maximum dose difference for leaf 17 [%]","maximum dose difference for leaf 18 [%]","maximum dose difference for leaf 19 [%]","maximum dose difference for leaf 20 [%]",
                                           "maximum dose difference for leaf 21 [%]","maximum dose difference for leaf 22 [%]","maximum dose difference for leaf 23 [%]","maximum dose difference for leaf 24 [%]","maximum dose difference for leaf 25 [%]","maximum dose difference for leaf 26 [%]","maximum dose difference for leaf 27 [%]","maximum dose difference for leaf 28 [%]","maximum dose difference for leaf 29 [%]","maximum dose difference for leaf 30 [%]",
                                           "maximum dose difference for leaf 31 [%]","maximum dose difference for leaf 32 [%]","maximum dose difference for leaf 33 [%]","maximum dose difference for leaf 34 [%]","maximum dose difference for leaf 35 [%]","maximum dose difference for leaf 36 [%]","maximum dose difference for leaf 37 [%]","maximum dose difference for leaf 38 [%]","maximum dose difference for leaf 39 [%]","maximum dose difference for leaf 40 [%]",
                                           "maximum dose difference for all leaves [%]",])
                for i, item in enumerate(ColumnTitleProfileT2):
                    worksheet.write(i+1,0,item)
                for i in range(0,40):
                    for j in range(0,7):
                        if(self.AnalysisSet.MaximumDifferenceForProfileAnalysisT2[i,j]==0):
                                worksheet.write(i+1,j+1,'-')
                        else:
                                worksheet.write(i+1,j+1,self.AnalysisSet.MaximumDifferenceForProfileAnalysisT2[i,j])
                for i in range(0,7):
                        worksheet.write(41,i+1,self.AnalysisSet.MaximumDifferenceForProfileAnalysisT2ForAllLeaves[i])

            workbook.save(str(fname))

#This function is called when you change the position of slider connected with dicom1. This function starts a timer and when the timer
#finishes it calls RefreshDicom1 function.
    def ChangeWindowDicom1(self):
        if self.AnalysisSet.Dicom1.Matrix is not None:
            self.TimerDicom1.start()

#This function is called when you change the position of slider connected with dicom2. This function starts a timer and when the timer
#finishes it calls RefreshDicom2 function.
    def ChangeWindowDicom2(self):
        if self.AnalysisSet.Dicom2.Matrix is not None:
            self.TimerDicom2.start()

#This function rescales test picture according to the sliders values and displays a new (rescaled) picture and sliders values
    def RefreshDicom1(self):
        self.AnalysisSet.Dicom1.WindowWidth=self.Slider1W.value()
        self.AnalysisSet.Dicom1.WindowCenter=self.Slider1C.value()

        self.LabelWindowWidthTestPicture.setText("Window width = "+str(self.AnalysisSet.Dicom1.WindowWidth)+" CU")
        self.LabelWindowCenterTestPicture.setText("Window center = "+str(self.AnalysisSet.Dicom1.WindowCenter)+" CU")

        self.AnalysisSet.Dicom1.ScaleMatrix()
        self.picture1.setPixmap(self.AnalysisSet.Dicom1.GetPixmap())
        self.show()

#This function rescales reference picture according to the sliders values and displays a new (rescaled) picture and sliders values
    def RefreshDicom2(self):
        self.AnalysisSet.Dicom2.WindowWidth=self.Slider2W.value()
        self.AnalysisSet.Dicom2.WindowCenter=self.Slider2C.value()

        self.LabelWindowWidthReferencePicture.setText("Window width = "+str(self.AnalysisSet.Dicom2.WindowWidth)+" CU")
        self.LabelWindowCenterReferencePicture.setText("Window center = "+str(self.AnalysisSet.Dicom2.WindowCenter)+" CU")

        self.AnalysisSet.Dicom2.ScaleMatrix()
        self.picture2.setPixmap(self.AnalysisSet.Dicom2.GetPixmap())
        self.show()

#This function is called when you change the position of normalization slider. This function starts a timer and when the timer
#finishes it calls RefreshGraph function.
    def ChangeNormalization(self):
        self.AnalysisSet.NormalizationFactor=float(self.NormalizationSlider.value())/1000
        self.NormalizationLabel.setText("Normalization factor = "+str(round(self.AnalysisSet.NormalizationFactor,3)))
        self.TimerNormalization.start()

#This function refreses graph
    def RefreshGraph(self):
        MaxRange=685
        MinRange=340
        self.AnalysisSet.NormaliseData()
        if(self.AnalysisSet.ProfileAnalysisT3==True):
            self.AnalysisSet.AnalyseProfilesForT3Test()
        if(self.AnalysisSet.ProfileAnalysisT2==True):
            self.AnalysisSet.AnalyseProfilesForT2Test()
        if(self.TestNormalisedGraph==None):
            return
        else:
            self.Plot.removeItem(self.TestNormalisedGraph)
            self.TestNormalisedGraph=pyqtgraph.PlotDataItem(range(MinRange,MaxRange),self.AnalysisSet.Dicom1.MatrixNormalised[self.AnalysisSet.Dicom1.MinimasPosition[self.AnalysisSet.SelectedRow],MinRange:MaxRange],pen=(255,0,0))
            self.Plot.addItem(self.TestNormalisedGraph)
        self.RefreshLabels()

#This function changes segment width
    def ChangeSegmentWidth(self):
        self.AnalysisSet.PATwidth=self.SegmentWidthSlider.value()
        self.SegmentWidthLabel.setText("Segment width = " +str(self.AnalysisSet.PATwidth)+" pixels (~"+str(round(self.AnalysisSet.PATwidth*0.0392,2))+"cm)")
        self.TimerSegmentWidth.start()

    def RefreshSegmentsWidth(self):
        if(self.AnalysisSet.ProfileAnalysisT2==True):
            self.AnalysisSet.AnalyseProfilesForT2Test()
            for i in range(0,7):
                if(self.rangeT2[i]==None):
                    return
                else:
                    self.Plot.removeItem(self.rangeT2[i])
            for i in range(0,7):
                Min=359-int(self.AnalysisSet.PATwidth/2)+i*51
                Max=Min+self.AnalysisSet.PATwidth-1
                self.rangeT2[i]=pyqtgraph.LinearRegionItem(values=[Min,Max], movable=False)
                self.Plot.addItem(self.rangeT2[i],ignoreBounds=True)

        elif(self.AnalysisSet.ProfileAnalysisT3==True):
            self.AnalysisSet.AnalyseProfilesForT3Test()
            for i in range(0,4):
                if(self.rangeT3[i]==None):
                    return
                else:
                    self.Plot.removeItem(self.rangeT3[i])
            for i in range(0,4):
                Min=395-int(self.AnalysisSet.PATwidth/2)+i*77
                Max=Min+self.AnalysisSet.PATwidth-1
                self.rangeT3[i]=pyqtgraph.LinearRegionItem(values=[Min,Max], movable=False)
                self.Plot.addItem(self.rangeT3[i],ignoreBounds=True)
        self.RefreshLabels()

#This function sets the flag 'self.AnalysisSet.testT3' to the TRUE value and the other flags to the False value, calculates values
#for T3 test and refreshes labels in the main window.
    def CalculateT3(self):
        self.AnalysisSet.testT2=False
        self.AnalysisSet.testT3=True
        self.AnalysisSet.DosimetryTest=False
        self.AnalysisSet.ProfileAnalysisT3=False
        self.AnalysisSet.ProfileAnalysisT2=False
        self.AnalysisSet.CalculateT3Test()
        self.ShowSliders()
        self.RefreshLabels()

#This function sets the flag 'self.AnalysisSet.testT2' to the TRUE value and the other flags to the False value, calculates values
#for T2 test and refreshes labels in the main window.
    def CalculateT2(self):
        self.AnalysisSet.testT2=True
        self.AnalysisSet.testT3=False
        self.AnalysisSet.DosimetryTest=False
        self.AnalysisSet.ProfileAnalysisT3=False
        self.AnalysisSet.ProfileAnalysisT2=False
        self.AnalysisSet.CalculateT2Test()
        self.ShowSliders()
        self.RefreshLabels()

#This function sets the flag 'self.AnalysisSet.DosimetryTest' to the TRUE value and the other flags to the False value, calculates value
#for Dosimetry test and refreshes labels in the main window.
    def CalculateDosimetry(self):
        self.AnalysisSet.testT2=False
        self.AnalysisSet.testT3=False
        self.AnalysisSet.DosimetryTest=True
        self.AnalysisSet.ProfileAnalysisT3=False
        self.AnalysisSet.ProfileAnalysisT2=False
        self.AnalysisSet.CalculateDosimetryTest()
        self.RefreshLabels()
        self.ShowSliders()

#This function sets the flag 'self.AnalysisSet.AnalyseProfilesForT2Test' to the TRUE value and the other flags to the False value, calculates values
#for profile analysis for T2 test and refreshes labels in the main window
    def ProfileAnalysisT2(self):
        self.AnalysisSet.testT2=False
        self.AnalysisSet.testT3=False
        self.AnalysisSet.DosimetryTest=False
        self.AnalysisSet.ProfileAnalysisT3=False
        self.AnalysisSet.ProfileAnalysisT2=True
        self.AnalysisSet.CalculateNormalizationFactor()
        self.AnalysisSet.NormaliseData()
        self.NormalizationSlider.setValue(int(self.AnalysisSet.NormalizationFactor*1000))
        self.NormalizationLabel.setText("Normalization factor = "+ str(round(self.AnalysisSet.NormalizationFactor,3)))
        self.AnalysisSet.PATwidth=25
        self.AnalysisSet.AnalyseProfilesForT2Test()
        self.SegmentWidthSlider.setValue(self.AnalysisSet.PATwidth)
        self.SegmentWidthLabel.setText("Segment width = " +str(self.AnalysisSet.PATwidth)+" pixels (~"+str(round(self.AnalysisSet.PATwidth*0.0392,2))+"cm)")
        self.RefreshLabels()
        self.HideSliders()
        self.Plot.clear()

#This function sets the flag 'self.AnalysisSet.AnalyseProfilesForT3Test' to the TRUE value and the other flags to the False value, calculates values
#for profile analysis for T3 test and refreshes labels in the main window.
    def ProfileAnalysisT3(self):
        self.AnalysisSet.testT2=False
        self.AnalysisSet.testT3=False
        self.AnalysisSet.DosimetryTest=False
        self.AnalysisSet.ProfileAnalysisT3=True
        self.AnalysisSet.ProfileAnalysisT2=False
        self.AnalysisSet.CalculateNormalizationFactor()
        self.AnalysisSet.NormaliseData()
        self.NormalizationSlider.setValue(int(self.AnalysisSet.NormalizationFactor*1000))
        self.NormalizationLabel.setText("Normalization factor = "+ str(round(self.AnalysisSet.NormalizationFactor,3)))
        self.AnalysisSet.PATwidth=31
        self.AnalysisSet.AnalyseProfilesForT3Test()
        self.SegmentWidthSlider.setValue(self.AnalysisSet.PATwidth)
        self.SegmentWidthLabel.setText("Segment width = " +str(self.AnalysisSet.PATwidth)+" pixels (~"+str(round(self.AnalysisSet.PATwidth*0.0392,2))+"cm)")
        self.RefreshLabels()
        self.HideSliders()
        self.Plot.clear()

#This function refreshes labels and charts
    def RefreshLabels(self):
        self.LabelTestPictureDate.setText(self.AnalysisSet.Dicom1.Date)
        self.LabelTestPictureTime.setText(self.AnalysisSet.Dicom1.Time)
        self.LabelReferencePictureDate.setText(self.AnalysisSet.Dicom2.Date)
        self.LabelReferencePictureTime.setText(self.AnalysisSet.Dicom2.Time)

        if self.AnalysisSet.testT3==True:

            self.TableT3H.setHorizontalHeaderLabels(["segment 1","segment 2","segment 3","segment 4","MAD"])
            self.TableT3H.setVerticalHeaderLabels(["mean value for test picture [CU]","mean value for reference (open) picture [CU]","test to open ratio [%]","ratio to mean [%]","tolerance [%]"])
            RowToleranceT3=["<"+unichr(0xB1)+"3","<"+unichr(0xB1)+"3","<"+unichr(0xB1)+"3","<"+unichr(0xB1)+"3","<"+unichr(0xB1)+"1.5"]

            for i, item in enumerate(self.AnalysisSet.Dicom1.MeanT3):
                if(item==0):
                    newitem=QtGui.QTableWidgetItem('-')
                    newitem.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                    self.TableT3H.setItem(0, i, newitem)
                else:
                    newitem = QtGui.QTableWidgetItem(str(round(item,4)))
                    newitem.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                    self.TableT3H.setItem(0, i, newitem)

            for i, item in enumerate(self.AnalysisSet.Dicom2.MeanT3):
                if(item==0):
                    newitem=QtGui.QTableWidgetItem('-')
                    newitem.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                    self.TableT3H.setItem(1, i, newitem)
                else:
                    newitem = QtGui.QTableWidgetItem(str(round(item,4)))
                    newitem.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                    self.TableT3H.setItem(1, i, newitem)

            for i, item in enumerate(self.AnalysisSet.PercentT3):
                if(item==0):
                    newitem=QtGui.QTableWidgetItem('-')
                    newitem.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                    self.TableT3H.setItem(2, i, newitem)
                else:
                    newitem = QtGui.QTableWidgetItem(str(round(item,4)))
                    newitem.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                    self.TableT3H.setItem(2, i, newitem)

            for i, item in enumerate(self.AnalysisSet.Percent2MeanT3):
                if(item==0):
                    newitem=QtGui.QTableWidgetItem('-')
                    newitem.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                    self.TableT3H.setItem(3, i, newitem)
                    self.TableT3H.item(3,i).setBackgroundColor(QtGui.QColor(255,255,255))
                elif((item>=3)or(item<=-3)):
                    newitem = QtGui.QTableWidgetItem(str(round(item,4)))
                    newitem.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                    self.TableT3H.setItem(3, i, newitem)
                    self.TableT3H.item(3,i).setBackgroundColor(QtGui.QColor(255,0,0))
                else:
                    newitem = QtGui.QTableWidgetItem(str(round(item,4)))
                    newitem.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                    self.TableT3H.setItem(3, i, newitem)
                    self.TableT3H.item(3,i).setBackgroundColor(QtGui.QColor(0,255,0))

            for i,item in enumerate(self.AnalysisSet.MADtestT3):
                if(item==0):
                    newitem=QtGui.QTableWidgetItem('-')
                    newitem.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                    self.TableT3H.setItem(i,4,newitem)
                    if(i==3):
                        self.TableT3H.item(3,4).setBackgroundColor(QtGui.QColor(255,255,255))
                else:
                    newitem=QtGui.QTableWidgetItem(str(round(item,4)))
                    newitem.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                    self.TableT3H.setItem(i,4,newitem)
                    if(i==3):
                        if((item>=1.5)or(item<=-1.5)):
                            self.TableT3H.item(3,4).setBackgroundColor(QtGui.QColor(255,0,0))
                        else:
                            self.TableT3H.item(3,4).setBackgroundColor(QtGui.QColor(0,255,0))

            for i, item in enumerate(RowToleranceT3):
                newitem=QtGui.QTableWidgetItem(item)
                newitem.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                self.TableT3H.setItem(4,i,newitem)

            self.TableT2H.hide()
            self.TableT3H.show()
            self.TableDosimetry.hide()
            self.TableProfileT3H.hide()
            self.TableProfileT2H.hide()

        if self.AnalysisSet.testT2==True:

            self.TableT2H.setHorizontalHeaderLabels(["segment 1","segment 2","segment 3","segment 4","segment 5","segment 6","segment 7","MAD"])
            self.TableT2H.setVerticalHeaderLabels(["mean value for test picture [CU]","mean value for reference (open) picture [CU]","test to reference ratio [%]","ratio to mean [%]", "tolerance [%]"])
            RowToleranceT2=["<"+unichr(0xB1)+"3","<"+unichr(0xB1)+"3","<"+unichr(0xB1)+"3","<"+unichr(0xB1)+"3","<"+unichr(0xB1)+"3","<"+unichr(0xB1)+"3","<"+unichr(0xB1)+"3","<"+unichr(0xB1)+"1.5"]

            for i, item in enumerate(self.AnalysisSet.Dicom1.MeanT2):
                if(item==0):
                    newitem=QtGui.QTableWidgetItem('-')
                    newitem.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                    self.TableT2H.setItem(0, i, newitem)
                else:
                    newitem = QtGui.QTableWidgetItem(str(round(item,4)))
                    newitem.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                    self.TableT2H.setItem(0, i, newitem)

            for i, item in enumerate(self.AnalysisSet.Dicom2.MeanT2):
                if(item==0):
                    newitem=QtGui.QTableWidgetItem('-')
                    newitem.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                    self.TableT2H.setItem(1, i, newitem)
                else:
                    newitem = QtGui.QTableWidgetItem(str(round(item,4)))
                    newitem.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                    self.TableT2H.setItem(1, i, newitem)

            for i, item in enumerate(self.AnalysisSet.PercentT2):
                if(item==0):
                    newitem=QtGui.QTableWidgetItem('-')
                    newitem.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                    self.TableT2H.setItem(2, i, newitem)
                else:
                    newitem = QtGui.QTableWidgetItem(str(round(item,4)))
                    newitem.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                    self.TableT2H.setItem(2, i, newitem)

            for i, item in enumerate(self.AnalysisSet.Percent2MeanT2):
                if(item==0):
                    newitem=QtGui.QTableWidgetItem('-')
                    newitem.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                    self.TableT2H.setItem(3, i, newitem)
                    self.TableT2H.item(3,i).setBackgroundColor(QtGui.QColor(255,255,255))
                elif((item>=3)or(item<=-3)):
                    newitem = QtGui.QTableWidgetItem(str(round(item,4)))
                    newitem.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                    self.TableT2H.setItem(3, i, newitem)
                    self.TableT2H.item(3,i).setBackgroundColor(QtGui.QColor(255,0,0))
                else:
                    newitem = QtGui.QTableWidgetItem(str(round(item,4)))
                    newitem.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                    self.TableT2H.setItem(3, i, newitem)
                    self.TableT2H.item(3,i).setBackgroundColor(QtGui.QColor(0,255,0))

            for i,item in enumerate(self.AnalysisSet.MADtestT2):
                if(item==0):
                    newitem=QtGui.QTableWidgetItem('-')
                    newitem.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                    self.TableT2H.setItem(i,7,newitem)
                    if(i==3):
                        self.TableT2H.item(3,7).setBackgroundColor(QtGui.QColor(255,255,255))
                else:
                    newitem=QtGui.QTableWidgetItem(str(round(item,4)))
                    newitem.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                    self.TableT2H.setItem(i,7,newitem)
                    if(i==3):
                        if((item>=1.5)or(item<=-1.5)):
                            self.TableT2H.item(3,7).setBackgroundColor(QtGui.QColor(255,0,0))
                        else:
                            self.TableT2H.item(3,7).setBackgroundColor(QtGui.QColor(0,255,0))

            for i, item in enumerate(RowToleranceT2):
                newitem=QtGui.QTableWidgetItem(item)
                newitem.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                self.TableT2H.setItem(4,i,newitem)

            self.TableT3H.hide()
            self.TableT2H.show()
            self.TableDosimetry.hide()
            self.TableProfileT3H.hide()
            self.TableProfileT2H.hide()

        if self.AnalysisSet.DosimetryTest == True:

            newitem=QtGui.QTableWidgetItem('mean value [CU]')
            self.TableDosimetry.setItem(0,0,newitem)
            if(self.AnalysisSet.DosimetryTestValue==0):
                newitem=QtGui.QTableWidgetItem('-')
                newitem.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                self.TableDosimetry.setItem(0,1,newitem)
            else:
                newitem=QtGui.QTableWidgetItem(str(round(self.AnalysisSet.DosimetryTestValue,4)))
                newitem.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                self.TableDosimetry.setItem(0,1,newitem)

            self.TableT3H.hide()
            self.TableT2H.hide()
            self.TableDosimetry.show()
            self.TableProfileT3H.hide()
            self.TableProfileT2H.hide()

        if self.AnalysisSet.ProfileAnalysisT3 == True:

            self.TableProfileT3H.setHorizontalHeaderLabels(["segment 1","segment 2","segment 3","segment 4"])
            self.TableProfileT3H.setVerticalHeaderLabels(["maximum dose difference for leaf 1 [%]","maximum dose difference for leaf 2 [%]","maximum dose difference for leaf 3 [%]","maximum dose difference for leaf 4 [%]","maximum dose difference for leaf 5 [%]","maximum dose difference for leaf 6 [%]","maximum dose difference for leaf 7 [%]","maximum dose difference for leaf 8 [%]","maximum dose difference for leaf 9 [%]","maximum dose difference for leaf 10 [%]",
                                                          "maximum dose difference for leaf 11 [%]","maximum dose difference for leaf 12 [%]","maximum dose difference for leaf 13 [%]","maximum dose difference for leaf 14 [%]","maximum dose difference for leaf 15 [%]","maximum dose difference for leaf 16 [%]","maximum dose difference for leaf 17 [%]","maximum dose difference for leaf 18 [%]","maximum dose difference for leaf 19 [%]","maximum dose difference for leaf 20 [%]",
                                                          "maximum dose difference for leaf 21 [%]","maximum dose difference for leaf 22 [%]","maximum dose difference for leaf 23 [%]","maximum dose difference for leaf 24 [%]","maximum dose difference for leaf 25 [%]","maximum dose difference for leaf 26 [%]","maximum dose difference for leaf 27 [%]","maximum dose difference for leaf 28 [%]","maximum dose difference for leaf 29 [%]","maximum dose difference for leaf 30 [%]",
                                                          "maximum dose difference for leaf 31 [%]","maximum dose difference for leaf 32 [%]","maximum dose difference for leaf 33 [%]","maximum dose difference for leaf 34 [%]","maximum dose difference for leaf 35 [%]","maximum dose difference for leaf 36 [%]","maximum dose difference for leaf 37 [%]","maximum dose difference for leaf 38 [%]","maximum dose difference for leaf 39 [%]","maximum dose difference for leaf 40 [%]",
                                                          "maximum dose difference for all leaves [%]",])

            for i in range (0,40):
                for j in range (0,4):
                    if(self.AnalysisSet.MaximumDifferenceForProfileAnalysisT3[i,j]==0):
                        newitem=QtGui.QTableWidgetItem('-')
                        newitem.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                        self.TableProfileT3H.setItem(i,j,newitem)
                        self.TableProfileT3H.item(i,j).setBackgroundColor(QtGui.QColor(255,255,255))
                    elif((self.AnalysisSet.MaximumDifferenceForProfileAnalysisT3[i,j]>=3)or(self.AnalysisSet.MaximumDifferenceForProfileAnalysisT3[i,j]<=-3)):
                        newitem=QtGui.QTableWidgetItem(str(round(self.AnalysisSet.MaximumDifferenceForProfileAnalysisT3[i,j],4)))
                        newitem.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                        self.TableProfileT3H.setItem(i,j,newitem)
                        self.TableProfileT3H.item(i,j).setBackgroundColor(QtGui.QColor(255,0,0))
                    else:
                        newitem=QtGui.QTableWidgetItem(str(round(self.AnalysisSet.MaximumDifferenceForProfileAnalysisT3[i,j],4)))
                        newitem.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                        self.TableProfileT3H.setItem(i,j,newitem)
                        self.TableProfileT3H.item(i,j).setBackgroundColor(QtGui.QColor(255,255,255))

            for j, item in enumerate(self.AnalysisSet.MaximumDifferenceForProfileAnalysisT3ForAllLeaves):
                if(item==0):
                    newitem=QtGui.QTableWidgetItem('-')
                    newitem.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                    self.TableProfileT3H.setItem(40,j,newitem)
                    self.TableProfileT3H.item(40,j).setBackgroundColor(QtGui.QColor(255,255,255))
                elif((item>=3)or(item<=-3)):
                    newitem=QtGui.QTableWidgetItem(str(round(item,4)))
                    newitem.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                    self.TableProfileT3H.setItem(40,j,newitem)
                    self.TableProfileT3H.item(40,j).setBackgroundColor(QtGui.QColor(255,0,0))
                else:
                    newitem=QtGui.QTableWidgetItem(str(round(item,4)))
                    newitem.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                    self.TableProfileT3H.setItem(40,j,newitem)
                    self.TableProfileT3H.item(40,j).setBackgroundColor(QtGui.QColor(255,255,255))
            self.TableT3H.hide()
            self.TableT2H.hide()
            self.TableDosimetry.hide()
            self.TableProfileT3H.show()
            self.TableProfileT2H.hide()

        if self.AnalysisSet.ProfileAnalysisT2 == True:

            self.TableProfileT2H.setHorizontalHeaderLabels(["segment 1","segment 2","segment 3","segment 4","segment 5","segment 6","segment 7"])
            self.TableProfileT2H.setVerticalHeaderLabels(["maximum dose difference for leaf 1 [%]","maximum dose difference for leaf 2 [%]","maximum dose difference for leaf 3 [%]","maximum dose difference for leaf 4 [%]","maximum dose difference for leaf 5 [%]","maximum dose difference for leaf 6 [%]","maximum dose difference for leaf 7 [%]","maximum dose difference for leaf 8 [%]","maximum dose difference for leaf 9 [%]","maximum dose difference for leaf 10 [%]",
                                                  "maximum dose difference for leaf 11 [%]","maximum dose difference for leaf 12 [%]","maximum dose difference for leaf 13 [%]","maximum dose difference for leaf 14 [%]","maximum dose difference for leaf 15 [%]","maximum dose difference for leaf 16 [%]","maximum dose difference for leaf 17 [%]","maximum dose difference for leaf 18 [%]","maximum dose difference for leaf 19 [%]","maximum dose difference for leaf 20 [%]",
                                                  "maximum dose difference for leaf 21 [%]","maximum dose difference for leaf 22 [%]","maximum dose difference for leaf 23 [%]","maximum dose difference for leaf 24 [%]","maximum dose difference for leaf 25 [%]","maximum dose difference for leaf 26 [%]","maximum dose difference for leaf 27 [%]","maximum dose difference for leaf 28 [%]","maximum dose difference for leaf 29 [%]","maximum dose difference for leaf 30 [%]",
                                                  "maximum dose difference for leaf 31 [%]","maximum dose difference for leaf 32 [%]","maximum dose difference for leaf 33 [%]","maximum dose difference for leaf 34 [%]","maximum dose difference for leaf 35 [%]","maximum dose difference for leaf 36 [%]","maximum dose difference for leaf 37 [%]","maximum dose difference for leaf 38 [%]","maximum dose difference for leaf 39 [%]","maximum dose difference for leaf 40 [%]",
                                                  "maximum dose difference for all leaves [%]",])

            for i in range (0,40):
                for j in range (0,7):
                    if(self.AnalysisSet.MaximumDifferenceForProfileAnalysisT2[i,j]==0):
                        newitem=QtGui.QTableWidgetItem('-')
                        newitem.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                        self.TableProfileT2H.setItem(i,j,newitem)
                        self.TableProfileT2H.item(i,j).setBackgroundColor(QtGui.QColor(255,255,255))
                    elif((self.AnalysisSet.MaximumDifferenceForProfileAnalysisT2[i,j]>=3)or(self.AnalysisSet.MaximumDifferenceForProfileAnalysisT2[i,j]<=-3)):
                        newitem=QtGui.QTableWidgetItem(str(round(self.AnalysisSet.MaximumDifferenceForProfileAnalysisT2[i,j],4)))
                        newitem.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                        self.TableProfileT2H.setItem(i,j,newitem)
                        self.TableProfileT2H.item(i,j).setBackgroundColor(QtGui.QColor(255,0,0))
                    else:
                        newitem=QtGui.QTableWidgetItem(str(round(self.AnalysisSet.MaximumDifferenceForProfileAnalysisT2[i,j],4)))
                        newitem.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                        self.TableProfileT2H.setItem(i,j,newitem)

            for i, item in enumerate(self.AnalysisSet.MaximumDifferenceForProfileAnalysisT2ForAllLeaves):
                if(item==0):
                    newitem=QtGui.QTableWidgetItem('-')
                    newitem.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                    self.TableProfileT2H.setItem(40,i,newitem)
                    self.TableProfileT2H.item(40,i).setBackgroundColor(QtGui.QColor(255,255,255))
                elif((item>=3)or(item<=-3)):
                    newitem=QtGui.QTableWidgetItem(str(round(item,4)))
                    newitem.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                    self.TableProfileT2H.setItem(40,i,newitem)
                    self.TableProfileT2H.item(40,i).setBackgroundColor(QtGui.QColor(255,0,0))
                else:
                    newitem=QtGui.QTableWidgetItem(str(round(item,4)))
                    newitem.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled)
                    self.TableProfileT2H.setItem(40,i,newitem)
                    self.TableProfileT2H.item(40,i).setBackgroundColor(QtGui.QColor(255,255,255))

            self.TableT3H.hide()
            self.TableT2H.hide()
            self.TableDosimetry.hide()
            self.TableProfileT3H.hide()
            self.TableProfileT2H.show()

    #This function hides all stuff connected with sliders
    def HideSliders(self):
        self.LabelWindowWidthTestPicture.hide()
        self.LabelWindowWidthReferencePicture.hide()
        self.LabelWindowCenterTestPicture.hide()
        self.LabelWindowCenterReferencePicture.hide()
        self.Slider1W.hide()
        self.Slider1C.hide()
        self.Slider2W.hide()
        self.Slider2C.hide()
        self.picture1.hide()
        self.picture2.hide()

        self.MyGraphicsWindow.show()
        self.NormalizationLabel.show()
        self.NormalizationSlider.show()
        self.SegmentWidthLabel.show()
        self.SegmentWidthSlider.show()

#This function hides graphs and shows pictures and sliders
    def ShowSliders(self):
        self.MyGraphicsWindow.hide()
        self.NormalizationLabel.hide()
        self.NormalizationSlider.hide()
        self.SegmentWidthLabel.hide()
        self.SegmentWidthSlider.hide()

        self.LabelWindowWidthTestPicture.show()
        self.LabelWindowWidthReferencePicture.show()
        self.LabelWindowCenterTestPicture.show()
        self.LabelWindowCenterReferencePicture.show()
        self.Slider1W.show()
        self.Slider1C.show()
        self.Slider2W.show()
        self.Slider2C.show()
        self.picture1.show()
        self.picture2.show()

#This function plots graph for a profile selected in table
    def PlotProfilesSet(self, selected, deselected):
        if (selected.isEmpty()==False):
            self.AnalysisSet.SelectedRow=selected.first().top()
            self.Plot.clear()
            MaxRange=685
            MinRange=340
            Markers=numpy.zeros(1024,int)

            if((self.AnalysisSet.SelectedRow<40) and (self.AnalysisSet.Dicom1.MinimasPosition[self.AnalysisSet.SelectedRow] <> 0.0)):
                self.ReferenceGraph=pyqtgraph.PlotDataItem(range(MinRange,MaxRange),self.AnalysisSet.Dicom2.Matrix[self.AnalysisSet.Dicom1.MinimasPosition[self.AnalysisSet.SelectedRow],MinRange:MaxRange],pen=(0,255,0))
                self.Plot.addItem(self.ReferenceGraph)
                self.TestNormalisedGraph=pyqtgraph.PlotDataItem(range(MinRange,MaxRange),self.AnalysisSet.Dicom1.MatrixNormalised[self.AnalysisSet.Dicom1.MinimasPosition[self.AnalysisSet.SelectedRow],MinRange:MaxRange],pen=(255,0,0))
                self.Plot.addItem(self.TestNormalisedGraph)
                self.Maximum=pyqtgraph.PlotDataItem(range(MinRange,MaxRange),1.03*self.AnalysisSet.Dicom2.Matrix[self.AnalysisSet.Dicom1.MinimasPosition[self.AnalysisSet.SelectedRow],MinRange:MaxRange],pen=(0,0,255))
                self.Plot.addItem(self.Maximum)
                self.Minimum=pyqtgraph.PlotDataItem(range(MinRange,MaxRange),0.97*self.AnalysisSet.Dicom2.Matrix[self.AnalysisSet.Dicom1.MinimasPosition[self.AnalysisSet.SelectedRow],MinRange:MaxRange],pen=(0,0,255))
                self.Plot.addItem(self.Minimum)

                self.Plot.setLabel('left',"Pixel value",'CU')
                self.Plot.setLabel('bottom',"Pixel number")

            if(self.AnalysisSet.ProfileAnalysisT3==True):
                for i in range(0,4):
                    Min=395-int(self.AnalysisSet.PATwidth/2)+i*77
                    Max=Min+self.AnalysisSet.PATwidth-1
                    self.rangeT3[i]=pyqtgraph.LinearRegionItem(values=[Min,Max], movable=False)
                    self.Plot.addItem(self.rangeT3[i],ignoreBounds=True)

            if(self.AnalysisSet.ProfileAnalysisT2==True):
                for i in range(0,7):
                    Min=359-int(self.AnalysisSet.PATwidth/2)+i*51
                    Max=Min+self.AnalysisSet.PATwidth-1
                    self.rangeT2[i]=pyqtgraph.LinearRegionItem(values=[Min,Max], movable=False)
                    self.Plot.addItem(self.rangeT2[i],ignoreBounds=True)

# Main  function
if __name__ == "__main__":
    app = QtGui.QApplication(sys.argv)
    MW = MWindow()
    MW.show()
    sys.exit(app.exec_())

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

from PIL import Image, ImageQt
import dicom
import numpy
from PyQt4 import QtGui
from PIL import Image, ImageQt
import xlwt

class MDicom:

    def __init__(self):                                         #MDicom constructor

        self.Portal_x_size=1024                                 #Portal x size (width)
        self.Portal_y_size=768                                  #Portal y size (length)
        self.DcFile=None                                        #Dicom file
        self.Matrix=numpy.zeros((self.Portal_x_size,self.Portal_y_size),dtype=float)          #Raw data image matrix
        self.MatrixRescaled=numpy.zeros((self.Portal_x_size,self.Portal_y_size),dtype=float)  #Data image matrix in CU
        self.MatrixNormalised=numpy.zeros((self.Portal_x_size,self.Portal_y_size),dtype=float)  #Data image matrix renormalised to reference picture
        self.WindowCenter=50                                    #Center of scaling window
        self.WindowWidth=200                                    #Width of scaling window
        self.Pixmap=None                                        #Pixmap (generated by function GetPixmapFromDicom)
        self.MeanT3=numpy.zeros(4,float)                        #Mean values for ROIs in T3 test
        self.MeanT2=numpy.zeros(7,float)                        #Mean values for ROIs in T2 test
        self.MeanDosimetry=0.0                                  #Mean value for ROI in dosimetry test
        self.Date=''                                            #Date of picture
        self.Time=''                                            #Time of picture
        self.MinimasPosition=numpy.zeros(40,int)                # Position of minimas (i.e. leaves cenres) on the basis of Dicom1

#This function reads data from Dicom file
    def ReadDicom(self,fname):
        self.DcFile=dicom.read_file(str(fname)) #Attention!
        self.Matrix=self.DcFile.pixel_array
        self.Matrix=self.Matrix*self.DcFile.RescaleSlope+self.DcFile.RescaleIntercept
        self.Date=self.DcFile.ContentDate[0:4]+'-'+self.DcFile.ContentDate[4:6]+'-'+self.DcFile.ContentDate[6:8]
        self.Time=self.DcFile.ContentTime[0:2]+':'+self.DcFile.ContentTime[2:4]

#This function scales matrix for displaying according to the sliders position
    def ScaleMatrix(self):
        MinPixVal=0.0
        MaxPixVal=255.0
        PixRange=255.0

        MinVal = self.WindowCenter - 0.5 - (self.WindowWidth - 1.0) / 2.0
        MaxVal = self.WindowCenter - 0.5 + (self.WindowWidth - 1.0) / 2.0

        min_mask = (MinVal >= self.Matrix)
        to_scale = (self.Matrix > MinVal) & (self.Matrix < MaxVal)
        max_mask = (self.Matrix >= MaxVal)

        self.MatrixRescaled=numpy.copy(self.Matrix)
        if min_mask.any():
            self.MatrixRescaled[min_mask] = MinPixVal
        if to_scale.any():
            self.MatrixRescaled[to_scale] = ((self.Matrix[to_scale] - (self.WindowCenter - 0.5)) /(self.WindowWidth - 1.0) + 0.5) * self.WindowWidth + MinPixVal
        if max_mask.any():
            self.MatrixRescaled[max_mask] = MaxPixVal

 #This function returns pixmap which is needed for displaying picture
    def GetPixmap(self):
        MatrixLocal=numpy.rint(self.MatrixRescaled).astype(numpy.int8)
        obrazek=Image.fromarray(MatrixLocal,'L')
        obrazek=obrazek.resize((512,384))
        obrazek1 = ImageQt.ImageQt(obrazek)
        obrazek2 = QtGui.QImage(obrazek1)
        pixmap = QtGui.QPixmap.fromImage(obrazek2)
        return pixmap

#This function calculates mean values for ROIs in T2,T3, and dosimetry tests
    def CalculateMeanValues(self):

        HelpMatrix1=self.Matrix[255:510,390:403]
        HelpMatrix2=self.Matrix[255:510,466:479]
        HelpMatrix3=self.Matrix[255:510,543:556]
        HelpMatrix4=self.Matrix[255:510,620:633]
        self.MeanT3[0]=HelpMatrix1.mean()
        self.MeanT3[1]=HelpMatrix2.mean()
        self.MeanT3[2]=HelpMatrix3.mean()
        self.MeanT3[3]=HelpMatrix4.mean()

        HelpMatrix21=self.Matrix[255:510,353:366]
        HelpMatrix22=self.Matrix[255:510,404:417]
        HelpMatrix23=self.Matrix[255:510,455:468]
        HelpMatrix24=self.Matrix[255:510,506:519]
        HelpMatrix25=self.Matrix[255:510,557:570]
        HelpMatrix26=self.Matrix[255:510,608:621]
        HelpMatrix27=self.Matrix[255:510,659:672]
        self.MeanT2[0]=HelpMatrix21.mean()
        self.MeanT2[1]=HelpMatrix22.mean()
        self.MeanT2[2]=HelpMatrix23.mean()
        self.MeanT2[3]=HelpMatrix24.mean()
        self.MeanT2[4]=HelpMatrix25.mean()
        self.MeanT2[5]=HelpMatrix26.mean()
        self.MeanT2[6]=HelpMatrix27.mean()

        HelpMatrix3=self.Matrix[263:503,472:552]
        self.MeanDosimetry=HelpMatrix3.mean()

#This function finds leaves centres needed for profile analysis
    def FindLeavesCentre(self):
        T3_x_size=306 #306, but it is not ok for old excell ! BE CAREFULL !
        T3_x_beginning=358
        T3_y_beginning=140
        T3_y_end=630
        First_minimum_segment_beginning=140
        First_minimum_segment_end=153
        Size_of_segment=13

        minimas_table=numpy.zeros((self.Portal_y_size,T3_x_size),int) #in this table, the number of row with minimum is denoted
        count_table=numpy.zeros((self.Portal_y_size,T3_x_size),int) #in this table 1 means the place where minimum is

        for j in range (0,T3_x_size):

            column=self.Matrix[:,T3_x_beginning+j]

            #Here you find all minimas
            for i in range(T3_y_beginning,T3_y_end):
                if((column[i]<column[i-1])and(column[i]<=column[i+1])):
                    minimas_table[i,j]=i
                    count_table[i,j]=1


            #Here you eliminate "additional" minimas
            for i in range(0,760):
                if(count_table[i,j] == 1):
                    minimum_value=column[i]
                    minimum_index=i

                    for p in range(1,8):
                        if((column[i+p]<minimum_value)and(count_table[i+p,j] == 1)):
                            minimas_table[minimum_index,j]=0
                            count_table[minimum_index,j]=0
                            minimum_value=column[i+p]
                            minimum_index=i+p
                        elif((column[i+p]>minimum_value)and(count_table[i+p,j] == 1)): #!!!!tricky!!! rethink
                            minimas_table[i+p,j]=0
                            count_table[i+p,j]=0
                        elif((column[i+p] == minimum_value)and(count_table[i+p,j] == 1)):
                            minimas_table[i+p,j]=0
                            count_table[i+p,j]=0

        for i in range(0,40):

            help_minimas=minimas_table[First_minimum_segment_beginning+Size_of_segment*i:First_minimum_segment_end+Size_of_segment*i,:]
            help_count=count_table[First_minimum_segment_beginning+Size_of_segment*i:First_minimum_segment_end+Size_of_segment*i,:]
            if(help_count.sum()==0):
                self.MinimasPosition[i]=0.0
            else:
                self.MinimasPosition[i]=round(float(help_minimas.sum())/float(help_count.sum()))













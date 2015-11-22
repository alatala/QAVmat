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

from MyDicom1 import MDicom
import numpy

class MyAnalysisSet:

    def __init__(self):

        self.testT2=False                                           # when this flag is true T2 test is performed
        self.testT3=False                                           # when this flag is true T3 test is performed
        self.DosimetryTest=False                                    # when this flag is true Dosimetry test is performed
        self.ProfileAnalysisT3=False                                # when this flag is true profile analysis for T3 test is performed
        self.ProfileAnalysisT2=False                                # when this flag is true profile analysis for T2 test is performed
        self.Dicom1=MDicom()                                        # Dicom file with test picture
        self.Dicom2=MDicom()                                        # Dicom file with reference picture
        self.PercentT3=numpy.zeros(4,float)                         # Ratio of mean pixel values in ROIs in T3 test (test picture to reference picture)
        self.MeanPercentT3=0.0                                      # Mean ratio pixel values in T3 test
        self.Percent2MeanT3=numpy.zeros(4,float)                    # Percent to mean values in T3 test
        self.PercentT2=numpy.zeros(7,float)                         # Ratio of mean pixel values in ROIs in T2 test (test picture to reference picture)
        self.Percent2MeanT2=numpy.zeros(7,float)                    # Percent to mean values in T2 test
        self.MeanPercentT2=0.0                                      # Mean ratio pixel values in T3 test
        self.MADtestT2=[0.0,0.0,0.0,0.0]                            # Mean absolute deviation in T2 test (for mean values of ROIs in test pictures, mean values of ROIs in reference
                                                                    # pictures, Ratio of mean pixel values in ROIs and percent to mean values in T2 test)
        self.MADtestT3=[0.0,0.0,0.0,0.0]                            # Mean absolute deviation in T3 test (for mean values of ROIs in test pictures, mean values of ROIs in reference
                                                                    # pictures, Ratio of mean pixel values in ROIs and percent to mean values in T3 test)
        self.DosimetryTestValue=0.0                                 # Result of dosimetry test
        self.NormalizationFactor=1.0                                # The factor to normalise test data for T2 or T3 profile analysis
        self.MaximumDifferenceForProfileAnalysisT3=numpy.zeros((40,4)) #Results of profile analysis for T3 test
        self.MaximumDifferenceForProfileAnalysisT3ForAllLeaves=numpy.zeros(4,float)
        self.MaximumDifferenceForProfileAnalysisT2=numpy.zeros((40,7)) #Results of profile analysis for T2 test
        self.MaximumDifferenceForProfileAnalysisT2ForAllLeaves=numpy.zeros(7,float)
        self.SelectedRow=None                                         #Last selected row
        self.PATwidth=31                                                #Segment width for profile analysis test

#This function calculates values for T3 test
    def CalculateT3Test(self):
        self.Dicom1.CalculateMeanValues()
        self.Dicom2.CalculateMeanValues()

        if((self.Dicom2.MeanT3[0]==0)or(self.Dicom2.MeanT3[1]==0)or(self.Dicom2.MeanT3[2]==0)or(self.Dicom2.MeanT3[3]==0)):
            return
        else:
            self.PercentT3=numpy.divide(self.Dicom1.MeanT3,self.Dicom2.MeanT3)*100
            self.MeanPercentT3=numpy.mean(self.PercentT3)
            self.Percent2MeanT3=numpy.divide(self.PercentT3,self.MeanPercentT3)*100-100
            self.MADtestT3[0]=numpy.mean(numpy.absolute(self.Dicom1.MeanT3-numpy.mean(self.Dicom1.MeanT3)))
            self.MADtestT3[1]=numpy.mean(numpy.absolute(self.Dicom2.MeanT3-numpy.mean(self.Dicom2.MeanT3)))
            self.MADtestT3[2]=numpy.mean(numpy.absolute(self.PercentT3-numpy.mean(self.PercentT3)))
            self.MADtestT3[3]=numpy.mean(numpy.absolute(self.Percent2MeanT3-numpy.mean(self.Percent2MeanT3)))

#This function calculates values for T2 test
    def CalculateT2Test(self):
        self.Dicom1.CalculateMeanValues()
        self.Dicom2.CalculateMeanValues()

        if((self.Dicom2.MeanT2[0]==0)or(self.Dicom2.MeanT2[1]==0)or(self.Dicom2.MeanT2[2]==0)or(self.Dicom2.MeanT2[3]==0)or(self.Dicom2.MeanT2[4]==0)or(self.Dicom2.MeanT2[5]==0)or(self.Dicom2.MeanT2[6]==0)):
            return
        else:
            self.PercentT2=numpy.divide(self.Dicom1.MeanT2,self.Dicom2.MeanT2)*100
            self.MeanPercentT2=numpy.mean(self.PercentT2)
            self.Percent2MeanT2=numpy.divide(self.PercentT2,self.MeanPercentT2)*100-100
            self.MADtestT2[0]=numpy.mean(numpy.absolute(self.Dicom1.MeanT2-numpy.mean(self.Dicom1.MeanT2)))
            self.MADtestT2[1]=numpy.mean(numpy.absolute(self.Dicom2.MeanT2-numpy.mean(self.Dicom2.MeanT2)))
            self.MADtestT2[2]=numpy.mean(numpy.absolute(self.PercentT2-numpy.mean(self.PercentT2)))
            self.MADtestT2[3]=numpy.mean(numpy.absolute(self.Percent2MeanT2-numpy.mean(self.Percent2MeanT2)))

# This function calculates results for dosimetry test
    def CalculateDosimetryTest(self):
            self.Dicom1.CalculateMeanValues()
            self.DosimetryTestValue=self.Dicom1.MeanDosimetry

# This function calculates normalization factor for T2 and T3 profile test
# Attention dividing by 0 - correct
    def CalculateNormalizationFactor(self):
            if(self.ProfileAnalysisT2==True):
                if(numpy.mean(self.Dicom1.Matrix[379:382,508:512])==0):
                    self.NormalizationFactor=1.0
                else:
                    self.NormalizationFactor=numpy.mean(self.Dicom2.Matrix[379:382,508:512])/numpy.mean(self.Dicom1.Matrix[379:382,508:512]) #Rethink that
            elif(self.ProfileAnalysisT3==True):
                if(numpy.mean(self.Dicom1.Matrix[379:382,537:561])==0):
                    self.NormalizationFactor=1.0
                else:
                    self.NormalizationFactor=numpy.mean(self.Dicom2.Matrix[379:382,537:561])/numpy.mean(self.Dicom1.Matrix[379:382,537:561]) #Rethink that

# This function normalizes test data to reference data (data from open picture)
    def NormaliseData(self):
            self.Dicom1.MatrixNormalised=self.NormalizationFactor*self.Dicom1.Matrix

# This function analyse profiles for T3 test. It:
#1. Takes all rows marked as leaves centre (from normalised and reference pictures)
#2. Calculates difference between normalised and reference profile (in %) in the centres (~1.2cm width) of 4 segments
#3. Finds the maximum (absolute) difference for each segment and leaf
#4. Finds the maximum (absolute) difference for all leaves in each segment
    def AnalyseProfilesForT3Test(self):

            #self.PATwidth=31 #default value for profile analysis for T3 test

            for i in range(0,40):
                if(self.Dicom1.MinimasPosition[i]==0):
                    for j in range(0,4):
                        self.MaximumDifferenceForProfileAnalysisT3[i,j]=0

                else:
                    TestRow=self.Dicom1.MatrixNormalised[self.Dicom1.MinimasPosition[i],:]
                    ReferenceRow=self.Dicom2.Matrix[self.Dicom1.MinimasPosition[i],:]

                    segmentHelp=numpy.zeros((4,self.PATwidth), dtype=float)

                    for k in range(0,4):
                        Min=395-int(self.PATwidth/2)+k*77
                        for j in range(0,self.PATwidth):
                            segmentHelp[k,j]=((TestRow[Min+j]-ReferenceRow[Min+j])/ReferenceRow[Min+j])*100

                    for k in range(0,4):
                            help_max=numpy.max(segmentHelp[k,:])
                            help_min=numpy.min(segmentHelp[k,:])
                            if(numpy.abs(help_max)>=numpy.abs(help_min)):
                                self.MaximumDifferenceForProfileAnalysisT3[i,k]=help_max
                            else:
                                self.MaximumDifferenceForProfileAnalysisT3[i,k]=help_min

            for i in range(0,4):
                help_max=numpy.max(self.MaximumDifferenceForProfileAnalysisT3[:,i])
                help_min=numpy.min(self.MaximumDifferenceForProfileAnalysisT3[:,i])
                if(numpy.abs(help_max)>=numpy.abs(help_min)):
                    self.MaximumDifferenceForProfileAnalysisT3ForAllLeaves[i]=help_max
                else:
                    self.MaximumDifferenceForProfileAnalysisT3ForAllLeaves[i]=help_min

# This function analyse profiles for T2 test. It:
#1. Takes all rows marked as leaves centre (from normalised and reference pictures)
#2. Calculates difference between normalised and reference profile (in %) in the centres (~0.98cm width) of 7 segments
#3. Finds the maximum (absolute) difference for each segment and leaf
#4. Finds the maximum (absolute) difference for all leaves in each segment
    def AnalyseProfilesForT2Test(self):

        #self.PATwidth=25 #Default value for profile analysis for T2 test

        for i in range(0,40):
            if(self.Dicom1.MinimasPosition[i]==0):
                for j in range(0,7):
                    self.MaximumDifferenceForProfileAnalysisT2[i,j]=0
            else:
                TestRow=self.Dicom1.MatrixNormalised[self.Dicom1.MinimasPosition[i],:]
                ReferenceRow=self.Dicom2.Matrix[self.Dicom1.MinimasPosition[i],:]

                segmentHelp=numpy.zeros((7,self.PATwidth), dtype=float)

                for k in range(0,7):
                    Min=359-int(self.PATwidth/2)+k*51
                    for j in range(0,self.PATwidth):
                        segmentHelp[k,j]=((TestRow[Min+j]-ReferenceRow[Min+j])/ReferenceRow[Min+j])*100

                for k in range(0,7):
                    help_max=numpy.max(segmentHelp[k,:])
                    help_min=numpy.min(segmentHelp[k,:])
                    if(numpy.abs(help_max)>=numpy.abs(help_min)):
                        self.MaximumDifferenceForProfileAnalysisT2[i,k]=help_max
                    else:
                        self.MaximumDifferenceForProfileAnalysisT2[i,k]=help_min

        for i in range(0,7):
            help_max=numpy.max(self.MaximumDifferenceForProfileAnalysisT2[:,i])
            help_min=numpy.min(self.MaximumDifferenceForProfileAnalysisT2[:,i])
            if(numpy.abs(help_max)>=numpy.abs(help_min)):
                self.MaximumDifferenceForProfileAnalysisT2ForAllLeaves[i]=help_max
            else:
                self.MaximumDifferenceForProfileAnalysisT2ForAllLeaves[i]=help_min

# encoding: utf8
# !/usr/bin/python
# ------------------------------------------------------
# Name:        PlanCheck vs 9b
# Purpose:     Check RS plan op verschillende punten
# Author:      M. Elford
# Created:     28-08-2018 voor RS vs 9b
# ------------------------------------------------------
# 28-08-2018:  methode toegevoegd om vsim plan te vergelijken met huidige plan
# 30-08-2018:  methode toegevoegd om structuren met densitiet override en vorm override te detecteren
# 30-08-2018:  methode toegevoegd om te controleren of vmat plan geyawed moet worden
# 03-09-2018:  Gui is geanchored
# 10-09-2018:  check toegevoegd om te controleren of final dose aanstaat bij een vmat plan
# 10-09-2018:  als de grid op 0.1cm staat mislukt de dosis check. Dit is opgelost door hem dan weg te laten
# 20-09-2018:  fix toegevoegd als de bundels gerenummerd worden
# 20-09-2018:  check toegevoegd om te controleren of intermediate dose aanstaat bij een imrt plan
# 15-10-2018:  check toegevoegd om te controleren of bundel is gekoppeld aan DSP punt
# 01-11-2018:  bug opgelost in de vsim vergelijking als vergeleken wordt met willekeurige plan
# 15-11-2018:  methode aangepast waarmee verschillende beamsets in 1 plan gechecked kunnen worden
# 03-12-2018:  methode toegevoegd om precription site te controleren
# 23-03-2019:  check toegevoegd voor machine, dosis algoritme en clinical status. Check voor recentste
#              CT en GUI in WPF.
# 16-09-2019:  Als 2 beamsets aanwezig zijn dan zoekt hij niet naar Edge setting(ivm conebeam verificatie)
# 16-09-2019:  Als 2 beamsets aanwezig zijn dan moet het isoc gelijk zijn voor beide beamsets
# 16-09-2019:  Max leaf travel mag groter zijn dan 0.4 voor prostaat hypo
# 01-10-2019:  In zeldzame gevallen wordt max dosis target afgerond op meer cijfers dan max dosis ext
# 03-12-2019:  checks voor edge en mip verwijderd aangezien dit nu door conebeam wordt overgenomen
# 04-12-2019:  Max leaf travel check vereenvoudigd
# 04-12-2019:  Bij overschrijding Agility leaf limieten wordt aangegeven waar de overschrijding is.
# 27-01-2020:  script aangepast voor RS9b. >1000MU regel alleen voor vmat ingesteld
# 16-03-2020:  logger toegevoegd
# 15-09-2020:  bug opgelost van leaf motion check
# 26-10-2020:  Nieuw machine toegevoegd (RIF Agility IEC)
# 15-12-2020:  Controle toegevoegd om merged beams te identificeren
# 01-03-2021:  Standaard melding over globale Dmax toegevoegd
# 22-03-2021:  Geschikt gemaakt voor de controle van meerdere beamsets
# 22-03-2021:  Alles weggehaald betreffende epi-plannen
# 01-04-2021:  Controle of 'SIB' in plannaam voorkomt
# 10-06-2021:  Leesteken control toegevoegd voor beamset naam en plannaam
# 08-07-2021:  diverse bugfixes
# 19-08-2021:  Jaw check efficienter gemaakt, 'ma'+'mam'toegevoegd in de filter van plannamen voor mamma planningen
# 19-08-2021:  regel 498: als ct serie anoniem is crashed het zoeken naar een CT decription, opgelost
# 19-08-2021:  get beamenergie routine aangepast voor meerdere beamsets
# 29-10-2021:  max dosis bepaling aangepast per target
# 10-01-2022:  geschikt gemaakt voor co-optimized beamsets
# 19-05 2022:  controle toegevoegd om overbodige plannamen te verwijderen
# 19-05 2022:  controle toegevoegd waarbij intermidiate dose bij vmat wordt waargenomen
# 19-05 2022:  veel overbodige variabelen zijn opgeruimd en code logischer gemaakt
# 19-05 2022:  LOG_AAN uitgezet
# 27-05-2022:  Rekengrid wordt getoond in de gui
# 27-05-2022:  Controle op, als bolus bestaat, dit wordt gebruikt in de bundels
# 27-05-2022:  Controle toegevoegd voor gantryarcspacing voor vmat bundels
# 27-05-2022:  Controle toegevoegd op overbodige plannen
# 06-06-2022:  Controle toegevoegd voor dosestatistics
import datetime
import math
import re
import string
import sys
import wpf

from operator import attrgetter

import clr
from System.Windows import Application, Window, MessageBox

clr.AddReference("System.Xml")
from System.IO import StringReader
from System.Xml import XmlReader

LOG_AAN = False


# TODO:

# tafeliso restricties toevoegen
# lijst maken van standaard doelgebieden per energie
class MyWindow(Window):
    def __init__(self):
        xaml = """
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="{Binding WindowTitle}"
        ResizeMode="CanResizeWithGrip"
        WindowStyle="ThreeDBorderWindow"
        MinWidth="370"
        MaxWidth="420"
        Style="{DynamicResource DefaultWindow}"
        Width="Auto"
        Height="Auto"
        SizeToContent="WidthAndHeight">
    <Grid Background="{DynamicResource {x:Static SystemColors.ControlBrushKey}}">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="1*" />
            <ColumnDefinition Width="1*" />
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>
        <StackPanel Grid.Column="0"
                    Grid.Row="0">
            <GroupBox Header="Plan CT"
                      BorderThickness="2"
                      BorderBrush="LightBlue"
                      HorizontalAlignment="Left"
                      Background="White"
                      VerticalAlignment="Top"
                      Margin="20,5"
                      Padding="5"
                      Height="Auto"
                      Width="130">
                <StackPanel>
                    <TextBox Name="label1"
                             HorizontalAlignment="Left"
                             VerticalAlignment="Top"
                             Width="auto"
                             BorderThickness="0"
                             Margin="3,3" />
                    <TextBox Name="label2"
                             HorizontalAlignment="Left"
                             VerticalAlignment="Top"
                             Width="auto"
                             BorderThickness="0"
                             Margin="3,3" />
                    <TextBox Name="label3"
                             HorizontalAlignment="Left"
                             VerticalAlignment="Top"
                             Width="auto"
                             BorderThickness="0"
                             Margin="3,3" />
                </StackPanel>
            </GroupBox>
            <StackPanel Grid.Column="0"
                        Grid.Row="0">
                <GroupBox Header="Techniek"
                          BorderThickness="2"
                          BorderBrush="LightBlue"
                          HorizontalAlignment="Left"
                          Background="White"
                          VerticalAlignment="Bottom"
                          Margin="20,5"
                          Padding="5"
                          Height="auto">
                    <StackPanel>
                        <TextBox Name="label5"
                                 HorizontalAlignment="Left"
                                 VerticalAlignment="Top"
                                 Width="auto"
                                 BorderThickness="0"
                                 Margin="3,3" />
                        <TextBox Name="label7"
                                 HorizontalAlignment="Left"
                                 VerticalAlignment="Top"
                                 Width="auto"
                                 BorderThickness="0"
                                 Margin="3,3" />
                        <TextBox Name="label8"
                                 HorizontalAlignment="Left"
                                 VerticalAlignment="Top"
                                 Width="auto"
                                 BorderThickness="0"
                                 Margin="3,3" />
                    </StackPanel>
                </GroupBox>
            </StackPanel>
        </StackPanel>
        <StackPanel Grid.Column="1"
                    Grid.Row="0">
            <Label Margin="0,5"
                   Content="Selecteer andere plan:" />
            <ComboBox x:Name="combobox"
                      SelectedIndex="-1"
                      HorizontalAlignment="Center"
                      VerticalAlignment="Top"
                      Width="100"
                      IsReadOnly="True"
                      SelectionChanged="OnChanged"
                      IsEditable="False" />
            <Label  Margin="0,5"
                    Content="Vergelijk met VSIM plan:" />
            <ComboBox x:Name="combobox2"
                      SelectedIndex="-1"
                      HorizontalAlignment="Center"
                      VerticalAlignment="Bottom"
                      Width="100"
                      SelectionChanged="OnChanged2"
                      IsReadOnly="True"
                      IsEditable="False" />
        </StackPanel>
        <StackPanel Grid.ColumnSpan="2"
                    Grid.Row="1"
                    Margin="10">
            <Border BorderThickness="1"
                    BorderBrush="Black"
                    Background="AntiqueWhite" 
                    CornerRadius="5">
                <TextBlock x:Name="DoseBox"
                           Padding="10"
                           FontStyle="Italic"
                           TextAlignment="Justify"
                       Foreground="Blue"
                       TextWrapping="Wrap"
                       HorizontalAlignment="Center" />
            </Border>
        </StackPanel>
        <StackPanel Grid.ColumnSpan="2"
                    Grid.Row="2">
            <TextBox Name="textbox4"
                     Margin="3"
                     BorderThickness="2"
                     BorderBrush="LightBlue"
                     HorizontalAlignment="Stretch"
                     VerticalAlignment="Top"
                     TextWrapping="Wrap"
                     Background="White"
                     AcceptsTab="True"
                     AcceptsReturn="True"
                     HorizontalScrollBarVisibility="Disabled"
                     VerticalScrollBarVisibility="Auto"
                     Height="222" />
        </StackPanel>
        <UniformGrid Rows="1"
                     Columns="1"
                     Margin="5"
                     Grid.ColumnSpan="2"
                     Grid.Row="3">
            <Button Content="Stop"
                    Padding="20,3"
                    HorizontalAlignment="Center"
                    Click="go2" />
        </UniformGrid>
    </Grid>
</Window>
"""

        xr = XmlReader.Create(StringReader(xaml))
        wpf.LoadComponent(self, xr)
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        MainLoop()
        self.Title = 'PlanCheck: ' + str(plan.Name)
        self.label1.Text = bset.PatientPosition
        self.label2.Text = bset.PatientSetup.LocalizationPoiGeometrySource.OnExamination.Name
        self.label3.Text = test_slice_thickness(exam)
        self.label5.Text = str(get_all_tech(plan)).strip('[]')
        self.label7.Text = str(getenergie(plan)).strip('[]') + GetAnn(gettech(bset))
        self.label8.Text = "Grid: " + str(["{0:0.2f}".format(i) for i in gridsize_test(bset)]).strip('[]')
        self.DoseBox.Text = test19(bset, plan, case)
        self.textbox4.Text = curmsg
        plannamen = [pl.Name for pl in case.TreatmentPlans]
        plannamen.insert(0, "")
        self.combobox.ItemsSource = plannamen
        self.combobox2.ItemsSource = plannamen

    def OnChanged(self, sender, event):
        patient.Save()
        curmsg = ""
        self.textbox4.Clear()
        self.label1.Text = ''
        self.label2.Text = ''
        self.label3.Text = ''
        self.label5.Text = ''
        self.label7.Text = ''
        self.label8.Text = ''
        self.DoseBox.Text = ''
        if self.combobox.SelectedValue != '':
            Selected_PlanName = self.combobox.SelectedItem
        else:
            return
        plan_to_load = case.TreatmentPlans[Selected_PlanName]
        plan_to_load.SetCurrent()
        self.textbox4.Text = MainLoop()
        self.label1.Text = bset.PatientPosition
        self.label2.Text = bset.PatientSetup.LocalizationPoiGeometrySource.OnExamination.Name
        self.label3.Text = test_slice_thickness(exam)
        self.label5.Text = str(get_all_tech(plan)).strip('[]')
        self.label7.Text = str(getenergie(plan)).strip('[]') + GetAnn(gettech(bset))
        self.label8.Text = "Grid: " + str(["{0:0.2f}".format(i) for i in gridsize_test(bset)]).strip('[]')
        self.DoseBox.Text = test19(bset, plan, case)
        self.Title = 'PlanCheck: ' + plan.Name

    def OnChanged2(self, sender, event):
        global VsimVlag, VSim_Beam_Set
        VsimVlag = True
        curmsg = ""
        if self.combobox2.SelectedValue != '':
            VSim_Beam_Set = case.TreatmentPlans[self.combobox2.SelectedItem].BeamSets[0]
        else:
            return
        self.textbox4.Clear()
        self.label1.Text = ''
        self.label2.Text = ''
        self.label3.Text = ''
        self.label5.Text = ''
        self.label7.Text = ''
        self.label8.Text = ''
        self.DoseBox.Text = ''
        self.textbox4.Text = MainLoop()
        self.label1.Text = bset.PatientPosition
        self.label2.Text = bset.PatientSetup.LocalizationPoiGeometrySource.OnExamination.Name
        self.label3.Text = test_slice_thickness(exam)
        self.label5.Text = str(get_all_tech(plan)).strip('[]')
        self.label7.Text = str(getenergie(plan)).strip('[]') + GetAnn(gettech(bset))
        self.label8.Text = "Grid: " + str(["{0:0.2f}".format(i) for i in gridsize_test(bset)]).strip('[]')
        self.Title = 'PlanCheck: ' + plan.Name
        VsimVlag = False

    def go2(self, sender, event):
        self.Close()
        sys.exit()


# nodige defs toevoegen
def FindPlanOpt(fplan, fbset):
    OptIndex = 0
    IndexNotFound = True
    while IndexNotFound:
        try:
            OptName = fplan.PlanOptimizations[OptIndex].OptimizedBeamSets[fbset.DicomPlanLabel].DicomPlanLabel
            IndexNotFound = False
        except Exception:
            IndexNotFound = True
            OptIndex += 1
    if IndexNotFound:
        sys.exit("Kan optimalisatie van beamset niet vinden")
    else:
        opt = fplan.PlanOptimizations[OptIndex]
    return opt


# -------------------------------------------------------------------------------------------------
# Function: fGetAverageDensityValues(sRoiName,hPlan)
# Description: Calculates the average density in a ROI by summation of the dosegrid voxels
#
# Arguments: 
#   - sRoiName: Name of the ROI to be evaluated
#   - hPlan: plan on which the density grid is defined
# -------------------------------------------------------------------------------------------------
def fGetAverageDensityValues(sRoiName, hPlan):
    # get the density values
    density = [d for d in hPlan.TreatmentCourse.TotalDose.OnDensity.DensityValues.DensityData]
    # get the dosegrid representation of the ROI
    RoiDoseGridVoxels = hPlan.TreatmentCourse.TotalDose.GetDoseGridRoi(RoiName=sRoiName)
    # calculate average, min and max density value over all voxels
    fAverageDensity = 0.0
    iVoxelIndices = RoiDoseGridVoxels.RoiVolumeDistribution.VoxelIndices
    fRelativeVolumes = RoiDoseGridVoxels.RoiVolumeDistribution.RelativeVolumes
    for i, v in zip(iVoxelIndices, fRelativeVolumes):
        fAverageDensity += density[i] * v
    # return values
    return fAverageDensity


def CheckVsimPlan(fbset, fVSim_Beam_Set):
    # #############################################################
    # ##  code om VSIM plan te vergelijken met geladen plan    ####
    # #############################################################
    fcurmsg = ""
    for iCounter in range(fVSim_Beam_Set.Beams.Count):
        VerkeerdVeldVlag = False
        Gevonden = False
        if fVSim_Beam_Set.Beams[iCounter].Segments.Count == 0:
            fcurmsg += "VSIM bundel %s bevat geen leafs en jaws.  \r\n" % fVSim_Beam_Set.Beams[iCounter].Name
            continue
        for bCounter in range(fbset.Beams.Count):
            Gevonden = False
            beam = fbset.Beams[bCounter]
            if beam.GantryAngle == fVSim_Beam_Set.Beams[iCounter].GantryAngle:
                Gevonden = True
                if fVSim_Beam_Set.Beams[iCounter].Isocenter.Position.x != beam.Isocenter.Position.x \
                        or fVSim_Beam_Set.Beams[iCounter].Isocenter.Position.y != beam.Isocenter.Position.y \
                        or fVSim_Beam_Set.Beams[iCounter].Isocenter.Position.z != beam.Isocenter.Position.z:
                    VerkeerdVeldVlag = True
                    fcurmsg += "Isoc verschilt voor bundel:  %s\r\n" % beam.Name
                if fVSim_Beam_Set.Beams[iCounter].CouchRotationAngle != beam.CouchRotationAngle:
                    VerkeerdVeldVlag = True
                    fcurmsg += "Tafel iso rotatie verschilt voor bundel:  %s\r\n" % beam.Name
                    # print beam.Segments.Count
                if fVSim_Beam_Set.Beams[iCounter].Segments.Count == 1:
                    if fVSim_Beam_Set.Beams[iCounter].Segments[0].CollimatorAngle != beam.Segments[0].CollimatorAngle:
                        VerkeerdVeldVlag = True
                        fcurmsg += "Colimator hoek verschilt voor bundel:  %s\r\n" % beam.Name
                    if fVSim_Beam_Set.Beams[iCounter].Segments[0].JawPositions != beam.Segments[0].JawPositions:
                        VerkeerdVeldVlag = True
                        fcurmsg += "Jaw positie verschilt voor bundel:  %s\r\n" % beam.Name
                    if fVSim_Beam_Set.Beams[iCounter].Segments[0].LeafPositions != beam.Segments[0].LeafPositions:
                        VerkeerdVeldVlag = True
                        fcurmsg += "Leaf posities verschillen voor bundel:  %s\r\n" % beam.Name
                break
            if Gevonden:
                fcurmsg += "Kan geen overeenkomende bundel vinden voor: %s \r\n" % fVSim_Beam_Set.Beams[iCounter].Name
                VerkeerdVeldVlag = True
        if VerkeerdVeldVlag != False or Gevonden != True:
            fcurmsg += "Kan geen overeenkomende bundel vinden voor: %s \r\n" % fVSim_Beam_Set.Beams[iCounter].Name
        else:
            fcurmsg += "Vergelijking klopt voor bundel:  %s\r\n" % beam.Name

    return fcurmsg


def gridsize_test(fbeamset):
    grid_size = [0, 0, 0]
    for doses in fbeamset.FractionDose.BeamDoses:
        current_grid = doses.InDoseGrid.VoxelSize
        grid_size = [max(grid_size[0], current_grid.x),
                     max(grid_size[1], current_grid.y),
                     max(grid_size[2], current_grid.z)]
    return grid_size


def GetAnn(ftechniek):
    # global FFF
    if ftechniek == 'Elektronen':
        ann = ' MeV'
    else:
        ann = ' MV'
    if FFF:
        ann = ann + " FFF"
    return ann


def test_slice_thickness(fexamination):
    image_stack = fexamination.Series[0].ImageStack
    if len(image_stack.SlicePositions) > 1:
        plak_dikte = round(abs(image_stack.SlicePositions[1] - image_stack.SlicePositions[0]), 2)
        text = "Slice dikte: " + str(plak_dikte) + " cm"
    return text


def letter_to_index(letter):
    _alphabet = 'abcdefghijklmnopqrstuvwxyz'
    return next((i + 1 for (i, _letter) in enumerate(_alphabet) if _letter == letter), None)


def find_external_and_targets(fcase):
    # Functie die de namen van de rois vindt gedefinieert als target en external
    # Retourneert de namen in een list
    external_name = None
    target_names = []
    for roi in fcase.PatientModel.RegionsOfInterest:
        if roi.Type == 'External':
            external_name = roi.Name
        elif roi.Type[-2:] == 'tv':
            target_names.append(roi.Name)
            print "target names: ", target_names
    return [external_name, target_names]


def get_change(current, previous):
    # #######################################################################
    # ##  Code om een percentage verschil tussen 2 waardes te verkrijgen ####
    # #######################################################################
    if current == previous:
        return 0
    try:
        return (abs((current - previous) / previous)) * 100.0
    except ZeroDivisionError:
        return 100


def is_max_dosis_in_target(fplan, fcase):
    MaxDoseBuitenPTV = False
    # Functie die controleert of de maximale dosis voor het plan zich in een doelgebied bevindt
    # Retourneert 'True' als de maximale dosis van de external roi hetzelfde is als de maximale dosis in een target
    # roi en anders 'False'
    [external, targets] = find_external_and_targets(fcase)
    external_max_dose = fplan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName=external, DoseType="Max")
    if len(targets) > 0:
        target_max_dose = [fplan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName=name, DoseType="Max") for name in
                           targets]
    try:
        if get_change(max(target_max_dose), external_max_dose) > 0.00001:
            MaxDoseBuitenPTV = True
    except:
        pass
    return MaxDoseBuitenPTV


def getenergie(fplan):
    Energielist = []
    CurrentBeam = 0
    for bs in fplan.BeamSets:
        while bs.Beams.Count > CurrentBeam:
            ActiveBeam = bs.Beams[CurrentBeam]
            BmEnergie = int(ActiveBeam.BeamQualityId.split()[0])
            if BmEnergie not in Energielist:
                Energielist.append(BmEnergie)
            CurrentBeam += 1
    return Energielist


def getenergiebset(fbset):
    Energielist = []
    CurrentBeam = 0
    while fbset.Beams.Count > CurrentBeam:
        ActiveBeam = fbset.Beams[CurrentBeam]
        BmEnergie = int(ActiveBeam.BeamQualityId.split()[0])
        if BmEnergie not in Energielist:
            Energielist.append(BmEnergie)
        CurrentBeam += 1
    return Energielist


def get_all_tech(fplan):
    # ###################################################
    # ##  code om techniek van alle bsets te bepalen  ###
    # ###################################################
    ftechniek = []
    for bs in fplan.BeamSets:
        if bs.DeliveryTechnique == 'SMLC' \
                and bs.PlanGenerationTechnique == 'Conformal':
            ftechniek.append('3DCRT')
        if bs.DeliveryTechnique == 'SMLC' \
                and bs.PlanGenerationTechnique == 'Imrt':
            ftechniek.append('IMRT')
        if bs.DeliveryTechnique == 'DynamicArc' and bs.PlanGenerationTechnique == 'Imrt':
            ftechniek.append('VMAT')
        if bs.Modality == 'Electrons':
            ftechniek.append('Elektronen')
    return ftechniek


def gettech(fbset):
    # ########################################################
    # ##  code om techniek te bepalen van beamset         ####
    # ########################################################
    ftechniek = ''
    if fbset.DeliveryTechnique == 'SMLC' \
            and fbset.PlanGenerationTechnique == 'Conformal':
        ftechniek = '3DCRT'
    if fbset.DeliveryTechnique == 'SMLC' \
            and fbset.PlanGenerationTechnique == 'Imrt':
        ftechniek = 'IMRT'
    if fbset.DeliveryTechnique == 'DynamicArc' and fbset.PlanGenerationTechnique == 'Imrt':
        ftechniek = 'VMAT'
    if fbset.Modality == 'Electrons':
        ftechniek = 'Elektronen'
    return ftechniek


def test0(fcase):
    # ##############################################################
    # ##  code om te bepalen of de recentste CT gebruikt wordt  ####
    # ##############################################################
    msg0 = "\r\n"
    meer_recent = None
    primairCT = get_current("Examination")
    try:
        primair_datum = datetime.datetime(primairCT.GetExaminationDateTime()).date()
    except:
        msg0 = '-' + 'Er is geen primair CT datum aanwezig.\r\n'
        return msg0
    for examination in fcase.Examinations:
        try:
            CT_datum = datetime.datetime(examination.GetExaminationDateTime()).date()
            if CT_datum > primair_datum:
                description = examination.GetAcquisitionDataFromDicom()['SeriesModule']['SeriesDescription']
                frame_uid = examination.EquipmentInfo.FrameOfReference
                modality = examination.EquipmentInfo.Modality
                if modality == 'CT' and 'CBCT' not in description and frame_uid != primairCT.EquipmentInfo.FrameOfReference:
                    meer_recent = examination.Name
        except:
            pass
    if meer_recent:
        msg0 = '-' + 'Er is een recentere CT serie aanwezig.\r\n'
    print "0"
    return msg0


def test1(fpatient, fplan, fss):
    # ##################################################################################################
    # ##  Code om te controlereen of de patientID, dosisalgoritme, dosestatistics en machine klopt  ####
    # ##################################################################################################
    msg1 = ""
    regCompiled = re.compile('^\d{6}[A-Z]{2}\d[A-Z]')
    if not re.match(regCompiled, fpatient.PatientID):
        msg1 += '-' + 'Klopt de patient ID?\r\n'
    dose_stats_bijwerken = [geom.OfRoi.Name for geom in fss.RoiGeometries if geom.HasContours() and fplan.TreatmentCourse.TotalDose.GetDoseGridRoi(RoiName=geom.OfRoi.Name).RoiVolumeDistribution is None]
    if dose_stats_bijwerken:
        msg1 += "- Dose statistics moet geupdate worden.\r\n"
    for bs in fplan.BeamSets:
        try:
            if not bs.FractionDose.DoseValues.AlgorithmProperties.DoseAlgorithm in ['CCDose', 'ElectronMonteCarlo']:
                msg1 += '-' + '<' + bs.DicomPlanLabel + '>' + ' Verkeerde dosis algoritme gebruikt\r\n'
        except:
            pass
        try:
            if not bs.MachineReference.MachineName in ('RIF Agility FFF', 'RIF Agility IEC'):
                msg1 += '-' + '<' + bs.DicomPlanLabel + '>' + ' Verkeerde machine gebruikt\r\n'
        except:
            pass
        try:
            if not bs.FractionDose.DoseValues.IsClinical:
                msg1 += '-' + '<' + bs.DicomPlanLabel + '>' + ' Dosis status is niet clinical\r\n'
        except:
            pass

    print "1"
    return msg1


def test2(fss, fcase, fplan):
    # ############################################################################################
    # ##  Code om ss op lege structuren, form override en densiteit override te controleren  #####
    # ############################################################################################
    msg2 = ""
    DensRoi = {'Tafel', 'HH tafel', 'contrast', 'Bolus', 'opbouw', 'Opbouw', 'skinflash', 'EXT min kunstheup',
               '_PTV in lucht', 'External met dichtheid 1', '_PTVin lucht'}
    LeegContoursVlag = False
    DensiteitOverrideVlag = False
    FormOverrideVlag = False
    BolusAan = False
    BolusGevonden = False
    DensiteitCurrent = []
    LeegCurrent = []
    FormOverride = []

    for current in fss.RoiGeometries:
        if fcase.PatientModel.RegionsOfInterest[current.OfRoi.Name].Type in ('Organ', 'Avoidance') and \
                fcase.PatientModel.RegionsOfInterest[current.OfRoi.Name].OrganData.OrganType == 'Target':
            msg2 += '-' + 'ROI: %s is type: Organ maar Organtype: Target\r\n' % current.OfRoi.Name
        if not current.PrimaryShape:
            LeegContoursVlag = True
            LeegCurrent.append(current.OfRoi.Name)
        if current.OfRoi.RoiMaterial is not None and current.OfRoi.Name not in DensRoi and current.PrimaryShape:
            DensiteitOverrideVlag = True
            DensiteitCurrent.append(current.OfRoi.Name)
        try:
            if current.PrimaryShape.DerivedRoiStatus.IsShapeDirty:
                FormOverrideVlag = True
                FormOverride.append(current.OfRoi.Name)
            if not current.PrimaryShape and current.OfRoi.DerivedRoiExpression:
                FormOverrideVlag = True
                FormOverride.append(current.OfRoi.Name)
        except:
            pass

        if current.OfRoi.Type == 'Bolus':
            BolusGevonden = True
            print 'bolus: ', current.OfRoi.Name
            for bs in fplan.BeamSets:
                for b in bs.Beams:
                    if b.Boli.Count > 0:
                        BolusAan = True
                        break
                if BolusAan:
                    break

    if not BolusAan and BolusGevonden:
        msg2 += '-' + 'Er is bolus gevonden maar geen bijbehorende beams\r\n'
    if LeegContoursVlag:
        msg2 += '-' + 'Er zijn geen geometrieen in ROI: %s\r\n' \
                % ", ".join(LeegCurrent)
    if DensiteitOverrideVlag:
        msg2 += '-' + 'Er is een onverwachte densiteit override in ROI: %s\r\n' \
                % ", ".join(DensiteitCurrent)
    if FormOverrideVlag:
        msg2 += '-' + 'Override is nodig voor ROI: %s\r\n' \
                % ", ".join(FormOverride)

    print "2"
    return msg2


def test3(fscan_naam, fss):
    # ##############################################
    # ##  Code om het tafelblad te controleren  ####
    # ##############################################
    msg3 = ""
    TafelRoi = {'Tafel', 'HH tafel'}
    TafelAanwezig = False
    for current in fss.RoiGeometries:
        if current.OfRoi.Name in TafelRoi and current.PrimaryShape:
            TafelAanwezig = True
    if not TafelAanwezig:
        msg3 += '-' + "Er is geen tafelblad in CT:  %s\r\n" % fscan_naam

    print "3"
    return msg3


def test4(fplan):
    # ##############################################
    # ##  Code om het aantal isoc's te bepalen  ####
    # ##############################################
    global Isoc_Multi, NumIso
    msg4 = ""
    Isoc = []
    Isoc_Multi = []
    NumIso = 1

    for bs in fplan.BeamSets:
        for b in bs.Beams:
            Isoc.append([b.Isocenter.Position.x, b.Isocenter.Position.y, b.Isocenter.Position.z])
    Isoc_Multi.append(Isoc[0])

    for bm in range(Isoc.Count):
        if Isoc[bm] in Isoc_Multi:
            continue
        else:
            Isoc_Multi.append(Isoc[bm])
            NumIso += 1

    if NumIso != 1:
        msg4 += '-' + "Er zijn %d verschillende isoc's in plan. Oligo plan? \r\n" \
                % NumIso

    print "4"
    return msg4


def test5(fNumPOI, fcase):
    # #################################################################
    # ##  Code om het aantal localization points te bepalen        ####
    # #################################################################
    msg5 = ""
    Aantal_LocPoints = 0
    for i in range(fNumPOI):
        if fcase.PatientModel.PointsOfInterest[i].Type \
                == 'LocalizationPoint':
            Aantal_LocPoints += 1

    if Aantal_LocPoints != 1:
        msg5 += '-' + 'Er is geen Localization Point\r\n'

    print "5"
    return msg5


def test6(fbset):
    # #################################################################
    # ##  Code om VMAT plan te controleren of gejawed is           ####
    # #################################################################
    msg6 = ""
    VmatJawVlag = False
    for beam in fbset.Beams:
        Jaw_1 = 0.0
        Jaw_2 = 0.0
        for segment in beam.Segments:
            Jaw_1 = segment.JawPositions[2]
            Jaw_2 = segment.JawPositions[3]
            if Jaw_1 % 0.5 != 0 or Jaw_2 % 0.5 != 0:
                VmatJawVlag = True
                break
        if VmatJawVlag:
            break
    if not VmatJawVlag:
        msg6 += '-' + '<' + fbset.DicomPlanLabel + '>' + ': VMAT beamset moet gejawed worden.\r\n'

    print "6"
    return msg6


def test7(fplan):
    # #################################################################
    # ##  Code om beamset naam te checken                          ####
    # #################################################################
    msg7 = ""
    match = False
    for b in fplan.BeamSets:
        for c in b.DicomPlanLabel:
            if c in string.punctuation:
                match = True
                break
        if match == True:
            msg7 += '-' + 'Er is een leesteken in de beamset naam: %s\r\n' % b.DicomPlanLabel

    match2 = False
    for c in fplan.Name:
        if c in string.punctuation:
            match2 = True
            break
    if match2 == True:
        msg7 += '-' + 'Er is een leesteken in de plan naam: %s\r\n' % fplan.Name

    try:
        PlanNaam = fplan.Name.split()
        if not PlanNaam[0].isdigit():
            msg7 += '-' + 'Plan: ' + '<' + fplan.Name + '>' \
                    + ' :begint niet met een getal.\r\n'
    except:
        msg7 += '-' + 'Plan naam: ' + '<' + fplan.Name + '>' \
                + ' :begint niet met een getal.\r\n'
    try:
        for bs in fplan.BeamSets:
            BSNaam = bs.DicomPlanLabel.split()
            if not BSNaam[0].isdigit():
                msg7 += '-' + 'Beamset: ' + '<' + bs.DicomPlanLabel + '>' \
                        + ' :begint niet met een getal.\r\n'
    except:
        msg7 += '-' + 'Beamset: ' + '<' + bs.DicomPlanLabel + '>' \
                + ' :begint niet met een getal.\r\n'

    print "7"
    return msg7


def test8(fplan):
    # #################################################################
    # ##  Code om plannaam te vergelijken met beam set naam        ####
    # #################################################################
    msg8 = ""
    for bs in fplan.BeamSets:
        ftechniek = gettech(bs)
        if not fplan.Name == bs.DicomPlanLabel and fplan.BeamSets.Count == 1:
            msg8 += '-' + 'Plannaam en beam set naam zijn ongelijk. \r\n'
        if ftechniek == '3DCRT' and fplan.BeamSets.Count == 1:
            try:
                if fplan.Name != bs.Prescription.DosePrescriptions[0].Description and fplan.BeamSets.Count != 2:
                    msg8 += '-' + 'Plannaam en dose prescription site zijn ongelijk \r\n'
            except:
                pass

    print "8"
    return msg8


def test9(fcase, fplan, fbset):
    # #################################################################
    # ##  Code om planned_by, body site en setup beams te checken  ####
    # #################################################################
    msg9 = ""
    if fplan.PlannedBy == '' or fplan.PlannedBy is None:
        msg9 += '-' + 'Planner info mist.\r\n'
    if fcase.BodySite == '' or fcase.BodySite is None:
        msg9 += '-' + 'Body Site info mist.\r\n'
    if fbset.PatientSetup.UseSetupBeams:
        msg9 += '-' + 'Setup beams staat nog aan.\r\n'
    print "9"

    return msg9


def test10(fplan, fbset):
    # ###################################################
    # ##  Code om rekengrid te checken                ###
    # ###################################################
    msg10 = ""
    RekenGrid = 0.3
    KleinGridList = [
        'srt',
        'oligo',
        'gbm',
        'glioom',
        'oligo',
        'hersenen',
        'hals',
        'hypofyse',
        'ependymoom',
        'meningeoom'
    ]
    for i in KleinGridList:
        if i in fplan.Name.lower():
            RekenGrid = 0.2
            break
    #print 'x: ', str(fbset.FractionDose.InDoseGrid.VoxelSize.x)
    #print 'y: ', str(fbset.FractionDose.InDoseGrid.VoxelSize.y)
    #print 'z: ', str(fbset.FractionDose.InDoseGrid.VoxelSize.z)
    if fbset.FractionDose.InDoseGrid.VoxelSize.x > RekenGrid \
            or fbset.FractionDose.InDoseGrid.VoxelSize.y > RekenGrid \
            or fbset.FractionDose.InDoseGrid.VoxelSize.z > RekenGrid:
        msg10 += '-' + 'Rekengrid is groter dan ' + str(RekenGrid) + '\r\n'

    print "10"
    return msg10


def test11(fss, fNumIso, fIsoc_Multi):
    # #################################################################
    # ##  Code om ref-isoc afstand te checken                      ####
    # #################################################################
    msg11 = ""
    Max_Afst = 10
    try:
        for i in range(0, fNumIso):
            Afst_x = abs(round(fss.LocalizationPoiGeometry.Point.x - fIsoc_Multi[i][0], 2))
            Afst_y = abs(round(fss.LocalizationPoiGeometry.Point.z - fIsoc_Multi[i][2], 2))
            Afst_z = abs(round(-fss.LocalizationPoiGeometry.Point.y + fIsoc_Multi[i][1], 2))
            if Afst_x > Max_Afst or Afst_y > Max_Afst or Afst_z > Max_Afst:
                msg11 += '-' + 'Isoc' + str(i + 1) + ': Ref-Isoc afstand is groter dan %d' % Max_Afst + ' cm\r\n'
    except:
        MessageBox.Show('Er is geen localization point gedefinieerd!')
        msg11 += '-' + 'Er is geen localization point gedefinieerd!\r\n'

    print "11"
    return msg11


def test12(fplan):
    # #################################################################
    # ##  Code om bundelnaam+description+MaxMU te controleren      ####
    # #################################################################
    msg12 = ""
    beam_names = []
    desc_names = []
    for bss in fplan.BeamSets:
        for beam in bss.Beams:
            beam_names.append(beam.Name)
            if beam.Description != "":
                desc_names.append(beam.Description)
    dup_names = [x for n, x in enumerate(beam_names) if x in beam_names[:n]]
    if len(dup_names) != 0:
        msg12 += '-' + 'Dubbele naam in bundel: %s\r\n' % ", ".join(dup_names)
    for bsss in fplan.BeamSets:
        try:
            BeamChar = bsss.Beams[0].Name[0]
        except:
            pass
        try:
            if int(letter_to_index(BeamChar.lower())) != int(bsss.DicomPlanLabel.partition(' ')[0]):
                msg12 += '-' + \
                         'BeamSet index en bundelnaam index komen niet overeen\r\n'
        except:
            msg12 += '-' + \
                     'BeamSet index en bundelnaam index komen niet overeen\r\n'
        try:
            PlanNummer = int(bsss.Beams[0].Name[1])
        except:
            pass
        try:
            BmNummer = int(bsss.Beams[0].Name[-2:])  # voor gewone velden
        except:
            BmNummer = 1
        regCompiled = re.compile('^[A-Z]\d\.\d{2}')
        for beam in sorted(bsss.Beams, key=attrgetter('Number')):
            CheckMatch = False
            if re.match(regCompiled, beam.Name):
                CheckMatch = True
            try:
                cond_list = (CheckMatch,
                             beam.Name[0] == BeamChar,
                             int(beam.Name[1]) == PlanNummer,
                             int(beam.Name[-2:]) == BmNummer)
                if not all(cond_list):
                    msg12 += '-' + 'Naamfout in bundel: %s\r\n' % beam.Name
            except:
                msg12 += '-' + 'Naamfout in bundel: %s\r\n' % beam.Name
            try:
                if not str(int(round(beam.GantryAngle, 1))) in beam.Description:
                    msg12 += '-' + 'Description fout in bundel: %s\r\n' % beam.Name
            except:
                msg12 += '-' + 'Description fout in bundel: %s\r\n' % beam.Name
            try:
                if not 'iso' in beam.Description and beam.CouchRotationAngle != 0:
                    msg12 += '-' + '"Iso" ontbreekt in description voor bundel: %s\r\n' % beam.Name
            except:
                msg12 += '-' + '"Iso" ontbreekt in description voor bundel: %s\r\n' % beam.Name
            if beam.BeamMU < 5:
                msg12 += '-' + 'Minder dan 5MU in bundel: %s\r\n' % beam.Name
            if beam.BeamMU > 1000 and gettech(bsss) != 'VMAT':
                msg12 += '-' + 'Meer dan 1000MU in bundel: %s\r\n' % beam.Name

            BmNummer += 1

    print "12"
    return msg12


def test13(fplan):
    # #################################################################
    # ##  code om elekta corners te controleren                    ####
    # #################################################################
    msg13 = ""
    ElektaLimietVlag = False

    limits = [20.0 for i in range(80)]
    limits[0] = limits[79] = 16.1
    limits[1] = limits[78] = 16.7
    limits[2] = limits[77] = 17.3
    limits[3] = limits[76] = 17.8
    limits[4] = limits[75] = 18.3
    limits[5] = limits[74] = 18.8
    limits[6] = limits[73] = 19.2
    limits[7] = limits[72] = 19.7

    for bs in fplan.BeamSets:
        for beam in bs.Beams:
            for [segment_index, segment] in enumerate(beam.Segments):
                leaf_positions = segment.LeafPositions
                for i in range(limits.Count):
                    if leaf_positions[0][i] < -limits[i]:
                        ElektaLimietVlag = True
                        msg13 += '-' + beam.Name + ", segm. nummer: " + str(segment_index + 1) + " leaf " + str(
                            i + 1) + ": Agility limiet\n"
                    if leaf_positions[1][i] > limits[i]:
                        ElektaLimietVlag = True
                        msg13 += '-' + beam.Name + ", segm. nummer: " + str(segment_index + 1) + " leaf " + str(
                            i + 1) + ": Agility limiet\n"
    if ElektaLimietVlag:
        msg13 += '-' + \
                 'Hoek leafs vallen buiten de Elekta Agility limieten \r\n'

    print "13"
    return msg13


def test14(opt, fplan, fbset):
    # ##############################################################################################
    # ##  code om vmat opt parameters + energie+final dose + gantry spacing te controleren      ####
    # ##############################################################################################
    msg14 = ""
    if opt.OptimizationParameters.DoseCalculation.ComputeIntermediateDose:
        msg14 += '-' + "Compute intermediate dose is aan \r\n"
    if opt.OptimizationParameters.Algorithm.OptimalityTolerance > 0.0001:
        msg14 += '-' + "Optimization tolerance > 10^-5 \r\n"
    if not opt.OptimizationParameters.DoseCalculation.ComputeFinalDose:
        msg14 += '-' + 'Final dose staat niet aan \r\n'
    if "srt" in fplan.Name.lower() and opt.OptimizationParameters.TreatmentSetupSettings[
        0].SegmentConversion.ArcConversionProperties.MaxLeafTravelDistancePerDegree != 0.2:
        msg14 += '-' + 'Leaf motion constraint is niet juist voor SRT\r\n'
    if 0.3 < opt.OptimizationParameters.TreatmentSetupSettings[
        0].SegmentConversion.ArcConversionProperties.MaxLeafTravelDistancePerDegree > 0.5:
        msg14 += '-' + 'Leaf motion constraint is niet juist\r\n'
    if not opt.OptimizationParameters.TreatmentSetupSettings[
        0].SegmentConversion.ArcConversionProperties.UseMaxLeafTravelDistancePerDegree:
        msg14 += '-' + 'Leaf motion constraint is niet ingesteld \r\n'
    if 10 in getenergiebset(fbset):
        msg14 += '-' + 'Verkeerde energie voor vmat \r\n'
    for ts in opt.OptimizationParameters.TreatmentSetupSettings:
        for beam in ts.BeamSettings:
            if not beam.ArcConversionPropertiesPerBeam.FinalArcGantrySpacing in [2, 3, 4]:
                msg14 += '-' + 'Beam: ' + '<' + beam.ForBeam.Name + '>' + ' GantrySpacing is niet 2, 3 of 4 \r\n'

    print "14"
    return msg14


def test15(fexam):
    # ####################################################
    # ##  Code om slice dikte te checken               ###
    # ####################################################
    msg15 = ""
    image_stack = fexam.Series[0].ImageStack
    if len(image_stack.SlicePositions) > 1:
        slice_dikte = round(abs(image_stack.SlicePositions[1] - image_stack.SlicePositions[0]), 2)
    if slice_dikte < 0.1:
        msg15 += '-' + 'Slice dikte is kleiner dan 1mm \r\n'
    elif slice_dikte > 0.3:
        msg15 += '-' + 'Slice dikte is groter dan 3mm \r\n'

    print "15"
    return msg15


def test16(fcase):
    # ##############################################################################
    # ### Code om te checken of een CT serie bestaat met evenveel slices als PET ###
    # ##############################################################################
    msg16 = ""
    petnaam = ""
    matchingCTFlag = False
    for ex in fcase.Examinations:
        if 'PET' in ex.Name:
            petnaam = ex.Name
            no_of_slicesPET = len(ex.Series[0].ImageStack.SlicePositions)
            for exa in fcase.Examinations:
                if 'CT' in exa.Name:
                    no_of_slicesCT = len(exa.Series[0].ImageStack.SlicePositions)
                    if no_of_slicesPET - 1 <= no_of_slicesCT <= no_of_slicesPET + 1:
                        matchingCTFlag = True

    if not matchingCTFlag and petnaam != "":
        msg16 += '-' + 'Pet serie: ' + petnaam + ' heeft geen CT serie met evenveel slices. Klopt dat?' + '\r\n'

    print "16"
    return msg16


def test17(fplan, fcase, fbset):
    # ###################################################
    # ### Code om de plek van dmax te checken         ###
    # ###################################################
    msg17 = ""
    Dmaxbericht = False
    if gettech(fbset) != '3DCRT' and gettech(fbset) != 'Elektronen':
        Dmaxbericht = is_max_dosis_in_target(fplan, fcase)
    if Dmaxbericht:
        msg17 += '-' + 'Maximale dosis valt buiten een TargetRoi.\r\n'

    print "17"
    return msg17


def test19(fbset, fplan, fcase):
    # ##################################################################
    # # verkrijg de opgetelde dosis in dsp-punt en globale Dmax        #
    # ##################################################################
    # global DSPtotaalDosis
    msgDose = ""
    try:
        MaxDoseValue = fplan.TreatmentCourse.TotalDose.GetDoseStatistic(RoiName='External', DoseType='Max')
        if fcase.CaseSettings.DoseColorMap.ColorMapReferenceType == 'RelativePrescription':
            PresDoseValue = fbset.Prescription.PrimaryDosePrescription.DoseValue
        else:
            PresDoseValue = fcase.CaseSettings.DoseColorMap.ReferenceValue
        maxperc = round((MaxDoseValue / PresDoseValue), 4) * 100
        PresGy = round(PresDoseValue / 100, 2)
        for bs in fplan.BeamSets:
            DSPtotaalDosis = 0.000
            for i in range(bs.Beams.Count):
                beam_nr = bs.Beams[i].Number
                temp = 0
                dsp_nr = bs.FractionDose.BeamDoses[temp].ForBeam.Number
                while beam_nr != dsp_nr:
                    temp += 1
                    dsp_nr = bs.FractionDose.BeamDoses[temp].ForBeam.Number
                dsp_dx = temp
                DSPtotaalDosis += round(0.01 * bs.FractionDose.BeamDoses[dsp_dx].DoseAtPoint.DoseValue, 2)
            msgDose += '<' + bs.DicomPlanLabel + '>' + ': Totale dsp-dosis: %s ' % DSPtotaalDosis + 'Gy \r\n'
        msgDose += 'Globale Dmax: %s' % maxperc + '%, ' + ' (100%% = %sGy)' % PresGy + '\r\n'
    except:
        msgDose += 'Er is geen dosis berekend voor dsp-punt \r\n'
    print "19"
    return msgDose


def test20(opt):
    # #######################################################
    # ##  code om imrt opt parameters te controleren     ####
    # #######################################################
    msg20 = ""
    if not opt.OptimizationParameters.DoseCalculation.ComputeFinalDose:
        msg20 += '-' + 'Final dose staat niet aan \r\n'
    if not opt.OptimizationParameters.DoseCalculation.ComputeIntermediateDose:
        msg20 += '-' + 'Intermediate dose staat niet aan \r\n'

    print "20"
    return msg20


def test21(fbset):
    # ###################################################
    # ### Code om elektronen instellingen te checken  ###
    # ###################################################
    msg21 = ""
    if fbset.FractionDose.InDoseGrid.VoxelSize.x > 0.2 \
            or fbset.FractionDose.InDoseGrid.VoxelSize.y > 0.2 \
            or fbset.FractionDose.InDoseGrid.VoxelSize.z > 0.2:
        msg21 += '-' + 'Rekengrid is groter dan 0.2cm (voor elektronen) \r\n'
    if fbset.AccurateDoseAlgorithm.MonteCarloHistoriesPerAreaFluence < 100000:
        msg21 += '-' + 'Aantal histories staat niet goed (voor elektronen) \r\n'

    print "21"
    return msg21


def test22(fplan):
    # ###################################################
    # ### Code om DSP punt koppeling te controleren   ###
    # ###################################################
    msg22 = ""
    for bs in fplan.BeamSets:
        if bs.DoseSpecificationPoints.Count == 0:
            msg22 += '<' + bs.DicomPlanLabel + '>' + ': Er is geen DSP punt aangemaakt \r\n'
    for bs in fplan.BeamSets:
        for di in range(bs.Beams.Count):
            beam = bs.Beams[di]
            if bs.FractionDose.ForBeamSet.FractionDose.ForBeamSet.FractionDose.BeamDoses[
                di].UserSetBeamDoseSpecificationPoint is None:
                msg22 += '-' + 'DSP punt is niet gekoppeld aan bundel: %s\r\n' % beam.Name

    print "22"
    return msg22


def test23(fbset, fplan):
    # ################################################
    # ### Code om op merged beams te controleren   ###
    # ################################################
    msg23 = ""
    MammaList = ["mamma",
                 "mam",
                 "ma",
                 "mst",
                 "thw",
                 "thoraxw",
                 "wand",
                 "th",
                 "thoraxwand",
                 "thwand"
                 ]
    member_of_mammalist = False
    gantry_and_coll_angles = set([])
    has_unmerged_beams = None
    for i in MammaList:
        if i in fplan.Name.lower():
            member_of_mammalist = True
            break
    if not member_of_mammalist:
        for beam in fbset.Beams:
            beam_angles = str(beam.GantryAngle) + '-' + str(beam.InitialCollimatorAngle)
            if beam_angles in gantry_and_coll_angles:
                has_unmerged_beams = True
            else:
                gantry_and_coll_angles.add(beam_angles)
    if has_unmerged_beams:
        msg23 += '-' + '<' + fbset.DicomPlanLabel + '>' + 'Bundels kunnen gemerged worden \r\n'

    print "23"
    return msg23


def test24(fplan, fss, fNumIso, fIsoc_Multi):
    # #################################################################
    # ##  Code om tafelblad-isoc afstand te checken                ####
    # #################################################################
    msg24 = ""
    tafelnaam = ""
    Max_Afst = 47
    TafelRoi = {'Tafel'}
    for current in fss.RoiGeometries:
        if current.OfRoi.Name in TafelRoi and current.PrimaryShape:
            tafelnaam = current.OfRoi.Name
    if tafelnaam != "":
        BBTafel = fss.RoiGeometries[tafelnaam].GetBoundingBox()

    try:
        for bs in fplan.BeamSets:
            no_of_beams = bs.Beams.Count
            tiso = False
            for dx in range(no_of_beams):
                if bs.Beams[dx].CouchRotationAngle != 0.0:
                    tiso = True
        if tiso:
            msg24 += '-' + 'Tafel-iso rotatie gevonden: Kan isoc-tafelrand afstand niet controleren.\r\n'
        else:
            for i in range(0, fNumIso):
                # print fIsoc_Multi[i][0]
                # print  fIsoc_Multi[i][1]
                # print "BBTafel[0].x - Isoc_Multi[i][0]", BBTafel[0].x - fIsoc_Multi[i][0]
                # print "BBTafel[0].y - Isoc_Multi[i][1])", BBTafel[0].y - (fIsoc_Multi[i][1])
                # print "BBTafel[1].x - Isoc_Multi[i][0]", BBTafel[1].x - fIsoc_Multi[i][0]
                # print "BBTafel[0].y - (Isoc_Multi[i][1])", BBTafel[0].y - (fIsoc_Multi[i][1])
                dist_1 = math.hypot(BBTafel[0].x - fIsoc_Multi[i][0], (BBTafel[0].y - (fIsoc_Multi[i][1])) + 1)
                dist_2 = math.hypot(BBTafel[1].x - fIsoc_Multi[i][0], (BBTafel[0].y - (fIsoc_Multi[i][1])) + 1)
                print "dist_1, dist_2", dist_1, dist_2
                if max(dist_1, dist_2) > Max_Afst:
                    msg24 += '-' + 'Isoc' + str(
                        i + 1) + ': Isoc-tafelrand afstand is groter dan  %d' % Max_Afst + ' cm\r\n'
    except:
        pass

    print "24"
    return msg24


def test25(fplan):
    # #############################################################
    # ##  Code om SIB in de plan naam te vinden                ####
    # #############################################################
    msg25 = ""
    NO_SIBList = ["prostaat",
                  "srt",
                  "mam"
                  ]
    member_of_NO_SIBList = False
    for i in NO_SIBList:
        if i in fplan.Name.lower():
            member_of_NO_SIBList = True
            break
    if not member_of_NO_SIBList:
        numEval = fplan.TreatmentCourse.EvaluationSetup.EvaluationFunctions.Count
        PTVnaam = set()
        PTVdose = set()
        numGevonden = 0
        for i in range(numEval):
            goal = fplan.TreatmentCourse.EvaluationSetup.EvaluationFunctions[i]
            goalname = goal.ForRegionOfInterest.Name
            tyPe = goal.PlanningGoal.Type
            crit = goal.PlanningGoal.GoalCriteria
            param = goal.PlanningGoal.ParameterValue
            if tyPe == "VolumeAtDose" and crit == "AtLeast":
                numGevonden += 1
                PTVnaam.add(goalname)
                PTVdose.add(param)
        if numGevonden > 1 and "sib" not in fplan.Name.lower() and PTVnaam.Count > 1 and PTVdose.Count > 1:
            msg25 += '-' + 'SIB bevindt zich niet in plannaam \r\n'

    print "25"
    return msg25


def test26(fcase):
    # ############################################################
    # ##  Code om overbodige plannen te vinden                ####
    # ############################################################
    msg26 = ""
    for pl in fcase.TreatmentPlans:
        if 'kanweg' in pl.Name.lower():
            msg26 += '-' + 'Plan: kanweg kan verwijderd worden\r\n'
        elif 'los_' in pl.Name.lower():
            msg26 += '-' + 'Plan: ' + pl.Name + ' kan verwijderd worden\r\n'
        elif '_kopie' in pl.Name.lower():
            msg26 += '-' + 'Plan: ' + pl.Name + ' kan verwijderd worden\r\n'

    print "26"
    return msg26


def test27(fplan, fcase, fexam):
    # ###############################################################################
    # ##  Code om controleren of een DSPpunt in lucht of bot bevindt             ####
    # ###############################################################################
    msg27 = ''
    if fplan.Review is None or fplan.Review.ApprovalStatus != "Approved":
        for bs in fplan.BeamSets:
            for dsp in bs.DoseSpecificationPoints:
                try:
                    dsp_roi_name = dsp.Name + '_temp'
                    dsp_geom = fcase.PatientModel.CreateRoi(Name=dsp_roi_name, Color='Brown', Type='Control')
                    dsp_geom.CreateSphereGeometry(Radius=0.2,
                                                  Examination=fexam,
                                                  Center={"x": dsp.Coordinates.x,
                                                          "y": dsp.Coordinates.y,
                                                          "z": dsp.Coordinates.z})
                    fplan.TreatmentCourse.TotalDose.UpdateDoseGridStructures()
                    AverageDensity = fGetAverageDensityValues(dsp_roi_name, fplan)
                    print 'AverageDensity: ' , AverageDensity
                    if 0 < AverageDensity < 0.25:
                        msg27 += '-' + 'DSPpunt: ' + dsp.Name + ' is in lucht' + '\r\n'
                    elif 1.2 < AverageDensity:
                        msg27 += '-' + 'DSPpunt: ' + dsp.Name + ' is in bot' + '\r\n'
                    dsp_geom.DeleteRoi()
                except:
                    pass

    print "27"
    return msg27


def MainLoop():
    from System.IO import File
    import logging
    if File.Exists(r'\\skf-rif.nl\rif\Raystation\Logfiles\plancheck.log'):
        pass
    else:
        sys.exit()
    logging.basicConfig(filename=r'\\skf-rif.nl\rif\Raystation\Logfiles\plancheck.log', level=logging.INFO,
                        format='%(asctime)s %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p')
    global bset, case, plan, curmsg, patient, FFF, VsimVlag, VSim_Beam_Set, exam, ss
    try:
        case = get_current('Case')
        patient = get_current('Patient')
        plan = get_current('Plan')
        bset = get_current('BeamSet')
        exam = get_current('Examination')
        ss = plan.GetStructureSet()
        machine_db = get_current('MachineDB')
        po = FindPlanOpt(plan, bset)
    except:
        MessageBox.Show('Een patient en plan moet geladen zijn.')
        sys.exit()
    #################################################
    ###  Code om te controlereen of er dosis is  ####
    #################################################
    dvlag = True
    if not VsimVlag:
        for bs in plan.BeamSets:
            if bs.FractionDose.DoseValues is None:
                dvlag = False
                break
        if not dvlag:
            MessageBox.Show('Geen dosis berekend voor beamset: {0}'.format(bs.DicomPlanLabel))

    # Haal veelgebruikte globals
    NumPOI = case.PatientModel.PointsOfInterest.Count
    machine = machine_db.GetTreatmentMachine(machineName=bset.MachineReference.MachineName, lockMode=None)
    FFF = machine.PhotonBeamQualities[0].FluenceMode
    scan_naam = bset.PatientSetup.LocalizationPoiGeometrySource.OnExamination.Name

    curmsg = ''
    if not VsimVlag:
        curmsg += test0(case)
        curmsg += test1(patient, plan, ss)
        curmsg += test2(ss, case, plan)
        curmsg += test3(scan_naam, ss)
        curmsg += test4(plan)
        curmsg += test5(NumPOI, case)
        curmsg += test7(plan)
        curmsg += test8(plan)
        curmsg += test10(plan, bset)
        curmsg += test11(ss, NumIso, Isoc_Multi)
        curmsg += test12(plan)
        curmsg += test15(exam)
        #curmsg += test16(case)
        curmsg += test17(plan, case, bset)
        curmsg += test22(plan)
        curmsg += test24(plan, ss, NumIso, Isoc_Multi)
        curmsg += test25(plan)
        curmsg += test26(case)
        for bs in plan.BeamSets:
            curmsg += test9(case, plan, bs)
            if gettech(bs) == 'VMAT':
                curmsg += test14(FindPlanOpt(plan, bs), plan, bs)
                curmsg += test6(bs)
            if gettech(bs) == 'IMRT':
                curmsg += test20(FindPlanOpt(plan, bs))
            if gettech(bs) == 'IMRT' or gettech(bs) == '3DCRT':
                curmsg += test23(bs, plan)
            if gettech(bs) == 'Elektronen':
                curmsg += test21(bs)
            if gettech(bs) != 'Elektronen':
                curmsg += test13(plan)
            #if gettech(bs) == 'VMAT' or gettech(bs) == 'IMRT':
                #curmsg += test27(plan, case, exam)
        patient.Save()
    else:
        curmsg += CheckVsimPlan(bset, VSim_Beam_Set)
        VsimVlag = False
    if LOG_AAN == True:
        logging.info(curmsg)
    return curmsg


from connect import *

if __name__ == '__main__':
    VsimVlag = False
    Application().Run(MyWindow())

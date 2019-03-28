'# MWS Version: Version 2015.0 - Jan 16 2015 - ACIS 24.0.2 -

'# length = mm
'# frequency = GHz
'# time = ns
'# frequency range: fmin = 50 fmax = 70
'# created = '[VERSION]2015.0|24.0.2|20150116[/VERSION]


'@ use template: Antenna - Waveguide_1

'[VERSION]2015.0|24.0.2|20150116[/VERSION]
'set the units
With Units
    .Geometry "mm"
    .Frequency "GHz"
    .Voltage "V"
    .Resistance "Ohm"
    .Inductance "NanoH"
    .TemperatureUnit  "Kelvin"
    .Time "ns"
    .Current "A"
    .Conductance "Siemens"
    .Capacitance "PikoF"
End With
'----------------------------------------------------------------------------
Plot.DrawBox True
With Background
     .Type "Normal"
     .Epsilon "1.0"
     .Mue "1.0"
     .XminSpace "0.0"
     .XmaxSpace "0.0"
     .YminSpace "0.0"
     .YmaxSpace "0.0"
     .ZminSpace "0.0"
     .ZmaxSpace "0.0"
End With
With Boundary
     .Xmin "expanded open"
     .Xmax "expanded open"
     .Ymin "expanded open"
     .Ymax "expanded open"
     .Zmin "expanded open"
     .Zmax "expanded open"
     .Xsymmetry "none"
     .Ysymmetry "none"
     .Zsymmetry "none"
End With
' switch on FD-TET setting for accurate farfields
FDSolver.ExtrudeOpenBC "True"
Mesh.FPBAAvoidNonRegUnite "True"
Mesh.ConsiderSpaceForLowerMeshLimit "False"
Mesh.MinimumStepNumber "5"
With MeshSettings
     .SetMeshType "Hex"
     .Set "RatioLimitGeometry", "20"
End With
With MeshSettings
     .SetMeshType "HexTLM"
     .Set "RatioLimitGeometry", "20"
End With
PostProcess1D.ActivateOperation "vswr", "true"
PostProcess1D.ActivateOperation "yz-matrices", "true"
'----------------------------------------------------------------------------
'set the frequency range
Solver.FrequencyRange "50", "70"
Dim sDefineAt As String
sDefineAt = "50;60;70"
Dim sDefineAtName As String
sDefineAtName = "50;60;70"
Dim sDefineAtToken As String
sDefineAtToken = "f="
Dim aFreq() As String
aFreq = Split(sDefineAt, ";")
Dim aNames() As String
aNames = Split(sDefineAtName, ";")
Dim nIndex As Integer
For nIndex = LBound(aFreq) To UBound(aFreq)
Dim zz_val As String
zz_val = aFreq (nIndex)
Dim zz_name As String
zz_name = sDefineAtToken & aNames (nIndex)
' Define E-Field Monitors
With Monitor
    .Reset
    .Name "e-field ("& zz_name &")"
    .Dimension "Volume"
    .Domain "Frequency"
    .FieldType "Efield"
    .Frequency zz_val
    .Create
End With
' Define H-Field Monitors
With Monitor
    .Reset
    .Name "h-field ("& zz_name &")"
    .Dimension "Volume"
    .Domain "Frequency"
    .FieldType "Hfield"
    .Frequency zz_val
    .Create
End With
' Define Power flow Monitors
With Monitor
    .Reset
    .Name "power ("& zz_name &")"
    .Dimension "Volume"
    .Domain "Frequency"
    .FieldType "Powerflow"
    .Frequency zz_val
    .Create
End With
' Define Power loss Monitors
With Monitor
    .Reset
    .Name "loss ("& zz_name &")"
    .Dimension "Volume"
    .Domain "Frequency"
    .FieldType "Powerloss"
    .Frequency zz_val
    .Create
End With
' Define Farfield Monitors
With Monitor
    .Reset
    .Name "farfield ("& zz_name &")"
    .Domain "Frequency"
    .FieldType "Farfield"
    .Frequency zz_val
    .ExportFarfieldSource "False"
    .Create
End With
Next
'----------------------------------------------------------------------------
With MeshSettings
     .SetMeshType "Hex"
     .Set "Version", 1%
End With
With Mesh
     .MeshType "PBA"
End With
'set the solver type
ChangeSolverType("HF Time Domain")

'@ define material: notAR

'[VERSION]2015.0|24.0.2|20150116[/VERSION]
With Material 
     .Reset 
     .Name "notAR"
     .Folder ""
     .FrqType "all"
     .Type "Normal"
     .MaterialUnit "Frequency", "GHz"
     .MaterialUnit "Geometry", "mm"
     .MaterialUnit "Time", "ns"
     .MaterialUnit "Temperature", "Kelvin"
     .Epsilon "2.9"
     .Mue "1"
     .Sigma "0"
     .TanD "0.01"
     .TanDFreq "60"
     .TanDGiven "True"
     .TanDModel "ConstSigma"
     .EnableUserConstTanDModelOrderEps "False"
     .ConstTanDModelOrderEps "1"
     .SetElParametricConductivity "False"
     .ReferenceCoordSystem "Global"
     .CoordSystemType "Cartesian"
     .SigmaM "0"
     .TanDM "0.01"
     .TanDMFreq "60"
     .TanDMGiven "False"
     .TanDMModel "ConstTanD"
     .EnableUserConstTanDModelOrderMue "False"
     .ConstTanDModelOrderMue "1"
     .SetMagParametricConductivity "False"
     .DispModelEps "None"
     .DispModelMue "None"
     .DispersiveFittingSchemeEps "Nth Order"
     .MaximalOrderNthModelFitEps "10"
     .ErrorLimitNthModelFitEps "0.1"
     .UseOnlyDataInSimFreqRangeNthModelEps "False"
     .DispersiveFittingSchemeMue "Nth Order"
     .MaximalOrderNthModelFitMue "10"
     .ErrorLimitNthModelFitMue "0.1"
     .UseOnlyDataInSimFreqRangeNthModelMue "False"
     .UseGeneralDispersionEps "False"
     .UseGeneralDispersionMue "False"
     .NLAnisotropy "False"
     .NLAStackingFactor "1"
     .NLADirectionX "1"
     .NLADirectionY "0"
     .NLADirectionZ "0"
     .Rho "0"
     .ThermalType "Normal"
     .ThermalConductivity "0"
     .HeatCapacity "0"
     .MetabolicRate "0"
     .BloodFlow "0"
     .VoxelConvection "0"
     .MechanicsType "Unused"
     .Colour "0", "1", "1" 
     .Wireframe "False" 
     .Reflection "False" 
     .Allowoutline "True" 
     .Transparentoutline "False" 
     .Transparency "0" 
     .Create
End With

'@ define material colour: notAR

'[VERSION]2015.0|24.0.2|20150116[/VERSION]
With Material 
     .Name "notAR"
     .Folder ""
     .Colour "0", "1", "1" 
     .Wireframe "False" 
     .Reflection "False" 
     .Allowoutline "True" 
     .Transparentoutline "False" 
     .Transparency "0" 
     .ChangeColour 
End With

'@ new component: notAR

'[VERSION]2015.0|24.0.2|20150116[/VERSION]
Component.New "notAR"

'@ define brick: notAR:notAR

'[VERSION]2015.0|24.0.2|20150116[/VERSION]
With Brick
     .Reset 
     .Name "notAR" 
     .Component "notAR" 
     .Material "notAR" 
     .Xrange "0", "2.5" 
     .Yrange "0", "2.5" 
     .Zrange "0", "length" 
     .Create
End With


'@ define frequency range

'[VERSION]2015.0|24.0.2|20150116[/VERSION]
Solver.FrequencyRange "50", "70"

'@ pick face

'[VERSION]2015.0|24.0.2|20150116[/VERSION]
Pick.PickFaceFromId "notAR:notAR", "2"

'@ define port: 1

'[VERSION]2015.0|24.0.2|20150116[/VERSION]
With Port 
     .Reset 
     .PortNumber "1" 
     .Label "" 
     .NumberOfModes "1" 
     .AdjustPolarization "False" 
     .PolarizationAngle "0.0" 
     .ReferencePlaneDistance "0" 
     .TextSize "50" 
     .Coordinates "Picks" 
     .Orientation "positive" 
     .PortOnBound "False" 
     .ClipPickedPortToBound "False" 
     .Xrange "0", "2.5" 
     .Yrange "0", "2.5" 
     .Zrange "0", "0" 
     .XrangeAdd "0.0", "0.0" 
     .YrangeAdd "0.0", "0.0" 
     .ZrangeAdd "0.0", "0.0" 
     .SingleEnded "False" 
     .Create 
End With

'@ pick face

'[VERSION]2015.0|24.0.2|20150116[/VERSION]
Pick.PickFaceFromId "notAR:notAR", "1"

'@ define port: 2

'[VERSION]2015.0|24.0.2|20150116[/VERSION]
With Port 
     .Reset 
     .PortNumber "2" 
     .Label "" 
     .NumberOfModes "1" 
     .AdjustPolarization "False" 
     .PolarizationAngle "0.0" 
     .ReferencePlaneDistance "0" 
     .TextSize "50" 
     .Coordinates "Picks" 
     .Orientation "positive" 
     .PortOnBound "False" 
     .ClipPickedPortToBound "False" 
     .Xrange "0", "2.5" 
     .Yrange "0", "2.5" 
     .Zrange "10", "10" 
     .XrangeAdd "0.0", "0.0" 
     .YrangeAdd "0.0", "0.0" 
     .ZrangeAdd "0.0", "0.0" 
     .SingleEnded "False" 
     .Create 
End With

'@ set parametersweep options

'[VERSION]2015.0|24.0.2|20150116[/VERSION]
With ParameterSweep
    .SetSimulationType "Transient" 
End With

'@ add parsweep sequence: Sequence 1

'[VERSION]2015.0|24.0.2|20150116[/VERSION]
With ParameterSweep
     .AddSequence "Sequence 1" 
End With

'@ add parsweep parameter: Sequence 1:length

'[VERSION]2015.0|24.0.2|20150116[/VERSION]
With ParameterSweep
     .AddParameter_Stepwidth "Sequence 1", "length", "1", "13", "1" 
End With

'@ define time domain solver parameters

'[VERSION]2015.0|24.0.2|20150116[/VERSION]
Mesh.SetCreator "High Frequency" 
With Solver 
     .Method "Hexahedral"
     .CalculationType "TD-S"
     .StimulationPort "All"
     .StimulationMode "All"
     .SteadyStateLimit "-30.0"
     .MeshAdaption "False"
     .AutoNormImpedance "False"
     .NormingImpedance "50"
     .CalculateModesOnly "False"
     .SParaSymmetry "False"
     .StoreTDResultsInCache  "False"
     .FullDeembedding "False"
     .SuperimposePLWExcitation "False"
     .UseSensitivityAnalysis "False"
End With

'@ delete port: port1

'[VERSION]2015.0|24.0.2|20150116[/VERSION]
Port.Delete "1" 


'@ delete port: port2

'[VERSION]2015.0|24.0.2|20150116[/VERSION]
Port.Delete "2" 


'@ define boundaries

'[VERSION]2015.0|24.0.2|20150116[/VERSION]
With Boundary
     .Xmin "electric" 
     .Xmax "electric" 
     .Ymin "magnetic" 
     .Ymax "magnetic" 
     .Zmin "expanded open" 
     .Zmax "expanded open" 
     .Xsymmetry "none" 
     .Ysymmetry "none" 
     .Zsymmetry "none" 
     .XminThermal "isothermal" 
     .XmaxThermal "isothermal" 
     .YminThermal "isothermal" 
     .YmaxThermal "isothermal" 
     .ZminThermal "isothermal" 
     .ZmaxThermal "isothermal" 
     .XsymmetryThermal "none" 
     .YsymmetryThermal "none" 
     .ZsymmetryThermal "none" 
     .ApplyInAllDirections "False" 
     .ApplyInAllDirectionsThermal "False" 
     .XminTemperature "" 
     .XminTemperatureType "None" 
     .XmaxTemperature "" 
     .XmaxTemperatureType "None" 
     .YminTemperature "" 
     .YminTemperatureType "None" 
     .YmaxTemperature "" 
     .YmaxTemperatureType "None" 
     .ZminTemperature "" 
     .ZminTemperatureType "None" 
     .ZmaxTemperature "" 
     .ZmaxTemperatureType "None" 
End With


'@ define port: 2

'[VERSION]2015.0|24.0.2|20150116[/VERSION]
With Port 
     .Reset 
     .PortNumber "2" 
     .Label "" 
     .NumberOfModes "1" 
     .AdjustPolarization "False" 
     .PolarizationAngle "0.0" 
     .ReferencePlaneDistance "0" 
     .TextSize "50" 
     .Coordinates "Full" 
     .Orientation "zmin" 
     .PortOnBound "True" 
     .ClipPickedPortToBound "False" 
     .Xrange "0", "2.5" 
     .Yrange "0", "2.5" 
     .Zrange "-1.2491352416667", "-1.2491352416667" 
     .XrangeAdd "0.0", "0.0" 
     .YrangeAdd "0.0", "0.0" 
     .ZrangeAdd "0.0", "0.0" 
     .SingleEnded "False" 
     .Create 
End With 


'@ define port: 1

'[VERSION]2015.0|24.0.2|20150116[/VERSION]
With Port 
     .Reset 
     .PortNumber "1" 
     .Label "" 
     .NumberOfModes "1" 
     .AdjustPolarization "False" 
     .PolarizationAngle "0.0" 
     .ReferencePlaneDistance "0" 
     .TextSize "50" 
     .Coordinates "Full" 
     .Orientation "zmax" 
     .PortOnBound "True" 
     .ClipPickedPortToBound "False" 
     .Xrange "0", "2.5" 
     .Yrange "0", "2.5" 
     .Zrange "11.249135241667", "11.249135241667" 
     .XrangeAdd "0.0", "0.0" 
     .YrangeAdd "0.0", "0.0" 
     .ZrangeAdd "0.0", "0.0" 
     .SingleEnded "False" 
     .Create 
End With 


'@ edit parsweep parameter: Sequence 1:length

'[VERSION]2015.0|24.0.2|20150116[/VERSION]
With ParameterSweep
     .DeleteParameter "Sequence 1", "length" 
     .AddParameter_Stepwidth "Sequence 1", "length", "2", "12", "1" 
End With


'@ delete parsweep parameter: Sequence 1:length

'[VERSION]2015.0|24.0.2|20150116[/VERSION]
With ParameterSweep
     .DeleteParameter "Sequence 1", "length" 
End With


'@ add parsweep parameter: Sequence 1:lengths

'[VERSION]2015.0|24.0.2|20150116[/VERSION]
With ParameterSweep
     .AddParameter_Stepwidth "Sequence 1", "lengths", "1", "13", "1" 
End With


'@ add parsweep parameter: Sequence 1:length

'[VERSION]2015.0|24.0.2|20150116[/VERSION]
With ParameterSweep
     .AddParameter_Samples "Sequence 1", "length", "13", "13", "1", "False" 
End With


'@ delete parsweep parameter: Sequence 1:length

'[VERSION]2015.0|24.0.2|20150116[/VERSION]
With ParameterSweep
     .DeleteParameter "Sequence 1", "length" 
End With


'@ add parsweep sequence: Sequence 2

'[VERSION]2015.0|24.0.2|20150116[/VERSION]
With ParameterSweep
     .AddSequence "Sequence 2" 
End With


'@ add parsweep parameter: Sequence 2:lengths

'[VERSION]2015.0|24.0.2|20150116[/VERSION]
With ParameterSweep
     .AddParameter_Stepwidth "Sequence 2", "lengths", "2", "13", "1" 
End With



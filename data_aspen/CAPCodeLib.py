from fileinput import filename
import os
from re import A
from tokenize import String
from typing import Union, Dict, Literal
import win32com.client as win32
import numpy as np
import time
#from scripy import optimize

class Simulation():
    """Class which starts a Simulation interface instance

    Args:
        AspenFileName: Name of the Aspenfile on which you are working with
        WorkingDirectoryPath: Path to the Folder where we will be working
        VISIBITLITY: Toggles the opening and interactive running of the Aspen simulation
    """

    def __init__(self, AspenFileName: str, WorkingDirectoryPath: str, VISIBILITY: bool = True):
        import os
        import pythoncom
        import win32com.client as win32

        pythoncom.CoInitialize()  # 노트북/스레드에서 COM 안정성에 도움

        print("The current Directory is :")
        print(os.getcwd())

        working_directory = os.path.abspath(WorkingDirectoryPath)
        os.chdir(working_directory)
        print("The new Directory where you should also have your Aspen file is :")
        print(os.getcwd())

        if os.path.isabs(AspenFileName):
            abs_path = AspenFileName
        else:
            abs_path = os.path.join(working_directory, AspenFileName)
        abs_path = os.path.abspath(abs_path)

        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"Aspen file not found: {abs_path}")

        self.working_directory = working_directory
        self.aspen_filename = os.path.basename(abs_path)
        self.abs_path = abs_path
        self.last_run_error = None
        self.last_per_error = None

        self.AspenSimulation = win32.DispatchEx("Apwn.Document")
        self.AspenSimulation.InitFromFile(self.abs_path)
        self.AspenSimulation.Visible = VISIBILITY
        print(f"Loaded by InitFromFile: {self.abs_path}")

    def CloseAspen(self):
        AspenFileName = self.Give_AspenDocumentName()
        print(AspenFileName)
        self.AspenSimulation.Close(os.path.abspath(AspenFileName))
        print("\nAspen should be closed now")

    # This just shortens the path you need to call for Streams and Blocks

    @property
    def BLK(self):
        """Property: Defines Path to the Block node in Aspen File system.

        Aspendocument is defined in the Class Simulation initialization
        """
        return self.AspenSimulation.Tree.Elements("Data").Elements("Blocks")

    @property
    def STRM(self):
        """Property: Defines Path to the Streamnode node in Aspen File system.

        Aspendocument is defined in the Class Simulation initialization
        """
        return self.AspenSimulation.Tree.Elements("Data").Elements("Streams")

    # Type definition to simplify the type hinting:
    Phnum = Literal[1, 2, 3]
    Ph = Literal["L", "V", "S"]

    def Give_AspenDocumentName(self) -> String:
        """Returns name of Aspen document"""
        return self.AspenSimulation.FullName















#STREAM OUTPUTS
    def STRM_Get_TotalMassFlow(self, Streamname: str) -> float:
        return self.STRM.Elements(Streamname).Elements("Output").Elements("RES_MASSFLOW").Value

    def STRM_Get_TotalMoleFlow(self, Streamname: str) -> float:
        return self.STRM.Elements(Streamname).Elements("Output").Elements("RES_MOLEFLOW").Value














#RPLUG INPUTS
    def BLK_RPLUG_GET_ME_ALL_INPUTS_BACK(self, Blockname: str) -> Dict[str, Union[str, float, int]]:
        """Retrieves all the Inputs and returns Dictionary with Values

        Does not include all aspects of a Aspen Simulationsheet, for this look at Exports

        Args:
            Blockname: String which gives the name of Block.
        """

        TYPE = self.BLK.Elements(Blockname).Elements("Input").Elements("TYPE").Value
        Operating_conditions = self.BLK.Elements(Blockname).Elements("Input").Elements(
            "OPT_TSPEC").Value  # Chose between INLET-TEMP, CONST-TEMP, TEMP-PROF
        ReactorTemperature = self.BLK.Elements(Blockname).Elements("Input").Elements("REAC_TEMP").Value

        Constant_Temp = self.BLK.Elements(Blockname).Elements("Input").Elements("CTEMP").Value
        OutletTemp = self.BLK.Elements(Blockname).Elements("Input").Elements("TEMP").Value
        U = self.BLK.Elements(Blockname).Elements("Input").Elements("U").Value
        Activate_YES_NO = self.BLK.Elements(Blockname).Elements("Input").Elements("CHK_NTUBE").Value
        Number_of_Tubes = self.BLK.Elements(Blockname).Elements("Input").Elements("NTUBE").Value
        TubeLength = self.BLK.Elements(Blockname).Elements("Input").Elements("LENGTH").Value
        TubeDiameter = self.BLK.Elements(Blockname).Elements("Input").Elements("DIAM").Value
        Phase = self.BLK.Elements(Blockname).Elements("Input").Elements("PHASE").Value  # This can be V L or S
        Phasenumber = self.BLK.Elements(Blockname).Elements("Input").Elements("NPHASE").Value  # This can be 1,2,3
        ThermFluidPhase = self.BLK.Elements(Blockname).Elements("Input").Elements("CPHASE").Value  # "V" or "L"
        ThermFluidPhaseNumber = self.BLK.Elements(Blockname).Elements("Input").Elements("CNPHASE").Value  # 1 ,2 ,3

        StreaminPortList = self.BLK.Elements(Blockname).Elements("Ports").Elements("P(OUT)").Elements
        ListingOfStreamnamesinProductphase = []
        for Streams in StreaminPortList:
            ListingOfStreamnamesinProductphase.append(Streams.Name)

        Streamphase = self.BLK.Elements(Blockname).Elements("Input").Elements("PROD_PHASE").Elements(
            ListingOfStreamnamesinProductphase[0]).Value
        ActivateReaction_YES_NO = self.BLK.Elements(Blockname).Elements("Input").Elements("REACSYS").Value
        InletProcessflowPressure = self.BLK.Elements(Blockname).Elements("Input").Elements("PRES").Value
        InletThermalfluidPressure = self.BLK.Elements(Blockname).Elements("Input").Elements("CPRES").Value
        PressuredropCalulationOption = self.BLK.Elements(Blockname).Elements("Input").Elements("OPT_PDROP").Value
        ThermalfluidPressureDrop = self.BLK.Elements(Blockname).Elements("Input").Elements("CPDROP").Value
        ProcessflowPressureDrop = self.BLK.Elements(Blockname).Elements("Input").Elements("PDROP").Value
        Roughnessvalue = self.BLK.Elements(Blockname).Elements("Input").Elements("ROUGHNESS").Value
        PressuredropCorrelation = self.BLK.Elements(Blockname).Elements("Input").Elements("DP_FCOR").Value
        CorrectionFactor = self.BLK.Elements(Blockname).Elements("Input").Elements("DP_MULT").Value
        HoldupCalculationOption = self.BLK.Elements(Blockname).Elements("Input").Elements("OPT_HOLDUP").Value
        HoldupCorrelation = self.BLK.Elements(Blockname).Elements("Input").Elements("DP_HCOR").Value
        CatalystPresentOption = self.BLK.Elements(Blockname).Elements("Input").Elements("CAT_PRESENT").Value
        IgnoreCatalystVolume = self.BLK.Elements(Blockname).Elements("Input").Elements("IGN_CAT_VOL").Value
        WeightOfCatalystLoaded = self.BLK.Elements(Blockname).Elements("Input").Elements("CATWT").Value
        ParticleDensity = self.BLK.Elements(Blockname).Elements("Input").Elements("CAT_RHO").Value
        BedVoidage = self.BLK.Elements(Blockname).Elements("Input").Elements("BED_VOIDAGE").Value

        Dictionary = {
            "TYPE": TYPE,
            "Operating_conditions": Operating_conditions,
            "ReactorTemperature": ReactorTemperature,
            "Constant_Temp": Constant_Temp,
            "OutletTemp": OutletTemp,
            "U": U,
            "Activate_YES_NO": Activate_YES_NO,
            "Number_of_Tubes": Number_of_Tubes,
            "TubeLength": TubeLength,
            "TubeDiameter": TubeDiameter,
            "Phase": Phase,
            "Phasenumber": Phasenumber,
            "ThermFluidPhase": ThermFluidPhase,
            "ThermFluidPhaseNumber": ThermFluidPhaseNumber,
            "Streamphase": Streamphase,
            "ActivateReaction_YES_NO": ActivateReaction_YES_NO,
            "InletProcessflowPressure": InletProcessflowPressure,
            "InletThermalfluidPressure": InletThermalfluidPressure,
            "PressuredropCalulationOption": PressuredropCalulationOption,
            "ThermalfluidPressureDrop": ThermalfluidPressureDrop,
            "ProcessflowPressureDrop": ProcessflowPressureDrop,
            "Roughnessvalue": Roughnessvalue,
            "ThermalfluidPressureDrop": ThermalfluidPressureDrop,
            "PressuredropCorrelation": PressuredropCorrelation,
            "CorrectionFactor": CorrectionFactor,
            "HoldupCalculationOption": HoldupCalculationOption,
            "HoldupCorrelation": HoldupCorrelation,
            "CatalystPresentOption": CatalystPresentOption,
            "IgnoreCatalystVolume": IgnoreCatalystVolume,
            "WeightOfCatalystLoaded": WeightOfCatalystLoaded,
            "ParticleDensity": ParticleDensity,
            "BedVoidage": BedVoidage
        }
        return Dictionary


    # PAGE 1     Reactor Type

    def BLK_RPLUG_Set_TYPE(self, Blockname: str, TYPE: Literal[
        "T-SPEC", "ADIABATIC", "TCOOL-SPEC", "CO-COOL", "TCOOL-PROF", "QFLUX-PROF"]) -> None:
        '''defining the typ of Reactor which changes the necessary Inputs to make it run. Possibilities are: “T-SPEC” ”ADIABATIC” ”TCOOL-SPEC” “CO-COOL” “TCOOL-PROF” “QFLUX-PROF” '''
        self.BLK.Elements(Blockname).Elements("Input").Elements("TYPE").Value = TYPE

    # You chose Reactor with specific temperature:
    def BLK_RPLUG_Set_T_SPEC_Operating_condition(self, Blockname: str, Operating_conditions: Literal[
        "INLET-TEMP", "CONST-TEMP", "TEMP-PROF"]) -> None:
        self.BLK.Elements(Blockname).Elements("Input").Elements(
            "OPT_TSPEC").Value = Operating_conditions  # Chose between INLET-TEMP, CONST-TEMP, TEMP-PROF
        # if you chose INLET-TEMP:
        # Nothing is needed
        # if you chose CONST-TEMP:

    def BLK_RPLUG_Set_T_SPEC_Constant_Temp(self, Blockname, ReactorTemperature):
        self.BLK.Elements(Blockname).Elements("Input").Elements("REAC_TEMP").Value = ReactorTemperature
        # if you chose Temperature Profile:

    # def BLK_RPLUG_Set_T_SPEC_TemperatureProfil(self, Blockname:str, TemperatureList: list[float], LocationList: list[float]) -> None:
    def BLK_RPLUG_Set_T_SPEC_TemperatureProfil(self, Blockname, TemperatureList, LocationList):
        """Sets the Temperature Profile in side of the Column

        Args:
            Blockname: String which gives the name of Block.
            LocationList: List of location values which define where what temperature is found in the column
            TemperatureList: List of temperature values
        """
        # Check to see if it is the same size
        if len(TemperatureList) != len(LocationList):
            raise Exception('TemperatureList and LocationList need to have the same length! PList:{} LList: {}'.format(
                len(TemperatureList), len(LocationList)))
        i = 0
        for Temp in TemperatureList:
            listpositionname = "#" + str(i)
            self.BLK.Elements(Blockname).Elements("Input").Elements("SPEC_TEMP").Elements(listpositionname).Value = Temp
            i = i + 1
        i = 0
        for Location in LocationList:
            listpositionname = "#" + str(i)
            self.BLK.Elements(Blockname).Elements("Input").Elements("LOC").Elements(listpositionname).Value = Location
            i = i + 1
        i = 0


















#RADFRAC INPUTS
    def BLK_RADFRAC_GET_ME_ALL_INPUTS_BACK(self, Blockname: str) -> Dict[str, Union[str, float, int]]:
        """Retrieves all the Inputs and returns Dictionary with Values

        Does not include all aspects of a Aspen Simulationsheet, for this look at Exports

        Args:
            Blockname: String which gives the name of Block.
        """

        # PAGE 1         Configuration
        CalculationType = self.BLK.Elements(Blockname).Elements("Input").Elements("CALC_MODE").Value
        NStage = self.BLK.Elements(Blockname).Elements("Input").Elements("NSTAGE").Value
        CondenserType = self.BLK.Elements(Blockname).Elements("Input").Elements("CONDENSER").Value
        ReboilerType = self.BLK.Elements(Blockname).Elements("Input").Elements("REBOILER").Value
        Phase = self.BLK.Elements(Blockname).Elements("Input").Elements("Phase").Value  # This can be V L or S
        Phasenumber = self.BLK.Elements(Blockname).Elements("Input").Elements("NPhase").Value  # This can be 1,2,3
        ConvergenceMethod = self.BLK.Elements(Blockname).Elements("Input").Elements("CONV_METH").Value
        Refluxratio = self.BLK.Elements(Blockname).Elements("Input").Elements("BASIS_RR").Value
        Refluxrate = self.BLK.Elements(Blockname).Elements("Input").Elements("BASIS_L1").Value
        BoilupRate = self.BLK.Elements(Blockname).Elements("Input").Elements("BASIS_VN").Value
        BoilupRatio = self.BLK.Elements(Blockname).Elements("Input").Elements("BASIS_BR").Value
        CondenserDuty = self.BLK.Elements(Blockname).Elements("Input").Elements("Q1").Value
        ReboilerDuty = self.BLK.Elements(Blockname).Elements("Input").Elements("QN").Value
        TotalDestillateFlowrate = self.BLK.Elements(Blockname).Elements("Input").Elements("BASIS_D").Value
        LiquidBottomRate = self.BLK.Elements(Blockname).Elements("Input").Elements("BASIS_B").Value
        DestillateToFeedRatio = self.BLK.Elements(Blockname).Elements("Input").Elements("BASIS_D:F").Value
        BottomToFeedRatio = self.BLK.Elements(Blockname).Elements("Input").Elements("BASIS_B:F").Value
        # Page 2     Streams
        FeedStreamNameNode = self.BLK.Elements(Blockname).Elements("Ports").Elements("F(IN)").Element
        for FeedStreamName in FeedStreamNameNode:
            FeedStage = self.BLK.Elements(Blockname).Elements("Input").Elements("FEED_STAGE").Elements(
                FeedStreamName).Value
            FeedStageLocation = self.BLK.Elements(Blockname).Elements("Input").Elements("FEED_CONVE2").Elements(
                FeedStreamName).Value

        CompleteProductStreamNameList = []
        ProductStreamNameList4LiquidDestillate = self.BLK.Elements(Blockname).Elements("Ports").Elements(
            "LD(OUT)").Element
        for ProductStreamName in ProductStreamNameList4LiquidDestillate:
            CompleteProductStreamNameList.append(ProductStreamName)
        ProductStreamNameList4Bottoms = self.BLK.Elements(Blockname).Elements("Ports").Elements("B(OUT)").Element
        for ProductStreamName in ProductStreamNameList4Bottoms:
            CompleteProductStreamNameList.append(ProductStreamName)

        ProductStageLocationList = []
        ProductPhaseList = []
        for ProductStreamName in CompleteProductStreamNameList:
            ProductStageLocationList.append(
                self.BLK.Elements(Blockname).Elements("Input").Elements("PROD_STAGE").Elements(ProductStreamName).Value)
            ProductPhaseList.append(
                self.BLK.Elements(Blockname).Elements("Input").Elements("PROD_PHASE").Elements(ProductStreamName).Value)

            # Page 3     PRESSURE
        PressurePerspectiveOption = self.BLK.Elements(Blockname).Elements("Input").Elements("VIEW_PRES").Value
        CondenserPressure = self.BLK.Elements(Blockname).Elements("Input").Elements("PRES1").Value
        CondenserPressureDrop = self.BLK.Elements(Blockname).Elements("Input").Elements("PRES2").Value
        StagePressureDrop = self.BLK.Elements(Blockname).Elements("Input").Elements("DP_STAGE").Value
        # PAGE 4         Condenser
        CoolRefluxandDestillate = self.BLK.Elements(Blockname).Elements("Input").Elements("SC_OPTION").Value
        CondenserTempOption = self.BLK.Elements(Blockname).Elements("Input").Elements("OPT_SUBCOOL").Value
        SubcooledTemp = self.BLK.Elements(Blockname).Elements("Input").Elements("SC_TEMP").Value
        DegreeSubcooled = self.BLK.Elements(Blockname).Elements("Input").Elements("DEGSUB").Value
        CondenserOption = self.BLK.Elements(Blockname).Elements("Input").Elements("OPT_COND").Value
        VaporTemp = self.BLK.Elements(Blockname).Elements("Input").Elements("T1").Value
        VaporFraction = self.BLK.Elements(Blockname).Elements("Input").Elements("BASIS_RDV").Value
        ThermosyphonOption = self.BLK.Elements(Blockname).Elements("Input").Elements("OPT_TH_REB").Value
        ReboilerCirculationFlow = self.BLK.Elements(Blockname).Elements("Input").Elements("TH_FLOW").Value
        OutletTemperature = self.BLK.Elements(Blockname).Elements("Input").Elements("TH_TEMP").Value
        ReboilerOutletPressure = self.BLK.Elements(Blockname).Elements("Input").Elements("TH_PRES").Value
        ReboilerReturnLocation = self.BLK.Elements(Blockname).Elements("Input").Elements("RETURN_CONV").Value
        ReboilerConfiguration = self.BLK.Elements(Blockname).Elements("Input").Elements("TSR_CONFIG").Value

        Dictionary = {
            # Page 1:
            "CalculationType": CalculationType,
            "NStage": NStage,
            "CondenserType": CondenserType,
            "ReboilerType": ReboilerType,
            "Phase": Phase,
            "Phasenumber": Phasenumber,
            "ConvergenceMethod": ConvergenceMethod,
            "Refluxratio": Refluxratio,
            "Refluxrate": Refluxrate,
            "BoilupRate": BoilupRate,
            "BoilupRatio": BoilupRatio,
            "CondenserDuty": CondenserDuty,
            "ReboilerDuty": ReboilerDuty,
            "TotalDestillateFlowrate": TotalDestillateFlowrate,
            "LiquidBottomRate": LiquidBottomRate,
            "DestillateToFeedRatio": DestillateToFeedRatio,
            "BottomToFeedRatio": BottomToFeedRatio,
            # Page 2
            "FeedStage": FeedStage,
            "FeedStageLocation": FeedStageLocation,
            "ProductStageLocation": ProductStageLocationList,
            "ProductPhaseList": ProductPhaseList,
            # Page 3
            "PressurePerspectiveOption": PressurePerspectiveOption,
            "CondenserPressure": CondenserPressure,
            "CondenserPressureDrop": CondenserPressureDrop,
            "StagePressureDrop": StagePressureDrop,
            # Page 4
            "CondenserTempOption": CondenserTempOption,
            "CoolRefluxandDestillate": CoolRefluxandDestillate,
            "CondenserTempOption": CondenserTempOption,
            "SubcooledTemp": SubcooledTemp,
            "DegreeSubcooled": DegreeSubcooled,
            "CoolRefluxandDestillate": CoolRefluxandDestillate,
            "CondenserOption": CondenserOption,
            "VaporTemp": VaporTemp,
            "VaporFraction": VaporFraction,
            "ThermosyphonOption": ThermosyphonOption,
            "ReboilerCirculationFlow": ReboilerCirculationFlow,
            "OutletTemperature": OutletTemperature,
            "ReboilerOutletPressure": ReboilerOutletPressure,
            "ReboilerReturnLocation": ReboilerReturnLocation,
            "ReboilerConfiguration": ReboilerConfiguration,
        }
        return Dictionary

    ##RADFRAC

    # PAGE 1         Configuration
    def BLK_RADFRAC_Set_CalculationType(self, Blockname: str, CalculationType: Literal[
        "RIG-RATE", "EQUILIBRIUM"]) -> None:  # This can be RIG-RATE,  EQUILIBRIUM
        self.BLK.Elements(Blockname).Elements("Input").Elements("CALC_MODE").Value = CalculationType

    def BLK_RADFRAC_Set_NSTAGE(self, Blockname, NStage):
        self.BLK.Elements(Blockname).Elements("Input").Elements("NSTAGE").Value = NStage

    def BLK_RADFRAC_Set_CondenserType(self, Blockname: str, CondenserType: Literal[
        "NONE", "TOTAL", "PARTIAL-V", "PARTIAL-V-L"]) -> None:  # THIS can be NONE, TOTAL, PARTIAL-V, PARTIAL-V-L        Very important for Page 4
        self.BLK.Elements(Blockname).Elements("Input").Elements("CONDENSER").Value = CondenserType

    def BLK_RADFRAC_Set_ReboilerType(self, Blockname: str, ReboilerType: Literal[
        "NONE", "KETTLE", "THERMOSYPHON"]):  # Can be NONE, KETTLE, THERMOSYPHON, This is important for Page 5
        self.BLK.Elements(Blockname).Elements("Input").Elements("REBOILER").Value = ReboilerType

    def BLK_RADFRAC_Set_Phases(self, Blockname: str, Phase: Ph, Phasenumber: Phnum):
        self.BLK.Elements(Blockname).Elements("Input").Elements("Phase").Value = Phase  # This can be V L or S
        self.BLK.Elements(Blockname).Elements("Input").Elements("NPhase").Value = Phasenumber  # This can be 1,2,3

    def BLK_RADFRAC_Set_ConvergenceMethod(self, Blockname: str, ConvergenceMethod: Literal[
        "STANDARD", "PETROLEUM", "NONIDEAL", "AZEOTROPIC", "CRYOGENIX", "OTHERS"]) -> None:  # This can be STANDARD, PETROLEUM, NONIDEAL, AZEOTROPIC, CRYOGENIX, OTHERS
        self.BLK.Elements(Blockname).Elements("Input").Elements("CONV_METH").Value = ConvergenceMethod

    def BLK_RADFRAC_Set_Refluxratio(self, Blockname, Refluxratio):
        self.BLK.Elements(Blockname).Elements("Input").Elements("BASIS_RR").Value = Refluxratio

    def BLK_RADFRAC_Set_Refluxrate(self, Blockname, Refluxrate):
        self.BLK.Elements(Blockname).Elements("Input").Elements("BASIS_L1").Value = Refluxrate

    def BLK_RADFRAC_Set_BoilupRate(self, Blockname, BoilupRate):
        self.BLK.Elements(Blockname).Elements("Input").Elements("BASIS_VN").Value = BoilupRate

    def BLK_RADFRAC_Set_BoilupRatio(self, Blockname, BoilupRatio):
        self.BLK.Elements(Blockname).Elements("Input").Elements("BASIS_BR").Value = BoilupRatio

    def BLK_RADFRAC_Set_CondenserDuty(self, Blockname, CondenserDuty):
        self.BLK.Elements(Blockname).Elements("Input").Elements("Q1").Value = CondenserDuty

    def BLK_RADFRAC_Set_ReboilerDuty(self, Blockname, ReboilerDuty):
        self.BLK.Elements(Blockname).Elements("Input").Elements("QN").Value = ReboilerDuty

    def BLK_RADFRAC_Set_TotalDestillateFlowrate(self, Blockname, TotalDestillateFlowrate):
        self.BLK.Elements(Blockname).Elements("Input").Elements("BASIS_D").Value = TotalDestillateFlowrate

    def BLK_RADFRAC_Set_LiquidBottomRate(self, Blockname, LiquidBottomRate):
        self.BLK.Elements(Blockname).Elements("Input").Elements("BASIS_B").Value = LiquidBottomRate

    def BLK_RADFRAC_Set_DestillateToFeedRatio(self, Blockname, DestillateToFeedRatio):
        self.BLK.Elements(Blockname).Elements("Input").Elements("BASIS_D:F").Value = DestillateToFeedRatio

    def BLK_RADFRAC_Set_BottomToFeedRatio(self, Blockname, BottomToFeedRatio):
        self.BLK.Elements(Blockname).Elements("Input").Elements("BASIS_B:F").Value = BottomToFeedRatio
















#REQUIL INPUTS
    def BLK_REQUIL_Set_Temperature(self, Blockname, RequilTemp):
        self.BLK.BLK.Elements(Blockname).Elements("Input").Elements("TEMP").Value = RequilTemp

    def BLK_REQUIL_Set_Pressure(self, Blockname, RequilPres):
        self.BLK.BLK.Elements(Blockname).Elements("Input").Elements("PRES").Value = RequilPres
















#RSTOIC INPUTS
    def BLK_RSTOIC_Set_Temperature(self, Blockname, RstoicTemp):
        self.BLK.BLK.Elements(Blockname).Elements("Input").Elements("TEMP").Value = RstoicTemp

    def BLK_RSTOIC_Set_Pressure(self, Blockname, RstoicPres):
        self.BLK.BLK.Elements(Blockname).Elements("Input").Elements("PRES").Value = RstoicPres














#RGIBBS INPUTS
    def BLK_RGIBBS_Set_Temperature(self, Blockname, RgibbsTemp):
        self.BLK.BLK.Elements(Blockname).Elements("Input").Elements("TEMP").Value = RgibbsTemp

    def BLK_RGIBBS_Set_Pressure(self, Blockname, RgibbsPres):
        self.BLK.BLK.Elements(Blockname).Elements("Input").Elements("PRES").Value = RgibbsPres













#DEFINE RUN
    def Run(self) -> bool:
        """Runs Aspen once and stores the last failure reason.

        Retry policy is handled by the caller (environment/reset logic) so we
        avoid nested retries here, which can make COM hangs harder to debug.
        """
        self.last_run_error = None
        self.last_per_error = None
        start = time.time()
        try:
            self.AspenSimulation.Engine.Run2()
            print(f"Runtime = {time.time() - start}")
            per_error = self.AspenSimulation.Tree.Elements("Data").Elements("Results Summary").Elements(
                           "Run-Status").Elements("Output").Elements("PER_ERROR").Value
            self.last_per_error = per_error
            print("per_error value : ", per_error)
        except Exception as err:
            self.last_run_error = f"Aspen run failed while executing or reading PER_ERROR: {err}"
            print(self.last_run_error)
            return False

        if per_error == 0:
            return True

        self.last_run_error = f"Aspen PER_ERROR was non-zero (PER_ERROR={self.last_per_error})."
        print("Aspen PER_ERROR was non-zero.")
        return False

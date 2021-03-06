﻿using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NationalInstruments.VisaNS;
using System.IO;
using NationalInstruments.DAQmx;
using NationalInstruments;
namespace AcutrolVS2010
{
    /*
     * @author: Hai Tang (william.sea@jhu.edu)
     * Remarks:
     * 1.Channel 1 is the only channel in use. So the default channel is channel 1.
     * 
     * TODO: 
     * 1.debug ECP87 remote touch commend: not working correctly. But when using loc mode touch screen, it's working correct
     * 2.real time graph with line header
     * 3.display coil signals takes too long time (reduce size)
     * 
     * DEBUG LIST:
     * 1. Counter (return real time - sometimes causing errors)
     * 
     * Additional:
     * 1.setup abs limits (1140,1141) and ROT/LIN?
     * 
     * 
     * 
Author: Hai Tang
Email: william.sea@jhu.edu
Phone: 937-238-3359

----------------- Where to Find my C# software -----------------
For Windows 7 Computer: c:\users\vnel\documents\visual studio 2013\Projects

For Windows XP Computer: C:\Documents and Settings\Administrator\Desktop\Acutrol_VS2010
Shortcut for Visual Studio 2010 on Desktop named: Hai Tang Acutrol Control START HERE THEN LOAD AcutrolVS2010


----------------- How to Use my C# software -----------------
0. For the Windows XP computer;
1. Turn on Acutrol, in the meanwhile make sure the coil frame is fixed (IMPORTANT);
2. Click the Visual Studio 2010 icon on Desktop, open project named AcutrolVS2010;
3. After the Acutrol is totally turned on, run the program (click the green triangle icon);
4. You can now see the real time chair movement waveform (Chair Position Chart) and it’s position/rate/acceleration values in text boxes (Parameter Reading Section);
5. Click Initialization (bold) button to set up the ECP environment (default ECP 87);
	5.1. Note: Do NOT manually rotate the chair or touch the Acutrol touch-screen during the initialization process, otherwise it might not work.
	5.2. You will hear 4 beeps which takes about 20 seconds in total.
6. Cut Analog Input by clicking the button named Cut Analog Input if necessary;
7. The default mode is Position mode. You can switch modes through the drop-down list under "Select Mode";
8. For executing a sequence of sinusoidal input
	8.1. Fill the sequence list (magnitude, frequency and number of cycles desired) in the right side of GUI (Sinusoidal Sequences Section).
	8.2. Fill the Desired Limits for rate and acceleration. The default values are 150 deg/s and 1200 deg/s^2 respectively.
	8.3. The max rate and max acceleration values can be calculated by clicking Calculation Max button. You can use it to check if any input exceed the limits. If some sequences exceed, a pop up window will show and indicate which sequences exceed.045
	8.4. For those sequence numbers (1 to 16) not in use, just put 0 in the Number of Cycles text box. Once the program detects one zero in Number of Cycles, all the rest sequences will not be considered, assuming all the rest are zeros as well.
	8.5. DO NOT put 0 in any frequency text box, instead use a very small number to ensure security (like 0.001 Hz). 
	8.6. Click Execute button to execute all the sequences with non-zero number of cycles. 
	8.7. Executability will be checked every time you click Execute button, i.e. if one of the sequence exceeds desired limits, the whole sequences will not be executed and a pop up window is shown to indicate what is wrong;
	8.8. You can save sequences data by clicking Save to File button after putting a File Name. The default saving path is Deskop/Saved_Sequences
	8.9. You can read saved sequences by clicking Load Sequences File, after choosing a file, all sequences information will be automatically filled.
9. The saved chair movement and eye coil data are in folder named Chair_Results and Coil_Results respectively, in Desktop;
	9.1. Once you start executing a sinusoidal sequence, the chair movements and eye coild signals will automatically start recording, and after all the sequences are executed, the recording automatically stops.
	9.2. You can manually Start and Stop eye coil recording by clicking corresponding buttons
10. Display. 
	10.1.For saved chair movement data and eye coil data, You can Select File to Display by clicking corresponding buttons.
	10.2.You can Return Real Time to stop displaying saved data to display real time waveform.
	10.3.You can click on the chart and select certain area to zoom in and see the details.
11. The Parameter Limitations Settings section is used to set and show all the limits. Fill the text boxes and click Execute to set limits, click Reset to put all values to zero.
12. The Simple Commends section is used to send simple position/rate/acceleration commends, make sure to select the correct Mode before you click Execute.
	12.0. Auto Interlock Close is turned off for this section for safety purpose. 
	12.1. Manual operation after initialization:
	12.2. Set limits: type in the limits you desire in the Parameter Limitation Settings section, and click Execute. The default values are 20 Deg/s for rate and 50 Deg/s^2 for acceleration.
	12.3. Check limits (IMPORTANT): click Show Limits in Windows Settings section, and check the values in the displaying windows of Acutrol to make sure they've been correctly set. Or you can choose each window to dispay any limit value you want by selecting from the drop-down lists.501100
	12.4. Interlock close: click the Interlock Close button to close interlock. It will slowly return the 0 position (under the default position mode).
	12.5. Select Mode: select the desired mode (position, relative/absolute rate).
	12.6. Type in a number in the textbox corresponding to your selected mode. For position and rate mode, type values in Position Commend and Rate Commend under Simple Commends section. 
	12.7. Click Execute. The chair starts moving according to your limits settings and commends. (Position commends: go to certain position. Rate commends: rotate in positive direction, i.e. Clickwisely, with certain rate).
	12.8. When you switching between position and rate modes, it will remember your last commend. For example, your last position commend is 0 Deg, and then you switch to rate mode with rate commend 20 Deg/s, then you directly switch mode back to position mode, it will go back to 0 position, then you switch back to rate mode, it starts rotating with 20 Deg/s again.
	12.9. Stop: click the Stop button, which will reset the default limits and go back to 0 position under position mode. Or you can manually go to 0 position under position mode, or set rate value to 0 under rate mode.
 	12.10. When you finish, click Interlock Open button.
13. The Sinusoidal Input section is used to execute sinusoidal input. 
	13.0. Auto Interlock Close is turned off for this section for safety purpose. 
	13.1. manual operations after initialization:
	13.2. Set limits for default 20 Deg/s for rate and 50 Deg/s^2 for acceleration, check the limits on Acutrol screen by clicking Show Limits.
	13.3. Click Interlock Close. You must apply lower limits to perform Interlock Close, otherwise it will cause a bumping sound.
	13.4. Go back to zero position using position mode if it is not in zero position.
	13.5. Increase Limits (IMPORTANT). The suggested higher value are 150 Deg/s and 1200 Deg/s^2.
	13.6. Go to Synthesis Mode.
	13.7. Type in commend for Magnitude, Frequency and Phase under Sinusoidal Input section.
	13.8. Click Execute, the legality of input will be checked automatically. If the calculated max rate and acceleratio exceed the limit, a pop up window will show and indicate the calculated max values.
	13.9. Click Stop to slowly stop the chair.
	13.10. Click Interlock Open.
14. The Windows Settings section is used to select what to show in the 6 windows on Acutrol. 
	14.1.The Default shows Estimated Position/Vecocity, FilteredAccelEstimate, ProfilerPosition/Velocity/AccelCommend
	14.2.The Show Limits shows some representative limits, which are MinimunPositionLimit, RateTripLimit, SynthesisModeVelocity/AccelLimit, VelocityAbsoluteLimit and AbortModeAccelLimit. You can change them to what you want in the code.
15. Buttons: function for Return Local Mode, Remote Mode, Interlock Close, Interlock Open, Interlock Reset are indicated in button names.
16. Pulse Inputs:
	16.1. Currently it is able to rotate the motor with correct pulse inputs using Thread.Sleep(delayTime) to make the delays between pulses, but is will pause all the threads and the recording for chair and coil are also paused.
	16.2. Need to use System.Timer to do the job for delays. 
	16.3. I've built the GUI and written all the codes needed. The only thing left is to debug the sysTimer to make the program running correctly.
	16.4. The related codes are at the end part of the whole codes.


----------------- How to Access my Git Repository -----------------
Visual Studio 2010 Version: https://github.com/williamsea/Acutrol_VS2010
Visual Studio 2013 Version: https://github.com/williamsea/Acutrol

You can Clone the project or just download the .zip package from website.
You can follow my updates/changes by checking my commits (https://github.com/williamsea/Acutrol_VS2010/commits/master).


----------------- Common Questions and Solutions -----------------
1. Chair movement waveform jumping zig-zagly without any meanings:
	Solution: Switch between ECP 4 and ECP 87, repeat several times if it is not solved. If the numerical values are meaningful but the waveform is not responding, do local mode and then back to remote mode.
2. Buffle overflow
	Solution: restart the program (click the stop icon on visual studio tool bar to close the form, and then click then start icon).


----------------- Additional Related Documents in Server under Hai Tang -----------------
1. C# Acutrol.docx 
	-Including information to set up environment for a new computer, like how to install libraries
	-Including common useful commends
2. Notebook.docx
	-Including the daily log
     * 
     * 
     * 
     */

    public partial class Form1 : Form
    {
        //GPIB Connection string
        string sAddress = "GPIB0::1::INSTR";
        //The VNA uses a message based session
        MessageBasedSession mbSession = null;
        //Open a generic Session first
        Session mySession = null;

        //parameters for sinusoidal input
        double posReadValue = 1000;//an arbitray non-zero number used to check if position has been set back to zero position
        double cycleCounter = 0;//used for synthesis mode counting cycles
        //List<double> TargetCycleCount = new List<double>();
        //List<double> TargetMagn = new List<double>();
        //List<double> TargetFreq = new List<double>();
        int totalSequencesNum = 16 ;
        double[] TargetCycleCount = new double[16] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };//Must be initially Set to 0!!! Otherwise may cause problem!
        double[] TargetMagn = new double[16] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
        double[] TargetFreq = new double[16] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
        System.Timers.Timer sysTimer;
        int seqCount = 0; //count the sequence to be executed
        double zeroTrigger = 0.1;//0.01;
        TextBox[] textBoxSeqMag = new TextBox[16];
        TextBox[] textBoxSeqFreq = new TextBox[16];
        TextBox[] textBoxSeqCycle = new TextBox[16];
        TextBox[] textBox_maxRate = new TextBox[16];
        TextBox[] textBox_maxAccel = new TextBox[16];
        Boolean isExecutable = false;
        string sequencesSavingPath = "C:\\Documents and Settings\\Administrator\\Desktop\\Saved_Sequences";
        string sequenceReadingPath;
        FileStream sequenceFileStream;
        StreamWriter sequenceFileStreamWriter;
        StreamReader sequenceFileStreamReader; 

        //Machine representation codes
        String Position = "P";
        String Rate = "R";
        String Acceleration = "A";
        String PositionMode = "P";
        String AbsRateMode = "A";
        String RelativeRateMode = "R";
        String RateMode = "R";
        String SynthesisMode = "S";
        String AbortMode = "A";

        //Parameters for contents of Display Windows
        String RawPositionFeedback = "1081";
        String EstimatedPosition = "1082";//default 1: position
        String EstimatedVelocity = "1083";//default 2: velocity
        String FilteredVelocityEstimate = "1038";
        String FilteredAccelEstimate = "1039";//default 3: acceleration
        String ProfilerPositionCommend = "1008";//default 4: position commend
        String ProfilerVelocityCommend = "1009";//default 5: velocity commend
        String ProfilerAccelCommend = "1010";//default 6: acceleration commend
        //Limits
        String MinimunPositionLimit = "1145"; String OvarMinPosLim = "1";
        String MaximunPositionLimit = "1144"; String OvarMaxPosLim = "2";
        String SynthesisModeVelocityLimit = "1153"; String OvarSynModeVeloLim = "3";
        String SynthesisModeAccelLimit = "1154"; String OvarSynModeAccLim = "4"; 
        String PositionModeVelocityLimit = "1147"; String OvarPosModeVeloLim = "5";
        //only 1-5 is allowed! for ECP variables
        String PositionModeAccelLimit = "1148"; String OvarPosModeAccLim = "6";
        String RateModeVelocityLimit = "1150"; String OvarRateModeVeloLim = "7";
        String RateModeAccelLimit = "1151"; String OvarRateModeAccLim = "8";
        String AbortModeVelocityLimit = "1159";
        String AbortModeAccelLimit = "1160";
        String VelocityAbsoluteLimit = "1140";
        String AccelAbsoluteLimit = "1141";
        String RateTripLimit = "1146";

        //display information
        double[] PosValArray = new double[100000]; //array used to store position values
        int disp = 0; //counter used for display data
        int counter = 0; //counter used to store data array
        int DispLength = 300; //30s
        double[,] coilValArrayReduced = new double[16, 100000];
        bool showCoilWaveform = false;
        int coilValueCounter = 0; //counter used to store coil data list, every time there are 100 more set (16 channels) of data get stored

        //parameters for document saving and reading (Chair Movements)
        string savingPath = "C:\\Documents and Settings\\Administrator\\Desktop\\Chair_Results";
        string readingPath;
        FileStream fileStream;
        StreamWriter streamWriter;
        Boolean RecordCtr = false;
        //parameters for document saving and reading (Coil Movements)
        string coilSavingPath = "C:\\Documents and Settings\\Administrator\\Desktop\\Coil_Results";
        string coilReadingPath;
        FileStream coilFileStream;
        StreamWriter fileStreamWriter;
        StreamReader fileStreamReader;
        Boolean CoilRecordCtr = false;
        ArrayList savedCoilData;
        //int DisplayedRowNum=100;

        //parameters for reading signals from coil from BNC-2090
        private AnalogMultiChannelReader analogInReader;
        private NationalInstruments.DAQmx.Task myTask; //given full reference to avoid confusion with system.threading.task
        private NationalInstruments.DAQmx.Task runningTask;
        private AsyncCallback analogCallback;
        //private AnalogWaveform<double>[] data;//need add "using NationalInstruments"
        private double[,] data;
        private DataColumn[] dataColumn = null;
        private DataTable dataTable = null;
        int samplesPerChannel = 100;//reduced to 100 to make the display more smooth (1000); finish 100 samples then callback
        int timingRate = 1000;//otherwise too fast, ADC conversion attempts before previous conversion was completed (10000); actual sampling rate


        //parameters for pulse inputs
        int targetCycleNum = 0;
        int currentCycleCount = 0;
        double pulseDelay = 0;
        string DepartOrReturn;

        public Form1()
        {
            InitializeComponent();
            //comboBoxSelectMode.Text = "---Please Select Mode---";//Already set the default mode to be position mode
            comboBoxSelectMode.Text = "Position Mode (default)";
            comboBoxSelectMode.Items.Add("Position Mode");
            comboBoxSelectMode.Items.Add("Relative Rate Mode");
            comboBoxSelectMode.Items.Add("Absolute Rate Mode");
            comboBoxSelectMode.Items.Add("Synthesis Mode");

            initComboBoxWindows(comboBox_window1);
            initComboBoxWindows(comboBox_window2);
            initComboBoxWindows(comboBox_window3);
            initComboBoxWindows(comboBox_window4);
            initComboBoxWindows(comboBox_window5);
            initComboBoxWindows(comboBox_window6);

            //comboBoxSelectMode.SelectedIndex = 1;
            openMySession();

            ////Setup ECP 80 then 87 using TOUCH commend
            //mbSession.Write("u:t 179,179,176,32,152,144,32,50,176,181,176,32,152,151,32,50,176,181,181,181 \n");

            ////Set analog inputs gain to 0 
            //System.Threading.Thread.Sleep(10000);//wait for 10s for ECP to be restored
            //mbSession.Write(":u:t 180,177,178,50,32,144,32,176,181,179,50,32,144,32,176,181,181,181 \n");

            ////Set ROT/LIN value to 1 to use Limited Motion
            //System.Threading.Thread.Sleep(1000);//wait for 1s
            //mbSession.Write(":u:t 180,179,180,32,145,32,50,176,181,181,181 \n");

            ////Set the Veloc/Accel Absolute Limits
            //System.Threading.Thread.Sleep(1000);//wait for 1s
            //mbSession.Write(":u:t 180,180,148,176,177,32,145,144,32,51,176,180,49,176,181,181,181,181 \n");

            //set the default mode to be position mode
            SelectMode(PositionMode);

            ////Set default limitations just after initialization, before all the other actions
            //SetAllLimits();//default values

            //MessageBox.Show("ECP has been set to 80 then 87; Analog input 1&2 gain has been set to 0; ROT/LIN value set to 1 to use Limited Motion; Veloc/Accel Absolute Limits are default;  Default mode is Position Mode; All limits are set to default.");

            //Setup Ovariable
            //mbSession.Write(":c:o " + OvarMinPosLim + " , " + MinimunPositionLimit + " \n");
            //mbSession.Write(":c:o " + OvarMaxPosLim + " , " + MaximunPositionLimit + " \n");
            //mbSession.Write(":c:o " + OvarSynModeVeloLim + " , " + SynthesisModeVelocityLimit + " \n");
            //mbSession.Write(":c:o " + OvarSynModeAccLim + " , " + SynthesisModeAccelLimit + " \n");
            //mbSession.Write(":c:o " + OvarPosModeVeloLim + " , " + PositionModeVelocityLimit + " \n");//Allows only 5 data selects per ECP available for overiable config
            //mbSession.Write(":c:o " + OvarPosModeAccLim + " , " + PositionModeAccelLimit + " \n");
            //mbSession.Write(":c:o " + OvarRateModeVeloLim + " , " + RateModeVelocityLimit + " \n");
            //mbSession.Write(":c:o " + OvarRateModeAccLim + " , " + RateModeAccelLimit + " \n");

            ////For coil data display and recording
            //dataTable = new DataTable();


            InitTextBoxControls();
        }

        #region InitTextBoxControls
        private void InitTextBoxControls()
        {
            //The parameters of Sinusoidal Inputs
            ////Magnitude
            textBoxSeqMag[0] = (TextBox)textBoxSeqMag1;
            textBoxSeqMag[1] = (TextBox)textBoxSeqMag2;
            textBoxSeqMag[2] = (TextBox)textBoxSeqMag3;
            textBoxSeqMag[3] = (TextBox)textBoxSeqMag4;
            textBoxSeqMag[4] = (TextBox)textBoxSeqMag5;
            textBoxSeqMag[5] = (TextBox)textBoxSeqMag6;
            textBoxSeqMag[6] = (TextBox)textBoxSeqMag7;
            textBoxSeqMag[7] = (TextBox)textBoxSeqMag8;
            textBoxSeqMag[8] = (TextBox)textBoxSeqMag9;
            textBoxSeqMag[9] = (TextBox)textBoxSeqMag10;
            textBoxSeqMag[10] = (TextBox)textBoxSeqMag11;
            textBoxSeqMag[11] = (TextBox)textBoxSeqMag12;
            textBoxSeqMag[12] = (TextBox)textBoxSeqMag13;
            textBoxSeqMag[13] = (TextBox)textBoxSeqMag14;
            textBoxSeqMag[14] = (TextBox)textBoxSeqMag15;
            textBoxSeqMag[15] = (TextBox)textBoxSeqMag16;
            //Frequency
            textBoxSeqFreq[0] = (TextBox)textBoxSeqFreq1;
            textBoxSeqFreq[1] = (TextBox)textBoxSeqFreq2;
            textBoxSeqFreq[2] = (TextBox)textBoxSeqFreq3;
            textBoxSeqFreq[3] = (TextBox)textBoxSeqFreq4;
            textBoxSeqFreq[4] = (TextBox)textBoxSeqFreq5;
            textBoxSeqFreq[5] = (TextBox)textBoxSeqFreq6;
            textBoxSeqFreq[6] = (TextBox)textBoxSeqFreq7;
            textBoxSeqFreq[7] = (TextBox)textBoxSeqFreq8;
            textBoxSeqFreq[8] = (TextBox)textBoxSeqFreq9;
            textBoxSeqFreq[9] = (TextBox)textBoxSeqFreq10;
            textBoxSeqFreq[10] = (TextBox)textBoxSeqFreq11;
            textBoxSeqFreq[11] = (TextBox)textBoxSeqFreq12;
            textBoxSeqFreq[12] = (TextBox)textBoxSeqFreq13;
            textBoxSeqFreq[13] = (TextBox)textBoxSeqFreq14;
            textBoxSeqFreq[14] = (TextBox)textBoxSeqFreq15;
            textBoxSeqFreq[15] = (TextBox)textBoxSeqFreq16;
            //Cycle Count
            textBoxSeqCycle[0] = (TextBox)textBoxSeqCycle1;
            textBoxSeqCycle[1] = (TextBox)textBoxSeqCycle2;
            textBoxSeqCycle[2] = (TextBox)textBoxSeqCycle3;
            textBoxSeqCycle[3] = (TextBox)textBoxSeqCycle4;
            textBoxSeqCycle[4] = (TextBox)textBoxSeqCycle5;
            textBoxSeqCycle[5] = (TextBox)textBoxSeqCycle6;
            textBoxSeqCycle[6] = (TextBox)textBoxSeqCycle7;
            textBoxSeqCycle[7] = (TextBox)textBoxSeqCycle8;
            textBoxSeqCycle[8] = (TextBox)textBoxSeqCycle9;
            textBoxSeqCycle[9] = (TextBox)textBoxSeqCycle10;
            textBoxSeqCycle[10] = (TextBox)textBoxSeqCycle11;
            textBoxSeqCycle[11] = (TextBox)textBoxSeqCycle12;
            textBoxSeqCycle[12] = (TextBox)textBoxSeqCycle13;
            textBoxSeqCycle[13] = (TextBox)textBoxSeqCycle14;
            textBoxSeqCycle[14] = (TextBox)textBoxSeqCycle15;
            textBoxSeqCycle[15] = (TextBox)textBoxSeqCycle16;

            //Tried to refactor but not working
            //textBoxSeqMag[i] = (TextBox)this.Controls.Find(string.Format("textBoxSeqMag{0}", i), false).FirstOrDefault();

            //Max Rate
            textBox_maxRate[0] = (TextBox) textBox_maxRate1;
            textBox_maxRate[1] = (TextBox) textBox_maxRate2;
            textBox_maxRate[2] = (TextBox) textBox_maxRate3;
            textBox_maxRate[3] = (TextBox) textBox_maxRate4;
            textBox_maxRate[4] = (TextBox) textBox_maxRate5;
            textBox_maxRate[5] = (TextBox) textBox_maxRate6;
            textBox_maxRate[6] = (TextBox) textBox_maxRate7;
            textBox_maxRate[7] = (TextBox) textBox_maxRate8;
            textBox_maxRate[8] = (TextBox) textBox_maxRate9;
            textBox_maxRate[9] = (TextBox) textBox_maxRate10;
            textBox_maxRate[10] = (TextBox) textBox_maxRate11;
            textBox_maxRate[11] = (TextBox) textBox_maxRate12;
            textBox_maxRate[12] = (TextBox) textBox_maxRate13;
            textBox_maxRate[13] = (TextBox) textBox_maxRate14;
            textBox_maxRate[14] = (TextBox) textBox_maxRate15;
            textBox_maxRate[15] = (TextBox) textBox_maxRate16;
            //Max Acceleration
            textBox_maxAccel[0] = (TextBox) textBox_maxAccel1;
            textBox_maxAccel[1] = (TextBox) textBox_maxAccel2;
            textBox_maxAccel[2] = (TextBox) textBox_maxAccel3;
            textBox_maxAccel[3] = (TextBox) textBox_maxAccel4;
            textBox_maxAccel[4] = (TextBox) textBox_maxAccel5;
            textBox_maxAccel[5] = (TextBox) textBox_maxAccel6;
            textBox_maxAccel[6] = (TextBox) textBox_maxAccel7;
            textBox_maxAccel[7] = (TextBox) textBox_maxAccel8;
            textBox_maxAccel[8] = (TextBox) textBox_maxAccel9;
            textBox_maxAccel[9] = (TextBox) textBox_maxAccel10;
            textBox_maxAccel[10] = (TextBox) textBox_maxAccel11;
            textBox_maxAccel[11] = (TextBox) textBox_maxAccel12;
            textBox_maxAccel[12] = (TextBox) textBox_maxAccel13;
            textBox_maxAccel[13] = (TextBox) textBox_maxAccel14;
            textBox_maxAccel[14] = (TextBox) textBox_maxAccel15;
            textBox_maxAccel[15] = (TextBox) textBox_maxAccel16;
        }
        #endregion

        private void ReadSinosoidalParas()
        {
            for (int i = 0; i < 16; i++)
            {
                TargetMagn[i] = Convert.ToDouble(textBoxSeqMag[i].Text);
                TargetFreq[i] = Convert.ToDouble(textBoxSeqFreq[i].Text);
                TargetCycleCount[i] = Convert.ToDouble(textBoxSeqCycle[i].Text);
            }
        }

        #region initComboBoxWindows
        private void initComboBoxWindows(ComboBox targetComboBox)
        {
            targetComboBox.Text = "---Please Assign Parameters---";
            //Readings of position, rate, acceleration and their commends
            targetComboBox.Items.Add("Raw Position Feedback");
            targetComboBox.Items.Add("Estimated Position");
            targetComboBox.Items.Add("Estimated Velocity");
            targetComboBox.Items.Add("Filtered Velocity Estimate");
            targetComboBox.Items.Add("Filtered Accel Estimate");
            targetComboBox.Items.Add("Profiler Position Commend");
            targetComboBox.Items.Add("Profiler Velocity Commend");
            targetComboBox.Items.Add("Profiler Accel Commend");
            //Limitations
            targetComboBox.Items.Add("Maximun Position Limit");
            targetComboBox.Items.Add("Minimun Position Limit");
            targetComboBox.Items.Add("Position Mode Velocity Limit");
            targetComboBox.Items.Add("Position Mode Accel Limit");
            targetComboBox.Items.Add("Rate Mode Velocity Limit");
            targetComboBox.Items.Add("Rate Mode Accel Limit");
            targetComboBox.Items.Add("Synthesis Mode Velocity Limit");
            targetComboBox.Items.Add("Synthesis Mode Accel Limit");
            targetComboBox.Items.Add("Abort Mode Velocity Limit");
            targetComboBox.Items.Add("Abort Mode Accel Limit");
            targetComboBox.Items.Add("Velocity Absolute Limit");
            targetComboBox.Items.Add("Accel Absolute Limit");
            targetComboBox.Items.Add("Rate Trip Limit");
        }
        #endregion

        private void comboBoxSelectMode_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxSelectMode.SelectedItem == "Position Mode")
            {
                SelectMode(PositionMode);
            }
            else if (comboBoxSelectMode.SelectedItem == "Relative Rate Mode")
            {
                SelectMode(RelativeRateMode);
            }
            else if (comboBoxSelectMode.SelectedItem == "Absolute Rate Mode")
            {
                SelectMode(AbsRateMode);
            }
            else if (comboBoxSelectMode.SelectedItem == "Synthesis Mode")
            {
                SelectMode(SynthesisMode);
            }
        }

        private void SelectMode(String chosenMode)
        {
            try
            {
                mbSession.Write(":M:"+chosenMode+" 1 \n");
            }

            catch (VisaException v_exp)
            {
                MessageBox.Show(v_exp.Message);
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message);
            }
        }


        private void openMySession()
        {
            closeMySession();

            //open a Session to the VNA
            mySession = ResourceManager.GetLocalManager().Open(sAddress);
            //cast this to a message based session
            mbSession = (MessageBasedSession)mySession;
        }

        private void closeMySession()
        {
            if (mbSession == null)
                return;

            // Toggle the hardware GPIB REN line. Return to Local.
            GpibSession gpib = (GpibSession)mySession;
            gpib.ControlRen(RenMode.DeassertAfterGtl);

            //Close the Session
            mbSession.Dispose();
            mbSession = null;
        }

        //Show (Read) the position, rate and acceleration of specific channel chosen
        private void ShowFrameAxis()
        {
            string responseString = null;

            responseString = ReadParameter(responseString, Position, textReadPos);
            posReadValue = Convert.ToDouble(responseString);

            //record position into txt file
            if (RecordCtr == true)
            {
                streamWriter.Write(responseString + ' ');//write position string
            }
            PosValArray[counter] = posReadValue;
            
            responseString = ReadParameter(responseString, Rate, textReadRate);
            responseString = ReadParameter(responseString, Acceleration, textReadAcc);

            counter++;
            DisplayData();

            //display cycle counts
            textBox_cycle_counter.Text = cycleCounter.ToString();
            /*Read Ovariables*/
            //responseString = ReadOvarParameter(responseString, OvarMinPosLim, textBoxPosLimLow);
            //responseString = ReadOvarParameter(responseString, OvarMaxPosLim, textBoxPosLimHigh);
            //responseString = ReadOvarParameter(responseString, OvarSynModeVeloLim, textBoxSynModeRateLim);
            //responseString = ReadOvarParameter(responseString, OvarSynModeAccLim, textBoxSynModeAccLim);
            //responseString = ReadOvarParameter(responseString, OvarPosModeVeloLim, textBoxPosModeRateLim); //Allows only 5 data selects per ECP available for overiable config
            //responseString = ReadOvarParameter(responseString, OvarPosModeAccLim, textBoxPosModeAccLim); 
            //responseString = ReadOvarParameter(responseString, OvarRateModeVeloLim, textBoxRateModeRateLim); 
            //responseString = ReadOvarParameter(responseString, OvarRateModeAccLim, textBoxRateModeAccLim); 
        }

        //display position in waveform format in real time
        private void DisplayData()
        {
            //Displaying Chair Movement Data
            this.pos_chart.Series["PosVal"].Points.Clear();

            if (counter >= DispLength)
            {
                for (disp = counter - (DispLength); disp < counter; disp++)
                {
                    this.pos_chart.Series["PosVal"].Points.AddXY((disp / 10).ToString(), PosValArray[disp]);
                }
            }
            else
            {
                for (disp = 0; disp < DispLength; disp++)
                {
                    this.pos_chart.Series["PosVal"].Points.AddXY((disp / 10).ToString(), PosValArray[disp]);
                }
            }

            //Displaying Eye Coil Data
            if (showCoilWaveform == true)
            {
                DisplayCoilWaveform();
            }

        }

        private void DisplayCoilWaveform()
        {
            //Displaying Eye Coil Data
            this.coil_chart.Series["A3XChannel8"].Points.Clear();
            this.coil_chart.Series["A3YChannel9"].Points.Clear();
            this.coil_chart.Series["A3ZChannel10"].Points.Clear();
            this.coil_chart_A4.Series["A4XChannel11"].Points.Clear();
            this.coil_chart_A4.Series["A4YChannel12"].Points.Clear();
            this.coil_chart_A4.Series["A4ZChannel13"].Points.Clear();

            //For displaying with 1% representative values in real time
            if (coilValArrayReduced != null)
            {
                if (coilValueCounter >= DispLength)
                {
                    for (disp = coilValueCounter - (DispLength); disp < coilValueCounter; disp++)//take representitive sample 1 out of 100 to ensure smooth display
                    {
                        //use disp/10 so that the unit shown in chart is 1 sec. (Since 1% representative, which is 10 samples per second)
                        this.coil_chart.Series["A3XChannel8"].Points.AddXY((disp / 10).ToString(), coilValArrayReduced[8, disp]);//coilValList[8][disp]);
                        this.coil_chart.Series["A3YChannel9"].Points.AddXY((disp / 10).ToString(), coilValArrayReduced[9, disp]); 
                        this.coil_chart.Series["A3ZChannel10"].Points.AddXY((disp / 10).ToString(), coilValArrayReduced[10, disp]); 
                        this.coil_chart_A4.Series["A4XChannel11"].Points.AddXY((disp / 10).ToString(), coilValArrayReduced[11, disp]);
                        this.coil_chart_A4.Series["A4YChannel12"].Points.AddXY((disp / 10).ToString(), coilValArrayReduced[12, disp]);
                        this.coil_chart_A4.Series["A4ZChannel13"].Points.AddXY((disp / 10).ToString(), coilValArrayReduced[13, disp]);
                    }
                }
                else
                {
                    for (disp = 0; disp < DispLength; disp++)
                    {
                        this.coil_chart.Series["A3XChannel8"].Points.AddXY((disp / 10).ToString(), coilValArrayReduced[8, disp]);//coilValList[8][disp]);
                        this.coil_chart.Series["A3YChannel9"].Points.AddXY((disp / 10).ToString(), coilValArrayReduced[9, disp]);
                        this.coil_chart.Series["A3ZChannel10"].Points.AddXY((disp / 10).ToString(), coilValArrayReduced[10, disp]);
                        this.coil_chart_A4.Series["A4XChannel11"].Points.AddXY((disp / 10).ToString(), coilValArrayReduced[11, disp]);
                        this.coil_chart_A4.Series["A4YChannel12"].Points.AddXY((disp / 10).ToString(), coilValArrayReduced[12, disp]);
                        this.coil_chart_A4.Series["A4ZChannel13"].Points.AddXY((disp / 10).ToString(), coilValArrayReduced[13, disp]);
                    }
                }
                
            }

            ////For Displaying all values
            //if (showCoilWaveform == true)
            //{
            //    if (coilValArray != null)
            //    {
            //        if (coilValueCounter * samplesPerChannel >= 30000)
            //        {
            //            for (disp = coilValueCounter * samplesPerChannel - (30000); disp < coilValueCounter * samplesPerChannel; disp += 100)//take representitive sample 1 out of 100 to ensure smooth display
            //            {
            //                this.coil_chart.Series["A3XChannel8"].Points.AddXY((disp / 10).ToString(), coilValArray[8, disp]);//coilValList[8][disp]);
            //            }
            //        }
            //        else
            //        {
            //            for (disp = 0; disp < 30000; disp += 100)
            //            {
            //                this.coil_chart.Series["A3XChannel8"].Points.AddXY((disp / 10).ToString(), coilValArray[8, disp]);//coilValList[8][disp]);
            //            }
            //        }
            //    }
            //}
        }

        private string ReadParameter(string responseString, string targetParameter, TextBox targetTextBox)
        {
            try
            {
                //Send target Query command
                mbSession.Write(":u:f f ; :R:"+targetParameter+" 1 \n");
                //Read the response
                responseString = mbSession.ReadString();
                targetTextBox.Text = responseString;
            }
            catch (VisaException v_exp)
            {
                MessageBox.Show(v_exp.Message);
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message);
            }
            return responseString;
        }

        private string ReadOvarParameter(string responseString, string Ovariable, TextBox targetTextBox)
        {
            try
            {
                //Send target Query command
                mbSession.Write(":u:f f ; :r:o " + Ovariable + " \n");
                //Read the response
                responseString = mbSession.ReadString();
                targetTextBox.Text = responseString;
            }
            catch (VisaException v_exp)
            {
                MessageBox.Show(v_exp.Message);
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message);
            }
            return responseString;
        }

        private void ShowAxis_Tick(object sender, EventArgs e)
        {
            ShowFrameAxis();
        }

        private void SetAxisParameters()
        {
            CommendParameter(Position, textBoxSetPos);
            CommendParameter(Rate, textBoxSetRate);
            CommendParameter(Acceleration, textBoxSetAcc);
        }


        private void CommendParameter(string targetParameter, TextBox targetTextBox)
        {
            try
            {
                //Set target command
                mbSession.Write(":D:" + targetParameter + " 1, " + targetTextBox.Text.ToString() + " \n");
            }
            catch (VisaException v_exp)
            {
                MessageBox.Show(v_exp.Message);
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message);
            }
        }

        private void ExecuteCommendButton_Click(object sender, EventArgs e)
        {
            //make variables for legality checking
            double rateLimit, PosLimitLow, PosLimitHigh, accLimit;
            double inputPos, inputRate, inputAcc;
            rateLimit = Convert.ToDouble(textBoxLimitRate.Text);
            PosLimitLow = Convert.ToDouble(textBoxLimitPosL.Text);
            PosLimitHigh = Convert.ToDouble(textBoxLimitPosH.Text);
            accLimit = Convert.ToDouble(textBoxLimitAcc.Text);
            if (textBoxSetPos.Text == "")
            {
                inputPos = 0;
            }
            else
            {
                inputPos = Convert.ToDouble(textBoxSetPos.Text);
            }
            if (textBoxSetRate.Text == "")
            {
                inputRate = 0;
            }
            else
            {
                inputRate = Convert.ToDouble(textBoxSetRate.Text);
            }
            if (textBoxSetAcc.Text == "")
            {
                inputAcc = 0;
            }
            else
            {
                inputAcc = Convert.ToDouble(textBoxSetAcc.Text);
            }
            //check legality of input
            if (inputPos > PosLimitHigh || PosLimitHigh < PosLimitLow || inputRate > rateLimit || inputAcc > accLimit)
            {
                MessageBox.Show("Exceed Limits!");
            }
            else
            {
                SetAxisParameters();
            }  
        }

        private void button_simpleCommend_stop_Click(object sender, EventArgs e)
        {
            textBoxSetPos.Text = "0";//set 0 commend
            SelectMode(Position);//go to position mode 
            comboBoxSelectMode.Text = "Position Mode";
            updateLimits("20", "50");//set default limits
            CommendParameter(Position, textBoxSetPos);//execute commend
        }

        private void ReturnLocalButton_Click(object sender, EventArgs e)
        {

            try
            {
                closeMySession();
                ShowAxis.Enabled = false;
            }
            catch (VisaException v_exp)
            {
                MessageBox.Show(v_exp.Message);
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message);
            }

        }

        private void SinInputButton_Click(object sender, EventArgs e)
        {
            //check executability (legal input)
            double maxRate, maxAccel, rateLimit, AccLimit, inputMag, inputFreq;
            rateLimit = Convert.ToDouble(textBoxLimitRate.Text);
            AccLimit = Convert.ToDouble(textBoxLimitAcc.Text);
            inputMag = Convert.ToDouble(textBoxSetMagn.Text);
            inputFreq = Convert.ToDouble(textBoxSetFreq.Text);

            maxRate = Math.Round((inputMag * 2 * 3.14 * inputFreq), 2);
            maxAccel = Math.Round((inputMag * Math.Pow(2 * 3.14 * inputFreq, 2)), 2);

            if (maxRate > rateLimit || maxAccel > AccLimit)
            {
                MessageBox.Show("Limit exceeded! maxRate: " + maxRate.ToString() + " maxAccel: " + maxAccel.ToString());
            }
            else
            {
                CommendSinusoidal();
            }
             
        }

        private void button_sinusoidal_stop_Click(object sender, EventArgs e)
        {
            textBoxSetMagn.Text = "0";
            textBoxSetFreq.Text = "0";
            textBoxSetPhase.Text = "0";
            CommendSinusoidal();
        }

        private void CommendSinusoidal()
        {
            try
            {
                //Set Magnitude, Frequency and Phase for sinusoidal input (default channel 1)
                 mbSession.Write(":D:O 1, " + textBoxSetMagn.Text.ToString() + ", " + textBoxSetFreq.Text.ToString() + ", " + textBoxSetPhase.Text.ToString() + " \n");
            }
            catch (VisaException v_exp)
            {
                MessageBox.Show(v_exp.Message);
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message);
            }
        }

        private void RemoteMode_Click(object sender, EventArgs e)
        {
            openMySession();
            ShowAxis.Enabled = true;
        }

        private void ExecuteLimitButton_Click(object sender, EventArgs e)
        {
            /*The following code set the limitations for all modes with the same choen limited values*/
            SetAllLimits();

            /*The following code set the limitations under a specific mode, i.e Pos, Velo, Acc value under one of Pos, Rate, Syn modes*/
            //if (comboBoxSelectMode.SelectedItem == "Position Mode")
            //{
            //    SetLimitations(PositionMode);
            //}
            //else if (comboBoxSelectMode.SelectedItem == "Relative Rate Mode" || comboBoxSelectMode.SelectedItem == "Absolute Rate Mode")
            //{
            //    SetLimitations(RateMode);
            //}
            //else if (comboBoxSelectMode.SelectedItem == "Synthesis Mode")
            //{
            //    SetLimitations(SynthesisMode);
            //}
           
        }

        private void SetAllLimits()
        {
            mbSession.Write(":L :L " + PositionMode + " 1 " + textBoxLimitPosL.Text.ToString() + "; :L :H " + PositionMode + " 1 " + textBoxLimitPosH.Text.ToString() + " \n");
            mbSession.Write(":L :R " + PositionMode + " 1 " + textBoxLimitRate.Text.ToString() + " \n");
            mbSession.Write(":L :A " + PositionMode + " 1 " + textBoxLimitAcc.Text.ToString() + " \n");

            mbSession.Write(":L :L " + RateMode + " 1 " + textBoxLimitPosL.Text.ToString() + "; :L :H " + RateMode + " 1 " + textBoxLimitPosH.Text.ToString() + " \n");
            mbSession.Write(":L :R " + RateMode + " 1 " + textBoxLimitRate.Text.ToString() + " \n");
            mbSession.Write(":L :A " + RateMode + " 1 " + textBoxLimitAcc.Text.ToString() + " \n");

            mbSession.Write(":L :L " + SynthesisMode + " 1 " + textBoxLimitPosL.Text.ToString() + "; :L :H " + SynthesisMode + " 1 " + textBoxLimitPosH.Text.ToString() + " \n");
            mbSession.Write(":L :R " + SynthesisMode + " 1 " + textBoxLimitRate.Text.ToString() + " \n");
            mbSession.Write(":L :A " + SynthesisMode + " 1 " + textBoxLimitAcc.Text.ToString() + " \n");

            mbSession.Write(":L :R " + AbortMode + " 1 " + textBoxAbortRateLmt.Text.ToString() + " \n");
            mbSession.Write(":L :A " + AbortMode + " 1 " + textBoxAbortAccLmt.Text.ToString() + " \n");

            mbSession.Write(":L :V " + PositionMode + " 1 " + textBoxVTripLmt.Text.ToString() + " \n");
            mbSession.Write(":L :V " + RateMode + " 1 " + textBoxVTripLmt.Text.ToString() + " \n");
            mbSession.Write(":L :V " + SynthesisMode + " 1 " + textBoxVTripLmt.Text.ToString() + " \n");

        }

        /*The following code set the limitations under a specific mode, i.e Pos, Velo, Acc value under one of Pos, Rate, Syn modes*/
        //private void SetLimitations(String selectedMode)
        //{
        //    try
        //    {
        //        //Set Limitations for different modes 
        //        mbSession.Write(":L :L " + selectedMode + " 1 " + textBoxLimitPosL.Text.ToString() + "; :L :H " + selectedMode + " 1 " + textBoxLimitPosH.Text.ToString() + " \n");
        //        mbSession.Write(":L :R " + selectedMode + " 1 " + textBoxLimitRate.Text.ToString() + " \n");
        //        mbSession.Write(":L :A " + selectedMode + " 1 " + textBoxLimitAcc.Text.ToString() + " \n");
        //    }
        //    catch (VisaException v_exp)
        //    {
        //        MessageBox.Show(v_exp.Message);
        //    }
        //    catch (Exception exp)
        //    {
        //        MessageBox.Show(exp.Message);
        //    }
        //}

        #region WindowDisplays
        private void comboBox_window1_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetupWindows("1", comboBox_window1);
        }

        private void comboBox_window2_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetupWindows("2", comboBox_window2);
        }

        private void comboBox_window3_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetupWindows("3", comboBox_window3);
        }

        private void comboBox_window4_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetupWindows("4", comboBox_window4);
        }

        private void comboBox_window5_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetupWindows("5", comboBox_window5);
        }

        private void comboBox_window6_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetupWindows("6", comboBox_window6);
        }

        private void SetupWindows(String windowNum, ComboBox targetComboBox)
        {
            try
            {

                if (targetComboBox.SelectedItem == "Raw Position Feedback")
                {
                    mbSession.Write(":C:W " + windowNum + ", " + RawPositionFeedback + "  \n");
                }
                else if (targetComboBox.SelectedItem == "Estimated Position")
                {
                    mbSession.Write(":C:W " + windowNum + ", " + EstimatedPosition + "  \n");
                }
                else if (targetComboBox.SelectedItem == "Estimated Velocity")
                {
                    mbSession.Write(":C:W " + windowNum + ", " + EstimatedVelocity + "  \n");
                }
                else if (targetComboBox.SelectedItem == "Filtered Velocity Estimate")
                {
                    mbSession.Write(":C:W " + windowNum + ", " +FilteredVelocityEstimate + "  \n");
                }
                else if (targetComboBox.SelectedItem == "Filtered Accel Estimate")
                {
                    mbSession.Write(":C:W " + windowNum + ", " + FilteredAccelEstimate + "  \n");
                }
                else if (targetComboBox.SelectedItem == "Profiler Position Commend")
                {
                    mbSession.Write(":C:W " + windowNum + ", " + ProfilerPositionCommend + "  \n");
                }
                else if (targetComboBox.SelectedItem == "Profiler Velocity Commend")
                {
                    mbSession.Write(":C:W " + windowNum + ", " + ProfilerVelocityCommend + "  \n");
                }
                else if (targetComboBox.SelectedItem == "Profiler Accel Commend")
                {
                    mbSession.Write(":C:W " + windowNum + ", " + ProfilerAccelCommend + "  \n");
                }
                else if (targetComboBox.SelectedItem == "Maximun Position Limit")
                {
                    mbSession.Write(":C:W " + windowNum + ", " + MaximunPositionLimit + "  \n");
                }
                else if (targetComboBox.SelectedItem == "Minimun Position Limit")
                {
                    mbSession.Write(":C:W " + windowNum + ", " + MinimunPositionLimit + "  \n");
                }
                else if (targetComboBox.SelectedItem == "Position Mode Velocity Limit")
                {
                    mbSession.Write(":C:W " + windowNum + ", " + PositionModeVelocityLimit + "  \n");
                }
                else if (targetComboBox.SelectedItem == "Position Mode Accel Limit")
                {
                    mbSession.Write(":C:W " + windowNum + ", " + PositionModeAccelLimit + "  \n");
                }
                else if (targetComboBox.SelectedItem == "Rate Mode Velocity Limit")
                {
                    mbSession.Write(":C:W " + windowNum + ", " + RateModeVelocityLimit + "  \n");
                }
                else if (targetComboBox.SelectedItem == "Rate Mode Accel Limit")
                {
                    mbSession.Write(":C:W " + windowNum + ", " + RateModeAccelLimit + "  \n");
                }
                else if (targetComboBox.SelectedItem == "Synthesis Mode Velocity Limit")
                {
                    mbSession.Write(":C:W " + windowNum + ", " + SynthesisModeVelocityLimit + "  \n");
                }
                else if (targetComboBox.SelectedItem == "Synthesis Mode Accel Limit")
                {
                    mbSession.Write(":C:W " + windowNum + ", " + SynthesisModeAccelLimit + "  \n");
                }
                else if (targetComboBox.SelectedItem == "Abort Mode Velocity Limit")
                {
                    mbSession.Write(":C:W " + windowNum + ", " + AbortModeVelocityLimit + "  \n");
                }
                else if (targetComboBox.SelectedItem == "Abort Mode Accel Limit")
                {
                    mbSession.Write(":C:W " + windowNum + ", " + AbortModeAccelLimit +"  \n");
                }
                else if (targetComboBox.SelectedItem == "Velocity Absolute Limit")
                {
                    mbSession.Write(":C:W " + windowNum + ", " +VelocityAbsoluteLimit +"  \n");
                }
                else if (targetComboBox.SelectedItem == "Accel Absolute Limit")
                {
                    mbSession.Write(":C:W " + windowNum + ", " +AccelAbsoluteLimit +"  \n");
                }
                else if (targetComboBox.SelectedItem == "Rate Trip Limit")
                {
                    mbSession.Write(":C:W " + windowNum + ", " + RateTripLimit + "  \n");
                }

            }
            catch (VisaException v_exp)
            {
                MessageBox.Show(v_exp.Message);
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message);
            }
        }
        #endregion

        private void button_default_windows_Click(object sender, EventArgs e)
        {
            mbSession.Write(":C:W 1, " + EstimatedPosition + "  \n");
            mbSession.Write(":C:W 2, " + EstimatedVelocity + "  \n");
            mbSession.Write(":C:W 3, " + FilteredAccelEstimate + "  \n");
            mbSession.Write(":C:W 4, " + ProfilerPositionCommend + "  \n");
            mbSession.Write(":C:W 5, " + ProfilerVelocityCommend + "  \n");
            mbSession.Write(":C:W 6, " + ProfilerAccelCommend + "  \n");
        }

        private void button_showLimits_Click(object sender, EventArgs e)
        {
            mbSession.Write(":C:W 1, " + MinimunPositionLimit + "  \n");
            //mbSession.Write(":C:W 2, " + MaximunPositionLimit + "  \n");
            mbSession.Write(":C:W 2, " + RateTripLimit + "  \n");
            mbSession.Write(":C:W 3, " + SynthesisModeVelocityLimit + "  \n");
            mbSession.Write(":C:W 4, " + SynthesisModeAccelLimit + "  \n");
            //mbSession.Write(":C:W 5, " + AbortModeVelocityLimit + "  \n");
            mbSession.Write(":C:W 5, " + VelocityAbsoluteLimit + "  \n");
            mbSession.Write(":C:W 6, " + AbortModeAccelLimit + "  \n");
            //mbSession.Write(":C:W 5, " + PositionModeVelocityLimit + "  \n");
            //mbSession.Write(":C:W 6, " + PositionModeAccelLimit + "  \n");
            //mbSession.Write(":C:W 5, " + RateModeVelocityLimit + "  \n");
            //mbSession.Write(":C:W 6, " + RateModeAccelLimit + "  \n");
        }

        private void button_reset_limits_Click(object sender, EventArgs e)
        {
            textBoxLimitPosL.Clear();
            textBoxLimitPosH.Clear();
            textBoxLimitRate.Clear();
            textBoxLimitAcc.Clear();
        }

        private void edit_default_button_Click(object sender, EventArgs e)
        {

        }

        private void button_interlock_open_Click(object sender, EventArgs e)
        {
            Interlock_Open();
            RecordCtr = false;
        }

        private void Interlock_Open()
        {
            if (mySession == null)
                openMySession();

            mbSession.Write(":I:O 1  \n");
        }

        private void button_interlock_close_Click(object sender, EventArgs e)
        {
            Interlock_Close();
        }

        private void Interlock_Close()
        {
            if (mySession == null)
                openMySession();

            mbSession.Write(":I:C 1  \n");
        }

        private void button_interlock_reset_Click(object sender, EventArgs e)
        {
            if (mySession == null)
                openMySession();

            mbSession.Write(":I:R \n");
        }

        private void button_ECP4_Click(object sender, EventArgs e)
        {
            if (mySession == null)
                openMySession();

            //Setup ECP 4 using TOUCH commend
            mbSession.Write("u:t 179,179,176,32,148,32,50,176,181,181,181 \n");
            System.Threading.Thread.Sleep(5000);
        }

        private void button_ECP_Click(object sender, EventArgs e)
        {
            if (mySession == null)
                openMySession();
            ////Setup ECP 80 then 87 using TOUCH commend
            //mbSession.Write("u:t 179,179,176,32,152,144,32,50,176,181,176,32,152,151,32,50,176,181,181,181 \n");

            //Setup ECP 87 using TOUCH commend
            mbSession.Write("u:t 179,179,176,32,152,151,32,50,176,181,181,181 \n");
            System.Threading.Thread.Sleep(5000);
        }

        private void button_cut_analog_input_Click(object sender, EventArgs e)
        {
            if (mySession == null)
                openMySession();
            //Set analog inputs gain to 0 
            //System.Threading.Thread.Sleep(10000);//wait for 10s for ECP to be restored
            mbSession.Write(":u:t 180,177,178,50,32,144,32,176,181,179,50,32,144,32,176,181,181,181 \n");
        }

        private void buttonSinuSeqExecute_Click(object sender, EventArgs e)
        {
            ////update the sinosoidal parameter inputs
            //ReadSinosoidalParas();

            //calculate the max rate and accel values, and check if the sequences can be executed or not
            CalculateMax();

            if (isExecutable == true)
            {
                //open text writer
                WriteIntoTxt();

                //Set all limitations
                SetAllLimits();

                //Go back to zero position under position mode
                textBoxSetPos.Text = "0"; //need to set earlier before commend, it takes some time to write
                SelectMode(PositionMode);
                CommendParameter(Position, textBoxSetPos);
                Interlock_Close();

                CheckZeroPosition.Enabled = true; //Once go back to zero, start sinusoidal cycles

                cycleCounter = 0;

                //start displaying coil signal
                showCoilWaveform = true;
                //start coil signal recording
                StartCoilSignalRecording();
            }
        }

        private void WriteIntoTxt()
        {
            if (!Directory.Exists(savingPath))
            {
                Directory.CreateDirectory(savingPath);
            }
            string path = savingPath + "\\" + DateTime.Now.ToString("dd-MM-yyyy_hh-mm-ss") + ".txt";
            fileStream = new FileStream(path, System.IO.FileMode.Create);
            streamWriter = new StreamWriter(fileStream);
            RecordCtr = true;//start recording position into txt
        }

        private void updateLimits(String newRateLimit, String newAccLimit){
            textBoxLimitRate.Text = newRateLimit;
            textBoxLimitAcc.Text = newAccLimit;
            SetAllLimits(); 
        }

        private void CheckZeroPosition_Tick(object sender, EventArgs e)
        {
            ExecuteOneSequence(TargetMagn[seqCount], TargetFreq[seqCount]);
        }

        private void ExecuteOneSequence(double targetMagn, double targetFreq)
        {
            if (Math.Abs(posReadValue) < zeroTrigger)
            {
                SelectMode(SynthesisMode); comboBoxSelectMode.Text = "Synthesis Mode";
                textBoxSetMagn.Text = targetMagn.ToString();//need to set earlier before commmend, it takes some time to write
                textBoxSetFreq.Text = targetFreq.ToString();

                //set up the timer used for cycle counting based on target frequency
                sysTimer = new System.Timers.Timer(1000 / targetFreq); //so that the time interval = 1/f (sec)
                sysTimer.Elapsed += new System.Timers.ElapsedEventHandler(theCount);
                sysTimer.AutoReset = true; // false:execute only once. true: keep execute

                //updateLimits("150", "1200");
                updateLimits(textBox_desiredRateLimit.Text, textBox_desiredAccLimit.Text);
                System.Threading.Thread.Sleep(50);//wait for calculation and write
                CommendSinusoidal();

                CheckZeroPosition.Enabled = false;
                CheckCycleCount.Enabled = true;
                sysTimer.Enabled = true;
            }

        }

        private void theCount(object source, System.Timers.ElapsedEventArgs e)
        {
            cycleCounter++; //increased by one in every time interval = 1/f (sec)
        }

        private void CheckCycleCount_Tick(object sender, EventArgs e)
        {
            if (cycleCounter == TargetCycleCount[seqCount] - 5)//start slowing down 5 cycles before finishing
            {
                textBoxSetMagn.Text = "0";
                CommendSinusoidal();
            }
            if (cycleCounter == TargetCycleCount[seqCount])
            {
                //go back to zero position again
                textBoxSetPos.Text = "0";
                updateLimits("20", "50");
                SelectMode(PositionMode); comboBoxSelectMode.Text = "Position Mode";
                sysTimer.Enabled = false;
                CheckCycleCount.Enabled = false;
                CommendParameter(Position, textBoxSetPos);

                ReturnZeroPosition.Enabled = true;

                if (TargetCycleCount[seqCount] == 0)//not need to check all the rest sequences with 0 desired cycle numbers
                {
                    seqCount = totalSequencesNum;// go to the last sequency directly
                }
            }
        }

        private void ReturnZeroPosition_Tick(object sender, EventArgs e)
        {
            if (seqCount < totalSequencesNum && Math.Abs(posReadValue) < zeroTrigger)
            {
                seqCount++;
                cycleCounter = 0;
                CheckZeroPosition.Enabled = true;//go to the next sequence
                ReturnZeroPosition.Enabled = false;
            }

            if (seqCount == totalSequencesNum && Math.Abs(posReadValue) < zeroTrigger)
            {
                Interlock_Open();
                CheckZeroPosition.Enabled = false;//no need to go back
                ReturnZeroPosition.Enabled = false;

                //stop position recording and dispose
                RecordCtr = false;
                streamWriter.Close();
                fileStream.Close();
                seqCount = 0;

                //stop coil signal recording and write to file
                StopCoilSignalRecording();
                //stop displaying coil signals
                showCoilWaveform = false;
            }
        }

        private void button_select_file_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = savingPath;
            openFileDialog.Filter = "txt Files(*.txt)|*.txt|All Files(*.*)|*.*";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                readingPath = openFileDialog.FileName;
            }
            textBoxReadingDirectory.Text = readingPath;
            Directory.SetCurrentDirectory(AppDomain.CurrentDomain.BaseDirectory);
        }

        private void button_display_file_Click(object sender, EventArgs e)
        {
            ShowAxis.Enabled = false;
            if (textBoxReadingDirectory.Text != "")
            {
                readingPath = textBoxReadingDirectory.Text;
                Read_Txt();
            }
            else
                MessageBox.Show("Please input a path of data file!");
        }

        private void Read_Txt()
        {
            StreamReader sr = new StreamReader(readingPath, Encoding.Default);
            string wholeString;
            string[] readPosValArray;
            Array.Clear(PosValArray, 0, PosValArray.Length);//clear original PosValArray
            wholeString = sr.ReadToEnd();//generate a string containing all the info in the file
            readPosValArray = wholeString.Split(default(string[]), StringSplitOptions.RemoveEmptyEntries);//split the whole string into a string array
            PosValArray = Array.ConvertAll(readPosValArray, Double.Parse);//convert string[] to double[]
            sr.Close();
            DisplayFileData();
        }
        //display position in waveform format
        private void DisplayFileData()
        {
            this.pos_chart.Series["PosVal"].Points.Clear();
            for (disp = 0; disp < PosValArray.Length; disp++)
            {
                this.pos_chart.Series["PosVal"].Points.AddXY((disp / 10).ToString(), PosValArray[disp]);
            }
        }

        private void button_return_realtime_Click(object sender, EventArgs e)
        {
            this.pos_chart.Series["PosVal"].Points.Clear();
            ShowAxis.Enabled = true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        //start recording coil signals
        private void StartCoilSignalRecording()
        {
            if (runningTask == null)
            {
                try
                {
                    button_start_coil_record.Enabled = false;
                    button_stop_coil_record.Enabled = true;

                    //open a new .txt file for coil data recording
                    CreateDataFile();

                    //create a new task
                    myTask = new NationalInstruments.DAQmx.Task();

                    //create a virtual channel
                    myTask.AIChannels.CreateVoltageChannel("Dev2/ai0:15", "", AITerminalConfiguration.Rse, -10, 10, AIVoltageUnits.Volts);//easier way:  ai0:n (0:15 in total)

                    //configure the timing parameters
                    myTask.Timing.ConfigureSampleClock("", timingRate, SampleClockActiveEdge.Rising, SampleQuantityMode.ContinuousSamples, samplesPerChannel);

                    //verify the Task
                    myTask.Control(TaskAction.Verify);

                    //prepare file for data
                    String[] channelNames = new String[myTask.AIChannels.Count];
                    int i = 0;
                    foreach (AIChannel a in myTask.AIChannels)
                    {
                        channelNames[i++] = a.PhysicalName;
                    }

                    ////prepare the table for data
                    //InitializeDataTable(channelNames, ref dataTable);
                    //acquisitionDataGrid.DataSource = dataTable;//show dataTable on DataGridView

                    //Add the channel names and any other information to the file
                    PrepareFileForData();
                    savedCoilData = new ArrayList();
                    for (i = 0; i < myTask.AIChannels.Count; i++)
                    {
                        savedCoilData.Add(new ArrayList());
                    }

                    runningTask = myTask;
                    analogInReader = new AnalogMultiChannelReader(myTask.Stream);
                    analogCallback = new AsyncCallback(AnalogInCallback);

                    //use SynchronizeCallbacks to specify that the object marshals callbacks across threads appropriately
                    analogInReader.SynchronizeCallbacks = true;
                    analogInReader.BeginReadMultiSample(samplesPerChannel, analogCallback, myTask);
                }

                catch (DaqException exception)
                {
                    //display errors
                    MessageBox.Show(exception.Message);
                    runningTask = null;
                    myTask.Dispose();
                }

            }
        }

        private void CreateDataFile()
        {
            try
            {
                if (!Directory.Exists(coilSavingPath))
                {
                    Directory.CreateDirectory(coilSavingPath);
                }
                string coil_path = coilSavingPath + "\\" + DateTime.Now.ToString("dd-MM-yyyy_hh-mm-ss") + ".txt";
                coilFileStream = new FileStream(coil_path, System.IO.FileMode.Create);
                fileStreamWriter = new StreamWriter(coilFileStream);
                CoilRecordCtr = true;//start recording coil position into txt
            }
            catch(System.IO.IOException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //Used by text files to write the channel name
        private void PrepareFileForData()
        {
            //prepare file for data: write the channel names
            int numChannels = myTask.AIChannels.Count;

            for (int i = 0; i < numChannels; i++)
            {
                fileStreamWriter.Write(myTask.AIChannels[i].PhysicalName);
                fileStreamWriter.Write("\t");
            }
            fileStreamWriter.WriteLine();
        }
        ////Clean up any resources being used
        //protected override void DisposeTask(bool disposing)
        //{
        //    if (disposing)
        //    {
        //        components.Dispose();
        //    }
        //    if (myTask != null)
        //    {
        //        runningTask = null;
        //        myTask.Dispose();
        //    }
        //    base.Dispose(disposing);
        //}

        private void AnalogInCallback(IAsyncResult ar)
        {
            try
            {
                if (runningTask != null && runningTask == ar.AsyncState)
                {
                    //Read the available data from the channels
                    data = analogInReader.EndReadMultiSample(ar);//get raw data from analogInReader

                    //Displays data in grid and writes to file
                    //DisplayData(data, ref dataTable);

                    //Copy and append coil data value to coilValArray for display purpose
                    for (int i = 0; i < 16; i++)//16 channels
                    {
                        //save all the samples to coilValArray
                        //for (int j = 0; j < samplesPerChannel; j++)//samplesPerChannel=100
                        //{
                        //    coilValArray[i, samplesPerChannel * coilValueCounter + j] = data[i, j];
                        //}

                        //take 1 representative sample out of 100 samples to be saved in coilValArray for display purpose on graph
                        for (int j = 0; j < samplesPerChannel; j += 100)
                        {
                            coilValArrayReduced[i, coilValueCounter + j] = data[i, j];
                        }
                    }
                    coilValueCounter++;

                    LogData(data);

                    analogInReader.BeginReadMultiSample(samplesPerChannel,
                        analogCallback, myTask);
                }
            }
            catch (DaqException exception)
            {
                // Display Errors
                MessageBox.Show(exception.Message);
                runningTask = null;
                myTask.Dispose();
            }
        }

        //private void DisplayData(double[,] sourceArray, ref DataTable dataTable)
        //{
        //    //Display data in the DataGrid
        //    try
        //    {
        //        int channelCount = sourceArray.GetLength(0);
        //        int dataCount;

        //        //if (sourceArray.GetLength(1) < DisplayedRowNum)
        //        //    dataCount = sourceArray.GetLength(1);
        //        //else
        //        dataCount = DisplayedRowNum;

        //        //write to DataTable
        //        for (int i = 0; i < dataCount; i++)
        //        {
        //            for (int j = 0; j < channelCount; j++)
        //            {
        //                // Writes data to data table
        //                dataTable.Rows[i][j] = sourceArray.GetValue(j, i);
        //            }
        //        }
        //    }

        //    catch (Exception e)
        //    {
        //        MessageBox.Show(e.ToString());
        //        runningTask = null;
        //        myTask.Dispose();
        //        button_start_coil_record.Enabled = true;
        //        button_stop_coil_record.Enabled = false;
        //    }
        //}

        private void LogData(double[,] data)
        {
            int channelCount = data.GetLength(0);
            int dataCount = data.GetLength(1);

            for (int i = 0; i < channelCount; i++)
            {
                ArrayList l = savedCoilData[i] as ArrayList;
                for (int j = 0; j < dataCount; j++)
                {
                    l.Add(data[i, j]);
                }
            }
        }

        private void StopCoilSignalRecording()
        {
            if (runningTask != null)
            {
                //dispose of the task
                CloseFile();

                // Dispose of the task
                runningTask = null;
                myTask.Dispose();
            }
        }

        private void CloseFile()
        {
            int channelCount = savedCoilData.Count;
            int dataCount = (savedCoilData[0] as ArrayList).Count;

            try
            {
                fileStreamWriter.WriteLine(dataCount.ToString());//write the num of rows in txt file

                for (int i = 0; i < dataCount; i++)
                {
                    for (int j = 0; j < channelCount; j++)
                    {
                        // Writes data to file
                        ArrayList l = savedCoilData[j] as ArrayList;
                        double dataValue = (double)l[i];
                        fileStreamWriter.Write(dataValue.ToString("e2"));//round to 0.01
                        fileStreamWriter.Write("\t"); //seperate the data for each channel
                    }
                    fileStreamWriter.WriteLine(); //new line of data (start next scan)
                }
                fileStreamWriter.Close();

                button_start_coil_record.Enabled = true;
                button_stop_coil_record.Enabled = false;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.TargetSite.ToString());
                runningTask = null;
                myTask.Dispose();
                button_start_coil_record.Enabled = true;
                button_stop_coil_record.Enabled = false;
            }
        }

        //public void InitializeDataTable(String[] channelNames, ref DataTable data)
        //{
        //    int numOfChannels = channelNames.GetLength(0);
        //    data.Rows.Clear();
        //    data.Columns.Clear();
        //    dataColumn = new DataColumn[numOfChannels];
        //    int numOfRows = DisplayedRowNum;

        //    for (int currentChannelIndex = 0; currentChannelIndex < numOfChannels; currentChannelIndex++)
        //    {
        //        dataColumn[currentChannelIndex] = new DataColumn();
        //        dataColumn[currentChannelIndex].DataType = typeof(double);
        //        dataColumn[currentChannelIndex].ColumnName = channelNames[currentChannelIndex];
        //    }

        //    data.Columns.AddRange(dataColumn);

        //    for (int currentDataIndex = 0; currentDataIndex < numOfRows; currentDataIndex++)
        //    {
        //        object[] rowArr = new object[numOfChannels];
        //        data.Rows.Add(rowArr);
        //    }
        //}

        private void button_select_coil_file_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = coilSavingPath;
            openFileDialog.Filter = "txt Files(*.txt)|*.txt|All Files(*.*)|*.*";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                coilReadingPath = openFileDialog.FileName;
            }
            textBoxReadingCoilDirectory.Text = coilReadingPath;
            Directory.SetCurrentDirectory(AppDomain.CurrentDomain.BaseDirectory);
        }

        private void button_display_coil_file_Click(object sender, EventArgs e)
        {
            //modify UI
            button_display_coil_file.Enabled = false;
            this.Cursor = Cursors.WaitCursor;

            //open file
            bool opened = OpenDataFile();

            //load data
            ReadTextData();
            fileStreamReader.Close();

            this.Cursor = Cursors.Default;
            button_display_coil_file.Enabled = true;
        }

        private bool OpenDataFile()
        {
            try
            {
                FileStream fs = new FileStream(coilReadingPath, FileMode.Open);

                fileStreamReader = new StreamReader(fs);

            }
            catch (System.IO.IOException ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
            return true;
        }

        private void ReadTextData()
        {
            try
            {
                char[] tab = { '\t' };
                String[] split = fileStreamReader.ReadLine().Replace("\n", "").Split(tab);
                String[] channels = new String[split.GetLength(0) - 1];
                Array.Copy(split, 0, channels, 0, split.GetLength(0) - 1);
                int samples = Int32.Parse(fileStreamReader.ReadLine().Replace("\n", ""));
                int channelCount = channels.GetLength(0);

                //if (samples > 10000)
                //    DisplayedRowNum = 10000;
                //else 
                //    DisplayedRowNum = samples;//display all the samples

                double[,] array = new double[channelCount, samples];

                String line;
                for (int iSample = 0; iSample < samples; iSample++)
                {
                    line = fileStreamReader.ReadLine();
                    String[] values = line.Split(tab);

                    for (int iChan = 0; iChan < channelCount; iChan++)
                    {
                        array[iChan, iSample] = Convert.ToDouble(values[iChan]);
                    }
                }

                //InitializeDataTable(channels, ref dataTable);
                //acquisitionDataGrid.DataSource = dataTable;

                //DisplayData(array, ref dataTable);//show in table

                DisplayCoilFileData(array);//show in graph

            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
                runningTask = null;
                button_display_coil_file.Enabled = true;
            }
        }

        //display coil signals in waveform format
        private void DisplayCoilFileData(double[,] array)
        {

            this.coil_chart.Series["A3XChannel8"].Points.Clear();
            this.coil_chart.Series["A3YChannel9"].Points.Clear();
            this.coil_chart.Series["A3ZChannel10"].Points.Clear();
            this.coil_chart_A4.Series["A4XChannel11"].Points.Clear();
            this.coil_chart_A4.Series["A4YChannel12"].Points.Clear();
            this.coil_chart_A4.Series["A4ZChannel13"].Points.Clear();
            int length = array.GetLength(1);
            for (disp = 0; disp < length; disp++)
            {
                this.coil_chart.Series["A3XChannel8"].Points.AddXY(disp, array[8, disp]); //showing all points (1000 samples per second)
                this.coil_chart.Series["A3YChannel9"].Points.AddXY(disp, array[9, disp]); //showing all points (1000 samples per second)
                this.coil_chart.Series["A3ZChannel10"].Points.AddXY(disp, array[10, disp]); //showing all points (1000 samples per second)
                this.coil_chart_A4.Series["A4XChannel11"].Points.AddXY(disp, array[11, disp]);
                this.coil_chart_A4.Series["A4YChannel12"].Points.AddXY(disp, array[12, disp]);
                this.coil_chart_A4.Series["A4ZChannel13"].Points.AddXY(disp, array[13, disp]);
            }
        }

        private void button_start_coil_record_Click(object sender, EventArgs e)
        {
            showCoilWaveform = true;
            StartCoilSignalRecording();
        }

        private void button_stop_coil_record_Click(object sender, EventArgs e)
        {
            showCoilWaveform = false;
            StopCoilSignalRecording();
        }

        private void button_initialization_Click(object sender, EventArgs e)
        {
            if (mySession == null)
                openMySession();
            //repeat 4-87-4-87 twice to fix the ECP bug
            //Setup ECP 4 using TOUCH commend
            mbSession.Write("u:t 179,179,176,32,148,32,50,176,181,181,181 \n");
            System.Threading.Thread.Sleep(5000);

            //Setup ECP 87 using TOUCH commend
            mbSession.Write("u:t 179,179,176,32,152,151,32,50,176,181,181,181 \n");
            System.Threading.Thread.Sleep(5000);

            //Setup ECP 4 using TOUCH commend
            mbSession.Write("u:t 179,179,176,32,148,32,50,176,181,181,181 \n");
            System.Threading.Thread.Sleep(5000);

            //Setup ECP 87 using TOUCH commend
            mbSession.Write("u:t 179,179,176,32,152,151,32,50,176,181,181,181 \n");
            System.Threading.Thread.Sleep(5000);

            //the following two steps fix the display bug
            //return local mode
            closeMySession();
            System.Threading.Thread.Sleep(1000);
            //go back to remote mode
            openMySession();
        }

        private void button_calculate_Click(object sender, EventArgs e)
        {
            CalculateMax();
        }

        private void CalculateMax()
        {
            ReadSinosoidalParas();
            double maxRate, maxAccel;
            isExecutable = true;
            for (int i = 0; i < 16; i++)
            {
                maxRate = Math.Round((TargetMagn[i] * 2 * 3.14 * TargetFreq[i]), 2);
                maxAccel = Math.Round((TargetMagn[i] * Math.Pow(2 * 3.14 * TargetFreq[i], 2)), 2);
                textBox_maxRate[i].Text = maxRate.ToString();
                textBox_maxAccel[i].Text = maxAccel.ToString();

                if (maxRate > Convert.ToDouble(textBox_desiredRateLimit.Text) || maxAccel > Convert.ToDouble(textBox_desiredAccLimit.Text))
                {
                    MessageBox.Show("Limit exceeded for the No. " + (i + 1) + " sequence!");
                    isExecutable = false;
                }
            }
        }

        private void button_saveSequenceFile_Click(object sender, EventArgs e)
        {
            // Create/Open file
            if (textBox_sequenceFileName.Text == "")
            {
                MessageBox.Show("Please Input File Name");
            }
            else
            {
                try
                {
                    if (!Directory.Exists(sequencesSavingPath))
                    {
                        Directory.CreateDirectory(sequencesSavingPath);
                    }
                    string sequenceFileName = sequencesSavingPath + "\\" + textBox_sequenceFileName.Text + ".txt";
                    sequenceFileStream = new FileStream(sequenceFileName, System.IO.FileMode.Create);
                    sequenceFileStreamWriter = new StreamWriter(sequenceFileStream);
                }
                catch (System.IO.IOException ex)
                {
                    MessageBox.Show(ex.Message);
                }

                //write to file
                for (int i = 0; i < totalSequencesNum; i++)
                {
                    sequenceFileStreamWriter.Write(textBoxSeqMag[i].Text + "\t");
                    sequenceFileStreamWriter.Write(textBoxSeqFreq[i].Text + "\t");
                    sequenceFileStreamWriter.Write(textBoxSeqCycle[i].Text + "\n");
                }
                sequenceFileStreamWriter.WriteLine();

                sequenceFileStreamWriter.Close();
                sequenceFileStream.Close();

                MessageBox.Show("Sequences File Saved");
            }      
        }

        private void button_loadSequenceFile_Click(object sender, EventArgs e)
        {
            // open file dialog
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = sequencesSavingPath;
            openFileDialog.Filter = "txt Files(*.txt)|*.txt|All Files(*.*)|*.*";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                sequenceReadingPath = openFileDialog.FileName;
            }

            Directory.SetCurrentDirectory(AppDomain.CurrentDomain.BaseDirectory);

            //read the .txt file and put into array
            sequenceFileStreamReader = new StreamReader(sequenceReadingPath, Encoding.Default);
            string wholeString;
            string[] ValArray;
            wholeString = sequenceFileStreamReader.ReadToEnd();//generate a string containing all the info in the file
            ValArray = wholeString.Split(default(string[]), StringSplitOptions.RemoveEmptyEntries);//split the whole string into a string array
            sequenceFileStreamReader.Close();

            //update the text boxes of all sequences
            for (int i = 0; i < totalSequencesNum; i++)
            {
                textBoxSeqMag[i].Text = ValArray[3*i];
                textBoxSeqFreq[i].Text = ValArray[3 * i + 1];
                textBoxSeqCycle[i].Text = ValArray[3 * i + 2];
            }

        }

        private void button_pulse_Execute_Click(object sender, EventArgs e)
        {
            targetCycleNum = Convert.ToInt16(textBox_cycle_num.Text);
            currentCycleCount = 0;
            pulseDelay = Convert.ToDouble(textBox_pulse_delay.Text);

            //open text writer for chair position
            WriteIntoTxt();

            //start coil recording
            showCoilWaveform = true;
            StartCoilSignalRecording();

            //set up the timer used for delay between pulses
            sysTimer = new System.Timers.Timer(1000 * pulseDelay); //so that the time interval = pulseDelay
            sysTimer.Elapsed += new System.Timers.ElapsedEventHandler(NextPulseAction);
            sysTimer.AutoReset = false;//true; // false:execute only once. true: keep execute

            Interlock_Close();
            //Go back to zero position
            textBoxSetPos.Text = "0";//set 0 commend
            SelectMode(Position);//go to position mode 
            comboBoxSelectMode.Text = "Position Mode";
            updateLimits("20", "50");//set default limits
            CommendParameter(Position, textBoxSetPos);//execute commend

            CheckPulseZeroPosition.Enabled = true;
        }

        private void NextPulseAction(object source, System.Timers.ElapsedEventArgs e)
        {
            if (DepartOrReturn == "Depart")
            {
                //Go to Target Position
                textBoxSetPos.Text = textBox_pulse_travelDist.Text;
                currentCycleCount++;
                textBox_currentCycleCounter.Text = currentCycleCount.ToString();
                updateLimits(textBox_pulse_rateLimit.Text, textBox_pulse_accelLimit.Text);//update limits

                //System.Threading.Thread.Sleep(50);//wait for calculation and write
                CommendParameter(Position, textBoxSetPos);//execute commend
                //CheckPulseZeroPosition.Enabled = false;
                CheckPulseTargetPosition.Enabled = true;
                //System.Threading.Thread.Sleep(1000);//wait for 1s
                sysTimer.Enabled = false;
            }
            else if (DepartOrReturn == "Return")
            {
                //Go back to Zero Position
                textBoxSetPos.Text = "0";
                //System.Threading.Thread.Sleep(50);//wait for calculation and write
                CommendParameter(Position, textBoxSetPos);//execute commend
                CheckPulseZeroPosition.Enabled = true;
                //CheckPulseTargetPosition.Enabled = false;
                //System.Threading.Thread.Sleep(1000);//wait for 1s

                //Stop Criteria (meet the number of total cycles)
                if (currentCycleCount == targetCycleNum)
                {
                    CheckPulseZeroPosition.Enabled = false;

                    Interlock_Open();

                    //stop position recording and dispose
                    RecordCtr = false;
                    streamWriter.Close();
                    fileStream.Close();
                    seqCount = 0;

                    //stop coil recording
                    showCoilWaveform = false;
                    StopCoilSignalRecording();
                }

                sysTimer.Enabled = false;
            }
        }

        private void CheckPulseZeroPosition_Tick(object sender, EventArgs e)
        {
            if (Math.Abs(posReadValue) < zeroTrigger)
            {
                DepartOrReturn = "Depart";
                sysTimer.Enabled = true;
                CheckPulseZeroPosition.Enabled = false;
            }
        }

        private void CheckPulseTargetPosition_Tick(object sender, EventArgs e)
        {
            double targetPosVal = Convert.ToDouble(textBox_pulse_travelDist.Text);
            if (Math.Abs(posReadValue-targetPosVal) < zeroTrigger)
            {
                DepartOrReturn = "Return";
                sysTimer.Enabled = true;
                CheckPulseTargetPosition.Enabled = false;
            }
        }

    }
}




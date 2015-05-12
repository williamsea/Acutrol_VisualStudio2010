using System;
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
     * 1. ECP87
     * 2. Counter (return real time - sometimes causing errors)
     * 
     * Additional:
     * 1.setup abs limits (1140,1141) and ROT/LIN.
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
        double posReadValue = 1000;//used to check if position has been set back to zero position
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


            textBoxSeqMag[0] = (TextBox)textBoxSeqMag1;
        }

        private void ReadSinosoidalParas()
        {
            //The parameters of Sinusoidal Inputs
            ////Magnitude
            TargetMagn[0] = Convert.ToDouble(textBoxSeqMag1.Text);
            TargetMagn[1] = Convert.ToDouble(textBoxSeqMag2.Text);
            TargetMagn[2] = Convert.ToDouble(textBoxSeqMag3.Text);
            TargetMagn[3] = Convert.ToDouble(textBoxSeqMag4.Text);
            TargetMagn[4] = Convert.ToDouble(textBoxSeqMag5.Text);
            TargetMagn[5] = Convert.ToDouble(textBoxSeqMag6.Text);
            TargetMagn[6] = Convert.ToDouble(textBoxSeqMag7.Text);
            TargetMagn[7] = Convert.ToDouble(textBoxSeqMag8.Text);
            TargetMagn[8] = Convert.ToDouble(textBoxSeqMag9.Text);
            TargetMagn[9] = Convert.ToDouble(textBoxSeqMag10.Text);
            TargetMagn[10] = Convert.ToDouble(textBoxSeqMag11.Text);
            TargetMagn[11] = Convert.ToDouble(textBoxSeqMag12.Text);
            TargetMagn[12] = Convert.ToDouble(textBoxSeqMag13.Text);
            TargetMagn[13] = Convert.ToDouble(textBoxSeqMag14.Text);
            TargetMagn[14] = Convert.ToDouble(textBoxSeqMag15.Text);
            TargetMagn[15] = Convert.ToDouble(textBoxSeqMag16.Text);

            ////Tried to refactor but it's not working
            //for (int i = 0; i < 16; i++)
            //{
            //    //textBoxSeqMag[i] = (TextBox)this.Controls.Find(string.Format("textBoxSeqMag{0}", i), false).FirstOrDefault();
            //    //TargetMagn[i] = Convert.ToDouble(textBoxSeqMag[i].Text);
            //}
            //textBoxSeqMag[0] = (TextBox)textBoxSeqMag1;

            //Frequency
            TargetFreq[0] = Convert.ToDouble(textBoxSeqFreq1.Text);
            TargetFreq[1] = Convert.ToDouble(textBoxSeqFreq2.Text);
            TargetFreq[2] = Convert.ToDouble(textBoxSeqFreq3.Text);
            TargetFreq[3] = Convert.ToDouble(textBoxSeqFreq4.Text);
            TargetFreq[4] = Convert.ToDouble(textBoxSeqFreq5.Text);
            TargetFreq[5] = Convert.ToDouble(textBoxSeqFreq6.Text);
            TargetFreq[6] = Convert.ToDouble(textBoxSeqFreq7.Text);
            TargetFreq[7] = Convert.ToDouble(textBoxSeqFreq8.Text);
            TargetFreq[8] = Convert.ToDouble(textBoxSeqFreq9.Text);
            TargetFreq[9] = Convert.ToDouble(textBoxSeqFreq10.Text);
            TargetFreq[10] = Convert.ToDouble(textBoxSeqFreq11.Text);
            TargetFreq[11] = Convert.ToDouble(textBoxSeqFreq12.Text);
            TargetFreq[12] = Convert.ToDouble(textBoxSeqFreq13.Text);
            TargetFreq[13] = Convert.ToDouble(textBoxSeqFreq14.Text);
            TargetFreq[14] = Convert.ToDouble(textBoxSeqFreq15.Text);
            TargetFreq[15] = Convert.ToDouble(textBoxSeqFreq16.Text);
            //Cycle Count
            TargetCycleCount[0] = Convert.ToDouble(textBoxSeqCycle1.Text);
            TargetCycleCount[1] = Convert.ToDouble(textBoxSeqCycle2.Text);
            TargetCycleCount[2] = Convert.ToDouble(textBoxSeqCycle3.Text);
            TargetCycleCount[3] = Convert.ToDouble(textBoxSeqCycle4.Text);
            TargetCycleCount[4] = Convert.ToDouble(textBoxSeqCycle5.Text);
            TargetCycleCount[5] = Convert.ToDouble(textBoxSeqCycle6.Text);
            TargetCycleCount[6] = Convert.ToDouble(textBoxSeqCycle7.Text);
            TargetCycleCount[7] = Convert.ToDouble(textBoxSeqCycle8.Text);
            TargetCycleCount[8] = Convert.ToDouble(textBoxSeqCycle9.Text);
            TargetCycleCount[9] = Convert.ToDouble(textBoxSeqCycle10.Text);
            TargetCycleCount[10] = Convert.ToDouble(textBoxSeqCycle11.Text);
            TargetCycleCount[11] = Convert.ToDouble(textBoxSeqCycle12.Text);
            TargetCycleCount[12] = Convert.ToDouble(textBoxSeqCycle13.Text);
            TargetCycleCount[13] = Convert.ToDouble(textBoxSeqCycle14.Text);
            TargetCycleCount[14] = Convert.ToDouble(textBoxSeqCycle15.Text);
            TargetCycleCount[15] = Convert.ToDouble(textBoxSeqCycle16.Text);
        }

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

        //display position in waveform format
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
            SetAxisParameters();
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
            //update the sinosoidal parameter inputs
            ReadSinosoidalParas();

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

                updateLimits("150", "1200");
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
                        fileStreamWriter.Write(dataValue.ToString("e2"));
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
            ReadSinosoidalParas();
            textBox_maxRate1.Text = (TargetMagn[0] * 2 * 3.14 * TargetFreq[0]).ToString();
        }

    }
}

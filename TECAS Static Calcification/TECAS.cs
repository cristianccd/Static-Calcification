using LabJack.LabJackUD;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Timers;

namespace TECAS_Static_Calcification
{
    public partial class TECAS : Form
    {
        //Variables
        //***************************
        //pH Calibration
        DateTime pHstartTime;
        double pHCalSlope = 0, pHCalIntercept = 0, pHR2 = 0, SumpHCal = 0, pHTicks = 0, pHAvgVal = 0;
        int pHSampleNo = 0, pHCalState = 0;

        //pH measurement
        double pHMeasureAvg = 0, pHMeasureTicks = 0, pHMeasureVal = 0, pHMeasureAcc=0;
        DateTime pHMeasureStart;
        DialogResult pHQuestSample;
        int Waitlbltick = 0;

        //Syringe Calibration
        double SyrR2=0, SyrCalIntercept=0, SyrCalSlope=0;
        int SampleNo;
        public static int State;
        string _Reading, _SampleUnits, _LoadCal;
        public static double SampleVolume;
        DialogResult SyrQuestSample;
        SyrSamInp SyrVolForm;

        //Labjack Variables (shared)
        private U3 u3;
        LJUD.IO ioType = 0;
        LJUD.CHANNEL channel = 0;
        double dblValue = 0, dummyDouble = 0, dblValueAcc = 0, AvgVoltage = 0;
        int dummyInt = 0;

        //Experiment
        double ExpTicks = 0, ExpAvgVal = 0, ExpAccVal = 0, Deviation=0, VoltoInf=0, TimeDif=0;
        DateTime ExpStart, ExpWaitTime, ExpWdrTime;
        bool InfStarted = false, TimeMix = false, TimeMix_Wdr = false;
        bool Paused = false;
        int ExpState = 0, GraphPt=1, graphUpdate;
        double AccumVolInf = 0, SubSampling=0, WdrVol=1000, CalcVol=0, ReadVol=0;
        static private System.Timers.Timer aTimer;
        StreamWriter sw, swexp, Error_SW;
        int Inf_Ticks = 0;

        bool newErrorLog = true;

        //Zoom
        double xMin, yMin, xMax, yMax;
        double posXStart, posYStart, posXFinish, posYFinish;

        //***************************

        public TECAS()
        {
            InitializeComponent();
            //Disable Tab2 phmeasurement
            EnableTab(tabPage2, false);
            //pH calibration initialization
            dataGridView2.Rows.Add(2);

            //add 2 rows for each table (minimum for calibration)
            dataGridView1.Rows.Add(2);
            dataGridView1.Columns[2].ReadOnly = true;
            dataGridView1.Columns[3].ReadOnly = true;
            //disable
        }
        
        //*********************************************************************************
        //***********************************PH CALIBRATION********************************
        
        //Add Row to the sample table
        private void button8_Click(object sender, EventArgs e)
        {
            dataGridView2.Rows.Add();
            if (dataGridView2.RowCount > 2)
                button9.Enabled = true;
        }
        
        //Delete Row to the sample table
        private void button9_Click(object sender, EventArgs e)
        {
            if (dataGridView2.RowCount > 2)
                dataGridView2.Rows.RemoveAt(dataGridView2.Rows.Count - 1);
            if (dataGridView2.RowCount <= 2)
                button9.Enabled = false;
        }

        //Check for errors
        private bool pHCalErr()
        {
            try
            {
                Convert.ToDouble(textBox1.Text);
                if (Convert.ToDouble(textBox1.Text) < 0)
                {
                    MessageBox.Show("Time must be possitive!", "Error!", MessageBoxButtons.OK,MessageBoxIcon.Error);
                    return false;
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Time format not valid", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            if (dataGridView2.RowCount < 2)
            {
                MessageBox.Show("Not eneough Samples", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            //Check the samples
            for (int i = 0; i < dataGridView2.RowCount; i++)
            {
                try
                {
                    Convert.ToDouble(dataGridView2[0,i].Value);
                    if (Convert.ToDouble(dataGridView2[0, i].Value) <= 0 || Convert.ToDouble(dataGridView2[0, i].Value) > 14)
                    {
                        MessageBox.Show("pH sample No." + Convert.ToString(i + 1) + " not correct!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                }
                catch (Exception )
                {
                    MessageBox.Show("pH format not valid", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }
            return true;
        }
        //Open labjack and start timer
        private void button10_Click(object sender, EventArgs e) //Start pH
        {
            //disable the timer for pH measurement (if running)
            timer2.Enabled = false;
            //Check for errors
            if(pHCalErr())
            {
                //Open labjack and catch exception if not able to open 
                try
                {
                    if (u3 == null)
                        u3 = new U3(LJUD.CONNECTION.USB, "1", true); // Connection through USB
                    LJUD.ePut(u3.ljhandle, LJUD.IO.PIN_CONFIGURATION_RESET, 0, 0, 0);
                    LJUD.ePut(u3.ljhandle, LJUD.IO.PUT_ANALOG_ENABLE_PORT, 0, 31, 16);//first 5 FIO analog b0000000000011111
                    LJUD.AddRequest(u3.ljhandle, LJUD.IO.GET_AIN_DIFF, 4, 0, 32, 0);//Request FIO4
                    LJUD.GoOne(u3.ljhandle);
                }
                catch (LabJackUDException h)
                {
                    MessageBox.Show("Error opening DAQ. "+h.ToString(), "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                foreach (var series in chart3.Series)
                    series.Points.Clear();
                //Block pH controls
                textBox1.Enabled = false;
                button8.Enabled = false;
                //Block delete button if less than 2 samples
                if (dataGridView2.RowCount <= 2)
                    button9.Enabled = false;
                button10.Enabled = false;
                button11.Enabled = false;
                button12.Enabled = false;
                label48.Visible = true;
                panel16.Enabled = false;
                button19.Enabled = true;
                //Set to 0 all the variables involved
                pHCalState = 0;
                pHCalSlope = 0;
                pHCalIntercept = 0;
                pHR2 = 0;
                SumpHCal = 0;
                pHTicks = 0;
                pHAvgVal = 0;
                pHSampleNo = 0;
                label10.Text = "0.00";
                label12.Text = "00:00";
                //Put in 0 the data from the table
                for (int i = 0; i < dataGridView2.RowCount; i++)
                    dataGridView2[1, i].Value = "";
                //Disable the rest of the tabs
                EnableTab(tabPage2, false);
                EnableTab(tabPage3, false);
                EnableTab(tabPage4, false);
                //enable timer again
                timer3.Enabled = true;
            }
        }

        //Timer for the calibration of the electrode
        private void timer3_Tick(object sender, EventArgs e) //every loop will enter the state machine and also read the DAQ
        {
            timer3.Enabled = false; //first disable the timer to perform actions
            bool requestedExit = false;
            //Blink label Please Wait...
            Waitlbltick++;
            if (Waitlbltick >= 10)
            {
                Waitlbltick = 0;
                label48.Visible = !label48.Visible;
            }
            //Start of calibration
            switch (pHCalState) // Check for errors
            {
                case 0://calibration started
                    pHQuestSample = MessageBox.Show("When you are ready to measure the Sample " + Convert.ToString(pHSampleNo + 1) + " press OK", "pH Meter Calibration", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                    //if the result is OK, change state and start counting time
                    if (pHQuestSample == DialogResult.OK)
                    {
                        pHCalState = 1;
                        pHstartTime = DateTime.Now;
                    }
                    //If cancelled stop and cancel the test
                    if (pHQuestSample == DialogResult.Cancel)
                    {
                        //Enable pH controls
                        CancelpHTest();
                        return;
                    }
                    timer3.Enabled = true;
                    break;
                case 1: //Start accumulating time and check if the target is reached
                    label12.Text = String.Format("{0:00}", (DateTime.Now - pHstartTime).Minutes) + ":" + String.Format("{0:00}", (DateTime.Now - pHstartTime).Seconds);
                    //If the time is reached -> state 2
                    if (Convert.ToDouble((DateTime.Now - pHstartTime).TotalMinutes) >= Convert.ToDouble(textBox1.Text))
                    {
                        pHCalState = 2; //Time reached
                        timer3.Enabled = true;
                        break;
                    }
                    //Start adding the ticks to make the avg
                    pHTicks++;
                    //Add the voltage value in an accum.
                    SumpHCal = SumpHCal + dblValue;
                    timer3.Enabled = true;
                    break;
                case 2://Time reached, change sample or exit
                    //Calculate the average voltage
                    pHAvgVal = SumpHCal / pHTicks;
                    dataGridView2[1, pHSampleNo].Value = String.Format("{0:0.0000}", pHAvgVal); ;
                    //reset values
                    pHSampleNo++;
                    pHTicks = 0;
                    SumpHCal = 0;
                    pHCalState = 0;
                    //If the sample to check is higher than the count its over.
                    if (pHSampleNo >= dataGridView2.RowCount)
                        pHCalState = 10;
                    timer3.Enabled = true;
                    break;
                case 10://Exit: calculate the results
                    MessageBox.Show("Finished, please check the results...", "Finished", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //auxiliar variables to obtain results
                    double a = 0, b = 0, c = 0, d = 0, f = 0;
                    for (int i = 0; i < dataGridView2.RowCount; i++)
                    {
                        //sumyx
                        a= a + Convert.ToDouble(dataGridView2[1, i].Value) * Convert.ToDouble(dataGridView2[0, i].Value);
                        //sumx
                        b= b + Convert.ToDouble(dataGridView2[0, i].Value);
                        //sumy
                        c= c +Convert.ToDouble(dataGridView2[1, i].Value);
                        //sumx2
                        d= d + Math.Pow(Convert.ToDouble(dataGridView2[0, i].Value), 2);
                        //sumy2
                        f= f + Math.Pow(Convert.ToDouble(dataGridView2[1, i].Value),2);
                        chart3.Series["Series2"].Points.AddXY(Convert.ToDouble(dataGridView2[0,i].Value),dataGridView2[1, i].Value);
                    }
                    /*R2 = ((Nxysum - xsumysum)^2)/((Nx^2sum - xsum*xsum)*(Ny^2sum - ysum*ysum))*/
                    pHR2=(Math.Pow((dataGridView2.RowCount*a-b*c),2))/((dataGridView2.RowCount*d-Math.Pow(b,2))*(dataGridView2.RowCount*f-Math.Pow(c,2)));
                    /*Slope(b) = (NΣXY - (ΣX)(ΣY)) / (NΣX2 - (ΣX)2) Intercept(a) = (ΣY - b(ΣX)) / N */
                    pHCalSlope = (dataGridView2.RowCount * a - b * c) / (dataGridView2.RowCount*d - Math.Pow(b, 2));
                    pHCalIntercept=(c-pHCalSlope*b)/dataGridView2.RowCount;
                    //Show the label for positive or negative according to the intercept
                    if(pHCalIntercept>=0)
                        label6.Text = "y=" + String.Format("{0:0.0000}", pHCalSlope) + " x+" + String.Format("{0:0.0000}", pHCalIntercept);
                    else
                        label6.Text = "y=" + String.Format("{0:0.0000}", pHCalSlope) + " x" + String.Format("{0:0.0000}", pHCalIntercept);
                    //Write the R2
                    label5.Text = "R =  " + String.Format("{0:0.0000}", pHR2);
                    //Graph
                    chart3.Series["Series1"].Points.AddXY(Convert.ToDouble(dataGridView2[0,0].Value),pHCalSlope*Convert.ToDouble(dataGridView2[0,0].Value)+pHCalIntercept);
                    chart3.Series["Series1"].Points.AddXY(Convert.ToDouble(dataGridView2[0, dataGridView2.RowCount-1].Value), pHCalSlope * Convert.ToDouble(dataGridView2[0, dataGridView2.RowCount - 1].Value) + pHCalIntercept);
                    //Show the equations
                    panel1.Visible = true;
                    //Enable add row
                    button8.Enabled = true;
                    //If the row amount is higher than 2 also enable the delete
                    if (dataGridView2.RowCount > 2)
                        button9.Enabled = true;
                    //Enable the rest of the things and tabs
                    button10.Enabled = true;
                    button11.Enabled = true;
                    button12.Enabled = true;
                    EnableTab(tabPage2, true);
                    EnableTab(tabPage3, true);
                    EnableTab(tabPage4, true);
                    //Set everything to 0 again (not pHCalSlope and pHCalIntercept)
                    pHTicks = 0;
                    SumpHCal = 0;
                    pHSampleNo = 0;
                    pHCalState = 0;
                    //Enable start in pH read
                    button6.Enabled = true;
                    button7.Enabled = false;
                    //Enable the textbox in pH read
                    textBox3.Enabled = true;
                    //Check the box for successful calibration
                    checkBox2.Checked = true;
                    //Disable the blinking text
                    label48.Visible = false;
                    //Disable Cancel
                    button19.Enabled = false;
                    panel16.Enabled = true;
                    //Enable the time textbox
                    textBox1.Enabled = true;
                    label23.Text = String.Format("{0:0.0000}", pHCalSlope);
                    return;
            }
            //Read the value in the DAQ
            while (!requestedExit)
            {
                try
                {
                    LJUD.GoOne(u3.ljhandle);
                    LJUD.GetFirstResult(u3.ljhandle, ref ioType, ref channel, ref dblValue, ref dummyInt, ref dummyDouble);
                }
                catch (LabJackUDException h)
                {
                    CancelpHTest();
                    MessageBox.Show("Error getting DAQ data. " + h.ToString(), "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    requestedExit = true;
                }
                if (ioType == LJUD.IO.GET_AIN_DIFF)
                    label10.Text = String.Format("{0:0.0000}", dblValue);
                try
                {
                    LJUD.GetNextResult(u3.ljhandle, ref ioType, ref channel, ref dblValue, ref dummyInt, ref dummyDouble);
                }
                catch (LabJackUDException h)
                {
                    if (h.LJUDError == U3.LJUDERROR.NO_MORE_DATA_AVAILABLE)
                        requestedExit = true;//no more data to read
                    else
                    {
                        CancelpHTest();
                        MessageBox.Show("Error getting DAQ data. " + h.ToString(), "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        requestedExit = true;
                    }
                }
            }
        }
        //Open and load the file from PC
        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                if (openFileDialog2.ShowDialog() == DialogResult.OK)
                {
                    foreach (var series in chart3.Series)
                        series.Points.Clear();
                    _LoadCal = File.ReadAllText(openFileDialog2.FileName);
                    pHCalSlope=Convert.ToDouble(_LoadCal.Split('#')[0]);
                    pHCalIntercept = Convert.ToDouble(_LoadCal.Split('#')[1]);
                    pHR2 = Convert.ToDouble(_LoadCal.Split('#')[2]);
                    if(pHCalIntercept>=0)
                        label6.Text = "y=" + String.Format("{0:0.0000}", pHCalSlope) + " x+" + String.Format("{0:0.0000}", pHCalIntercept);
                    else
                        label6.Text = "y=" + String.Format("{0:0.0000}", pHCalSlope) + " x" + String.Format("{0:0.0000}", pHCalIntercept);
                    label5.Text = "R  =" + String.Format("{0:0.0000}", pHR2);
                    label23.Text = String.Format("{0:0.0000}", pHCalSlope);
                    panel1.Visible = true;
                    chart3.Series["Series1"].Points.AddXY(0, pHCalIntercept);
                    chart3.Series["Series1"].Points.AddXY(14, pHCalSlope * 14 + pHCalIntercept);
                    button12.Enabled = true;
                    button6.Enabled = true;
                    button7.Enabled = false;
                    textBox3.Enabled = true;
                    checkBox2.Checked = true;
                    EnableTab(tabPage2, true);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Calibration", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        //Save file to PC
        private void button12_Click(object sender, EventArgs e)
        {
            try
            {
                if (saveFileDialog2.ShowDialog() == DialogResult.OK)
                {
                    sw = new StreamWriter(saveFileDialog2.FileName);
                    sw.Write(Convert.ToString(pHCalSlope)+"#"+Convert.ToString(pHCalIntercept)+"#"+Convert.ToString(pHR2));
                    sw.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Calibration", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //If Cancel button is pressed
        private void button19_Click(object sender, EventArgs e)
        {
            CancelpHTest();
            checkBox2.Checked = false;
            EnableTab(tabPage2, false);
            EnableTab(tabPage3, true);
            EnableTab(tabPage4, true);
        }
        //Cancel pH calibration
        private void CancelpHTest()
        {
            //disable timer
            timer3.Enabled = false;
            //Enable the controls
            textBox1.Enabled = true;
            button8.Enabled = true;
            button9.Enabled = false;
            button10.Enabled = true;
            button11.Enabled = true;
            button12.Enabled = false;
            //Set everything back to normal state
            pHCalState = 0;
            pHCalSlope = 0;
            pHCalIntercept = 0;
            pHR2 = 0;
            SumpHCal = 0;
            pHTicks = 0;
            pHAvgVal = 0;
            pHSampleNo = 0;
            label48.Visible = false;
            button19.Enabled = false;
            panel16.Enabled = true;
            label10.Text = "0.00";
            label12.Text = "00:00";
            //Clear the values and add two rows
            dataGridView2.Rows.Clear();
            dataGridView2.Rows.Add(2);
            EnableTab(tabPage3, true);
            EnableTab(tabPage4, true);
        }

        //*********************************************************************************
        //***********************************END*******************************************

        //*********************************************************************************
        //******************************PH MEASUREMENT*************************************

        private void button6_Click(object sender, EventArgs e)
        {
            //Check for calibration
            if (!checkBox2.Checked)
            {
                MessageBox.Show("pH Calibration not done!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //Read the setpoint
            if (textBox3.Text != "")
            {
                try
                {
                    Convert.ToDouble(textBox3.Text);
                }
                catch (Exception)
                {
                    MessageBox.Show("Invalid Set Point Format!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            else
            {
                textBox3.Text = "0";
            }
            //disable start and enable stop
            button6.Enabled = false;
            textBox3.Enabled = false;
            button7.Enabled = true;

            //Configure DAQ
            //Open labjack and catch exception if not able to open          
            try
            {
                if (u3 == null)
                    u3 = new U3(LJUD.CONNECTION.USB, "1", true); // Connection through USB
                ConfigLJ();
            }
            catch (LabJackUDException h)
            {
                MessageBox.Show("Error opening DAQ. "+ h.ToString(), "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            pHMeasureStart = DateTime.Now;
            chart2.ChartAreas[0].AxisY.LabelStyle.Format = "#.###";
            chart2.ChartAreas[0].AxisX.LabelStyle.Format = "#";
            chart2.ChartAreas[0].AxisY.ScrollBar.Enabled = false;
            chart2.ChartAreas[0].AxisX.ScrollBar.Enabled = false;
            foreach (var series in chart2.Series)
                series.Points.Clear();
            pHMeasureVal = 0;
            pHMeasureAcc = 0;
            dblValueAcc = 0;
            pHMeasureTicks =0;
            pHMeasureTicks = 0;
            //Enable timer for pH reading
            timer2.Enabled = true;
        }
        
        //Read pH value

        private void timer2_Tick(object sender, EventArgs e)
        {
            //disable timer
            timer2.Enabled = false;
            //Read from DAQ
            bool requestedExit = false;
            while (!requestedExit)
            {
                //Read first value and check the voltage
                try
                {
                    LJUD.GoOne(u3.ljhandle);
                    LJUD.GetFirstResult(u3.ljhandle, ref ioType, ref channel, ref dblValue, ref dummyInt, ref dummyDouble);
                }
                catch (LabJackUDException h)
                {
                    //MessageBox.Show("Error getting the DAQ data. " + h.ToString(), "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UpdateStatStrip(h, false);
                    System.Threading.Thread.Sleep(10000);
                    ConfigLJ();
                }
                if (ioType == LJUD.IO.GET_AIN_DIFF)
                {
                    pHMeasureVal = (dblValue - pHCalIntercept) / pHCalSlope;
                    pHMeasureAcc += pHMeasureVal;
                    dblValueAcc += dblValue;
                    pHMeasureTicks++; //Start accum of read values every 75ms to make an average
                    if (pHMeasureTicks >= 10)//average between N samples
                    {
                        pHMeasureAvg = pHMeasureAcc / 10;
                        AvgVoltage = dblValueAcc / 10;
                        label3.Text = String.Format("{0:0.000}", pHMeasureAvg);
                        label42.Text = String.Format("{0:0.0000 V}", AvgVoltage);
                        if (Convert.ToDouble(textBox3.Text) == 0)
                            label37.Text = "0.000";
                        else
                            label37.Text = String.Format("{0:0.000}", pHMeasureAvg - Convert.ToDouble(textBox3.Text));
                        chart2.Series["Series1"].Points.AddXY((DateTime.Now - pHMeasureStart).TotalSeconds, pHMeasureAvg);
                        //If the seconds are more than 300 start resizing
                        if ((DateTime.Now - pHMeasureStart).TotalSeconds>300)
                            chart2.ChartAreas[0].AxisX.ScaleView.Position = (DateTime.Now-pHMeasureStart).TotalSeconds - 300;
                        //Reset the accum.
                        pHMeasureAcc = 0;
                        dblValueAcc = 0;
                        pHMeasureTicks = 0;
                    }
                }
                try
                {
                    LJUD.GetNextResult(u3.ljhandle, ref ioType, ref channel, ref dblValue, ref dummyInt, ref dummyDouble);
                }
                catch (LabJackUDException h)
                {
                    if (h.LJUDError == U3.LJUDERROR.NO_MORE_DATA_AVAILABLE)
                        requestedExit = true;//no more data to read
                    else
                    {
                        //MessageBox.Show("Error getting DAQ data. " + h.ToString(), "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        UpdateStatStrip(h, false);
                        System.Threading.Thread.Sleep(10000);
                        ConfigLJ();
                    }
                }
            }
            //enable timer again
            timer2.Enabled = true;
        }
        //Stop the reading
        private void button7_Click(object sender, EventArgs e)
        {
            button7.Enabled = false;
            button6.Enabled = true;
            textBox3.Enabled = true;
            timer2.Enabled = false;
            //Disable the watchdog
            if(u3 != null)
                LJUD.ePut(u3.ljhandle, LJUD.IO.SWDT_CONFIG, LJUD.CHANNEL.SWDT_DISABLE, 0, 0); 
        }
        //*********************************************************************************
        //***********************************END*******************************************

        //*********************************************************************************
        //******************************SYRINGE CALIBRATION********************************

        //Add row
        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Add();
            if (dataGridView1.RowCount > 2)
                button2.Enabled = true;
        }
        
        //Del row
        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount > 2)
                dataGridView1.Rows.RemoveAt(dataGridView1.Rows.Count - 1);
            if (dataGridView1.RowCount <= 2)
                button2.Enabled = false;
        }

        //Check for errors
        private bool CheckSyrCalErr()
        {
            if (comboBox22.SelectedIndex == -1)
            {
                MessageBox.Show("COM port not selected!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            try//if diameter and capacity are numbers
            {
                Convert.ToDouble(textBox21.Text);
                Convert.ToDouble(textBox22.Text);
            }
            catch (Exception) //Diameter or capacity not double
            {
                MessageBox.Show("Wrong input parameters!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                try//if the values in the data grid are numbers 
                {
                    //lets check numbers first
                    Convert.ToDouble(dataGridView1[0, i].Value);
                    Convert.ToDouble(dataGridView1[2, i].Value);
                }
                catch (Exception) //Diameter or apacity not double
                {
                    MessageBox.Show("Please check the sample's volumes", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                if (dataGridView1[1, i].Value == null)
                {
                    MessageBox.Show("Units are not properly selected", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }

            if (dataGridView1.RowCount < 2)
            {
                MessageBox.Show("Not enough samples for calibration!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            return true;//Everything OK
        }

        //Start Syr Calibration
        private void button3_Click(object sender, EventArgs e)
        {
            if (CheckSyrCalErr())//Check errors
            {
                //open port if not open
                if (comboBox22.SelectedIndex != -1)
                {
                    if (serialPort1.IsOpen)
                        serialPort1.Close();
                    try
                    {
                        serialPort1.PortName = "COM" + Convert.ToString(comboBox22.SelectedIndex);
                        serialPort1.Open();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Opening Port", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    } 
                }
                //Set everything before starting
                SampleNo = 0;
                State = 0;
                //Send config commands to the pump
                serialPort1.Write("STP\r\n");
                System.Threading.Thread.Sleep(20);
                serialPort1.Write("CLD INF\r\n");
                System.Threading.Thread.Sleep(20);
                serialPort1.Write("CLD WDR\r\n");
                System.Threading.Thread.Sleep(20);
                serialPort1.Write("DIA " + textBox21.Text + "\r\n");
                System.Threading.Thread.Sleep(20);
                serialPort1.Write("DIR INF\r\n");
                System.Threading.Thread.Sleep(20);
                serialPort1.Write("RAT 800 MH\r\n");
                System.Threading.Thread.Sleep(20);
                //Clear the series
                foreach (var series in chart1.Series)
                    series.Points.Clear();
                //Block controls
                dataGridView1.Enabled = false;
                comboBox22.Enabled = false;
                textBox21.Enabled = false;
                textBox22.Enabled = false;
                checkBox4.Enabled = false;
                panel14.Enabled = false;
                button4.Enabled = false;
                button5.Enabled = false;
                button1.Enabled = false;
                button2.Enabled = false;
                button3.Enabled = false;
                button20.Enabled = true;
                //Enable timer
                timer1.Enabled = true;
            }  
        }

        public void UpdateStatStrip(Exception Ex, bool ReadExp)
        {
            if (newErrorLog)
            {
                string Path = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\Calcification Experiments";
                if (!System.IO.Directory.Exists(Path))
                    System.IO.Directory.CreateDirectory(Path);
                Error_SW = new StreamWriter(Path + @"\ErrorLog_" + DateTime.Now.ToString("yyyy-MM-dd HHmmss") + ".txt");
                newErrorLog = false;
            }
            if (ReadExp)
            {
                Error_SW.Write("Error during experiment @ " + DateTime.Now.ToString() + " - " + Ex.Message + "\r\n");
                toolStripStatusLabel1.Text = "Error during experiment @ " + DateTime.Now.ToString() + " - " + Ex.Message;
            }
            else
            {
                Error_SW.Write("Error reading pH @ " + DateTime.Now.ToString() + " - " + Ex.Message + "\r\n");
                toolStripStatusLabel1.Text = "Error reading pH @ " + DateTime.Now.ToString() + " - " + Ex.Message;
            }
            Error_SW.Flush();
            statusStrip1.BackColor = Color.Red;
           
        }


        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            switch (State)
            { 
                case 0://sampling started
                    SyrQuestSample=MessageBox.Show("When you are ready to measure the sample " + Convert.ToString(SampleNo+1) + " press OK", "Infuse Sample", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                    if (SyrQuestSample == DialogResult.OK)
                    {
                        //set the infuse volumes according to the number of the sample
                        serialPort1.Write("VOL " + dataGridView1[0, SampleNo].Value.ToString() + "\r\n");
                        System.Threading.Thread.Sleep(20);
                        //if ml, set to ml
                        if (dataGridView1[1, SampleNo].Value.ToString() == "ml")
                            serialPort1.Write("VOL ML\r\n");
                        //otherwise ul
                        else
                            serialPort1.Write("VOL UL\r\n");
                        System.Threading.Thread.Sleep(20);
                        serialPort1.Write("RUN\r\n"); //start running
                        State = 1; //switch state to running, needs to wait the ammount desired
                        SampleNo++;
                        label1.Text = "Sample No.: " + Convert.ToString(SampleNo);
                    }
                    if (SyrQuestSample == DialogResult.Cancel)
                    {
                        dataGridView1.Enabled = true;
                        comboBox22.Enabled = true;
                        textBox21.Enabled = true;
                        textBox22.Enabled = true;
                        checkBox4.Enabled = true;
                        button16.Enabled = true;
                        button17.Enabled = true;
                        button4.Enabled = true;
                        button5.Enabled = false;
                        button1.Enabled = true;
                        if (dataGridView1.RowCount > 2)
                            button2.Enabled = true;
                        button3.Enabled = true;
                        button20.Enabled = false;
                        serialPort1.Close();
                        return;
                    }
                    timer1.Enabled = true;    
                    break;
                case 1:
                    serialPort1.ReadExisting();
                    serialPort1.Write("DIS\r\n");
                    System.Threading.Thread.Sleep(20);
                    _Reading=serialPort1.ReadExisting();
                    label33.Text = "Current Volume: " + _Reading.Substring(5, 5) + _Reading.Substring(16, 2).ToLower();
                    //Infusion completed -> State 2
                    if (Convert.ToDouble(_Reading.Substring(5, 5)) >= Convert.ToDouble(dataGridView1[0, SampleNo-1].Value.ToString()))
                        State = 2;//input volume form2
                    timer1.Enabled = true;
                    break;
                case 2:
                    _SampleUnits = dataGridView1[1, SampleNo-1].Value.ToString();
                    //Create a new form with the sample sent to write
                    SyrVolForm = new SyrSamInp(_SampleUnits);
                    SyrVolForm.Show();
                    this.Enabled = false;
                    State = 3;
                    timer1.Enabled = true;
                    break;
                case 3:
                    timer1.Enabled = true;
                    break;
                case 4:
                    //enable form again
                    this.Enabled = true;
                    dataGridView1[2, SampleNo - 1].Value = SampleVolume;
                    _SampleUnits = Convert.ToString(dataGridView1[1, SampleNo - 1].Value);
                    dataGridView1[3, SampleNo - 1].Value = _SampleUnits;
                    if (SampleNo == dataGridView1.Rows.Count) //end of samples
                    {
                        label1.Text = "Sample No.: " + Convert.ToString(SampleNo);
                        State = 10; //quit
                    }
                    else // go one more sample
                        State = 0;
                    timer1.Enabled = true;
                    break;
                case 10:
                    MessageBox.Show("Finished, please check the results...", "Finished", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    double a = 0, b = 0, c = 0, d = 0, f = 0;
                    for (int i = 0; i < dataGridView1.RowCount; i++)
                    {
                        if (dataGridView1[1, i].Value.ToString() == "ml")
                        {
                            double Aux1 = 0, Aux2 = 0;
                            Aux1 = Convert.ToDouble(dataGridView1[0, i].Value) * 1000;
                            Aux2 = Convert.ToDouble(dataGridView1[2, i].Value) * 1000;
                            //sumxy
                            a = a + Aux1*Aux2;
                            //sumx
                            b = b + Aux1;
                            //sumy
                            c = c + Aux2;
                            //sumx2
                            d = d + Math.Pow(Aux1, 2);
                            //sumy2
                            f = f + Math.Pow(Aux2, 2);
                            chart1.Series["Series1"].Points.AddXY(Aux1, Aux2);
                            if (i == 0 || i == dataGridView1.RowCount-1)
                                chart1.Series["Series2"].Points.AddXY(Aux1, Aux2);
                        }
                        else
                        {
                            //sumxy
                            a = a + Convert.ToDouble(dataGridView1[0, i].Value) * Convert.ToDouble(dataGridView1[2, i].Value);
                            //sumx
                            b = b + Convert.ToDouble(dataGridView1[0, i].Value);
                            //sumy
                            c = c + Convert.ToDouble(dataGridView1[2, i].Value);
                            //sumx2
                            d = d + Math.Pow(Convert.ToDouble(dataGridView1[0, i].Value), 2);
                            //sumy2
                            f = f + Math.Pow(Convert.ToDouble(dataGridView1[2, i].Value), 2);
                            chart1.Series["Series1"].Points.AddXY(Convert.ToDouble(dataGridView1[0, i].Value), dataGridView1[2, i].Value);
                            if (i == 0 || i == dataGridView1.RowCount - 1)
                                chart1.Series["Series2"].Points.AddXY(Convert.ToDouble(dataGridView1[0, i].Value), dataGridView1[2, i].Value);
                        }
                    }
                    /*R2 = ((Nxysum - xsumysum)^2)/(Nx^2sum - xsum*xsum)*(Ny^2sum - ysum*ysum)*/
                SyrR2 =(Math.Pow((dataGridView1.RowCount*a-b*c),2))/((dataGridView1.RowCount*d-Math.Pow(b,2))*(dataGridView1.RowCount*f-Math.Pow(c,2)));
                    /*Slope(b) = (NΣXY - (ΣX)(ΣY)) / (NΣX2 - (ΣX)2) Intercept(a) = (ΣY - b(ΣX)) / N */
                    SyrCalSlope = (dataGridView1.RowCount * a - b * c) / (dataGridView1.RowCount*d - Math.Pow(b, 2));
                    SyrCalIntercept=(c-SyrCalSlope*b)/dataGridView1.RowCount;
                    //Write text values with the correct intercept
                    label30.Text = "R =  " + String.Format("{0:0.0000}", SyrR2);
                    if (SyrCalIntercept>=0)
                        label29.Text = "y=" + String.Format("{0:0.0000}", SyrCalSlope) + " x+" + String.Format("{0:0.0000}", SyrCalIntercept);
                    else
                        label29.Text = "y=" + String.Format("{0:0.0000}", SyrCalSlope) + " x" + String.Format("{0:0.0000}", SyrCalIntercept);
                    serialPort1.Close();
                    //Check the box for correct calibration
                    checkBox1.Checked = true;
                    //Make everything visible again
                    panel5.Visible = true;
                    dataGridView1.Enabled = true;
                    comboBox22.Enabled = true;
                    textBox21.Enabled = true;
                    textBox22.Enabled = true;
                    //Enable checkbox, disable controls of manual infusion
                    checkBox4.Checked = false;
                    checkBox4.Enabled = true;
                    panel14.Enabled = true;
                    button16.Enabled = false;
                    button17.Enabled = false;
                    button18.Enabled = false;
                    //Enable buttons for open and save
                    button4.Enabled = true;
                    button5.Enabled = true;
                    button1.Enabled = true;
                    //Enable delete only if there are more than 2 samples
                    if(dataGridView1.RowCount>2)
                        button2.Enabled = true;
                    button3.Enabled = true;
                    button20.Enabled = false;
                    label35.Text = String.Format("{0:0.0000}", SyrCalSlope);
                    return;
            }
        }

        //Cancel sampling
        private void button20_Click(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            dataGridView1.Enabled = true;
            comboBox22.Enabled = true;
            textBox21.Enabled = true;
            textBox22.Enabled = true;
            checkBox4.Enabled = true;
            checkBox4.Checked = false;
            panel14.Enabled = true;
            button16.Enabled = false;
            button17.Enabled = false;
            button4.Enabled = true;
            button5.Enabled = false;
            button1.Enabled = true;
            checkBox1.Enabled = false;
            //Enable delete only if there are more than 2 samples
            if (dataGridView1.RowCount > 2)
                button2.Enabled = true;
            button3.Enabled = true;
            button20.Enabled = false;
            serialPort1.Close();
        }

        //Open calibration
        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    foreach (var series in chart1.Series)
                        series.Points.Clear();
                    _LoadCal = File.ReadAllText(openFileDialog1.FileName);
                    /*COM#Diameter#Capacity#x1#y1#x2#y2#slope#intercept#r2#*/
                    comboBox22.SelectedIndex = Convert.ToInt16(_LoadCal.Split('#')[0]);
                    textBox21.Text = _LoadCal.Split('#')[1];
                    textBox22.Text = _LoadCal.Split('#')[2];
                    SyrCalSlope = Convert.ToDouble(_LoadCal.Split('#')[7]);
                    SyrCalIntercept = Convert.ToDouble(_LoadCal.Split('#')[8]);
                    SyrR2 = Convert.ToDouble(_LoadCal.Split('#')[9]);
                    if (SyrCalIntercept >= 0)
                        label29.Text = "y=" + String.Format("{0:0.0000}", SyrCalSlope) + " x+" + String.Format("{0:0.0000}", SyrCalIntercept);
                    else
                        label29.Text = "y=" + String.Format("{0:0.0000}", SyrCalSlope) + " x" + String.Format("{0:0.0000}", SyrCalIntercept);
                    label30.Text = "R =  " + String.Format("{0:0.0000}", SyrR2);
                    label35.Text = String.Format("{0:0.0000}", SyrCalSlope);
                    panel5.Visible = true;
                    chart1.Series["Series2"].Points.AddXY(Convert.ToDouble(_LoadCal.Split('#')[3]), Convert.ToDouble(_LoadCal.Split('#')[4]));
                    chart1.Series["Series2"].Points.AddXY(Convert.ToDouble(_LoadCal.Split('#')[5]), Convert.ToDouble(_LoadCal.Split('#')[6]));
                    checkBox1.Checked = true;
                    button5.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Calibration", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        //Save
        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    sw = new StreamWriter(saveFileDialog1.FileName);
                    /*COM#Diameter#Capacity#x1#y1#x2#y2#slope#intercept#r2#*/
                    sw.Write(comboBox22.SelectedIndex.ToString() + "#" + textBox21.Text + "#" + textBox22.Text + "#" + Convert.ToString(chart1.Series["Series2"].Points[0].XValue) + "#" + Convert.ToString(chart1.Series["Series2"].Points[0].YValues[0]) + "#" + Convert.ToString(chart1.Series["Series2"].Points[1].XValue) + "#" + Convert.ToString(chart1.Series["Series2"].Points[1].YValues[0]) + "#" + Convert.ToString(SyrCalSlope) + "#" + Convert.ToString(SyrCalIntercept) + "#" + Convert.ToString(SyrR2));
                    sw.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Calibration", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }  
        }

        private void dataGridView1_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            if (dataGridView1.RowCount > 2)
            {
                button2.Enabled = true;
                button3.Enabled = true;
            }
        }

        private void dataGridView1_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            if (dataGridView1.RowCount <= 2)
                button2.Enabled = false;
        }

        //Manual infusion
        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked == true)
            {
                //Check for errors
                if (string.IsNullOrWhiteSpace(textBox21.Text) || string.IsNullOrWhiteSpace(textBox22.Text))
                {
                    MessageBox.Show("Please input the syringe specifications", "Diameter and capacity not set", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    checkBox4.Checked = false;
                    return;
                }
                try
                {
                    if (!serialPort1.IsOpen)
                    {
                        serialPort1.PortName = "COM" + Convert.ToString(comboBox22.SelectedIndex);
                        serialPort1.Open();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Port opening", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    checkBox4.Checked = false;
                    return;
                }
                try
                {
                    //set the initial conditions to 0
                    serialPort1.Write("VOL 0\r\n");
                    System.Threading.Thread.Sleep(20);
                    serialPort1.Write("VOL UL\r\n");
                    System.Threading.Thread.Sleep(20);
                    serialPort1.Write("CLD INF\r\n");
                    System.Threading.Thread.Sleep(20);
                    serialPort1.Write("CLD WDR\r\n");
                    System.Threading.Thread.Sleep(20);
                    serialPort1.Write("DIA " + textBox21.Text + "\r\n");
                    System.Threading.Thread.Sleep(20);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Port communication!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    checkBox4.Checked = false;
                    return;
                }
                label44.Text = "0 ul";
                label46.Text = "0 ul";
                button16.Enabled = true;
                button17.Enabled = true;
                button18.Enabled = true;
                
            }
            else
            {
                button16.Enabled = false;
                button17.Enabled = false;
                button18.Enabled = false;
                try
                {
                    if (serialPort1.IsOpen)
                        serialPort1.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Port opening", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }
        //Infuse Manual
        private void button17_MouseDown(object sender, MouseEventArgs e)
        {
            try
            {
                if (!serialPort1.IsOpen)
                {
                    serialPort1.PortName = "COM" + Convert.ToString(comboBox22.SelectedIndex);
                    serialPort1.Open();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Port opening", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            serialPort1.Write("STP\r\n");
            System.Threading.Thread.Sleep(20);
            serialPort1.Write("DIR INF\r\n");
            System.Threading.Thread.Sleep(20);
            serialPort1.Write("RAT 200 MH\r\n"); //Rate fixed
            System.Threading.Thread.Sleep(20);
            serialPort1.Write("RUN\r\n");
            System.Threading.Thread.Sleep(20);
        }
        private void button17_MouseUp(object sender, MouseEventArgs e)
        {
            serialPort1.Write("STP\r\n");
            System.Threading.Thread.Sleep(20);
            _Reading = serialPort1.ReadExisting();
            serialPort1.Write("DIS\r\n");
            System.Threading.Thread.Sleep(20);
            _Reading = serialPort1.ReadExisting();
            label44.Text = String.Format("{0:0.0000}",(Convert.ToDouble(_Reading.Substring(5, 5))))+ " " + _Reading.Substring(16, 2).ToLower(); 
            serialPort1.Close();
        }
        //Withdraw Manual
        private void button16_MouseDown(object sender, MouseEventArgs e)
        {
            try
            {
                if (!serialPort1.IsOpen)
                {
                    serialPort1.PortName = "COM" + Convert.ToString(comboBox22.SelectedIndex);
                    serialPort1.Open();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Port opening", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            serialPort1.Write("STP\r\n");
            System.Threading.Thread.Sleep(20);
            serialPort1.Write("DIR WDR\r\n");
            System.Threading.Thread.Sleep(20);
            serialPort1.Write("RAT 200 MH\r\n"); //Rate fixed
            System.Threading.Thread.Sleep(20);
            serialPort1.Write("RUN\r\n");
            System.Threading.Thread.Sleep(20);
        }
        private void button16_MouseUp(object sender, MouseEventArgs e)
        {
            _Reading = serialPort1.ReadExisting();
            serialPort1.Write("DIS\r\n");
            System.Threading.Thread.Sleep(20);
            _Reading = serialPort1.ReadExisting();
            label46.Text = String.Format("{0:0.0000}", (Convert.ToDouble(_Reading.Substring(11, 5)))) + " "+ _Reading.Substring(16, 2).ToLower();//_Reading.Substring(11, 5).ToLower() + " " + _Reading.Substring(16, 2).ToLower();
            serialPort1.Write("STP\r\n");
            System.Threading.Thread.Sleep(20);
            serialPort1.Close();
        }
        //Clear Values manual infusion
        private void button18_Click(object sender, EventArgs e)
        {
            try
            {
                if (!serialPort1.IsOpen)
                {
                    serialPort1.PortName = "COM" + Convert.ToString(comboBox22.SelectedIndex);
                    serialPort1.Open();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Port opening", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            label44.Text="0 ul";
            label46.Text = "0 ul";
            serialPort1.Write("STP\r\n");
            System.Threading.Thread.Sleep(20);
            serialPort1.Write("CLD INF\r\n");
            System.Threading.Thread.Sleep(20);
            serialPort1.Write("CLD WDR\r\n");
            serialPort1.Close();
        }

        //*********************************************************************************
        //***********************************END*******************************************

        //*********************************************************************************
        //******************************STATIC EXPERIMENT**********************************

        private bool CheckExpErr() //Check errors
        {
            try
            {
                if (Convert.ToDouble(textBox2.Text) < 0 || Convert.ToDouble(textBox6.Text) < 5)
                    throw new ArgumentException();
                if (Convert.ToDouble(textBox5.Text) < 20 || Convert.ToDouble(textBox5.Text) > 500)
                    throw new ArgumentException();
            }
            catch (Exception)
            {
                MessageBox.Show("Volume or time are below or over the limits!","Error!",MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            try
            {
                if (Convert.ToInt16(textBox4.Text) < 0)
                    throw new ArgumentException();
            }
            catch (Exception)
            {
                MessageBox.Show("Initial time", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            if (checkBox1.Checked != true || checkBox2.Checked != true)
            {
                MessageBox.Show("Calibrations not done, please load the calibration files!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            if (comboBox1.SelectedIndex==-1)
            {
                MessageBox.Show("Please choose a sampling rate!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            if (comboBox3.SelectedIndex == -1)
            {
                MessageBox.Show("Please choose a type of deviation!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            return true;
        }
        private void button14_Click(object sender, EventArgs e)
        {
            if (Paused) //If paused and pressed resume
            {
                if (CheckExpErr())
                {
                    button13.Enabled = true;
                    button15.Enabled = true;
                    button14.Enabled = false;
                    Paused = false;
                    //Change text of the button
                    button14.Text = "Start";
                    try
                    {
                        if (!serialPort1.IsOpen)
                        {
                            serialPort1.PortName = "COM" + Convert.ToString(comboBox22.SelectedIndex);
                            serialPort1.Open();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Port opening", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LJUD.ePut(u3.ljhandle, LJUD.IO.SWDT_CONFIG, LJUD.CHANNEL.SWDT_DISABLE, 0, 0); 
                        return;
                    }
                    //Clear Values
                    ExpState = 0;
                    ExpTicks = 0;
                    ExpAccVal = 0;
                    Inf_Ticks = 0;
                    //Set the controls to disable and the tabs
                    textBox2.Enabled = false;
                    textBox5.Enabled = false;
                    textBox6.Enabled = false;
                    textBox4.Enabled = false;
                    EnableTab(tabPage1, false);
                    EnableTab(tabPage2, false);
                    EnableTab(tabPage3, false);
                    //enable the timer again for showing the graph
                    LJUD.AddRequest(u3.ljhandle, LJUD.IO.PUT_CONFIG, LJUD.CHANNEL.SWDT_ENABLE, 10, 0, 0);
                    LJUD.ePut(u3.ljhandle, LJUD.IO.PUT_ANALOG_ENABLE_PORT, 0, 31, 16);//first 5 FIO analog b0000000000011111
                    LJUD.AddRequest(u3.ljhandle, LJUD.IO.GET_AIN_DIFF, 4, 0, 32, 0);//Request FIO4
                    LJUD.GoOne(u3.ljhandle);
                    timer4.Enabled = true;
                    timer2.Enabled = false;
                    aTimer.Enabled = true;
                }
                return;
            }
            //If it is the start, Check for errors
            if (!CheckExpErr())
            {
                if (serialPort1.IsOpen)
                    serialPort1.Close();
                return;
            }
            try
            {
                if (!serialPort1.IsOpen)
                {
                    serialPort1.PortName = "COM" + Convert.ToString(comboBox22.SelectedIndex);
                    serialPort1.Open();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Port opening", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if(u3 != null)
                    LJUD.ePut(u3.ljhandle, LJUD.IO.SWDT_CONFIG, LJUD.CHANNEL.SWDT_DISABLE, 0, 0); 
                return;
            }
            //Send commands to the pump
            serialPort1.Write("STP\r\n");
            System.Threading.Thread.Sleep(20);
            serialPort1.Write("CLD INF\r\n");
            System.Threading.Thread.Sleep(20);
            serialPort1.Write("CLD WDR\r\n");
            System.Threading.Thread.Sleep(20);
            serialPort1.Write("DIR INF\r\n");
            System.Threading.Thread.Sleep(20);
            serialPort1.Write("DIA " + textBox21.Text + "\r\n");
            System.Threading.Thread.Sleep(20);
            serialPort1.Write("RAT 800 MH\r\n"); //Rate fixed
            System.Threading.Thread.Sleep(20);
            serialPort1.Write("VOL UL\r\n");
            System.Threading.Thread.Sleep(20);
            //Clear the accum
            AccumVolInf = 0;
            dblValueAcc = 0;
            dblValue = 0;
            ExpState = 0;
            ExpTicks = 0;
            ExpAccVal = 0;
            dblValue = 0;
            dblValueAcc = 0;
            WdrVol = 1000;
            //Check the subsampling rate and load the divider for the graph
            try
            {
                switch (comboBox1.SelectedIndex)
                {
                    case 0:
                        SubSampling = 500;
                        GraphPt = 10; //Every 5 secs, refresh graph
                        break;
                    case 1:
                        SubSampling = 1000;
                        GraphPt = 5;
                        break;
                    case 2:
                        SubSampling = 2000;
                        GraphPt = 3;
                        break;
                    case 3:
                        SubSampling = 3000;
                        GraphPt = 2;
                        break;
                    case 4:
                        SubSampling = 4000;
                        GraphPt = 1;
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Sumbsampling rate", MessageBoxButtons.OK, MessageBoxIcon.Error);
                serialPort1.Close();
                if (u3 != null)
                    LJUD.ePut(u3.ljhandle, LJUD.IO.SWDT_CONFIG, LJUD.CHANNEL.SWDT_DISABLE, 0, 0); 
                return;
            }
            //Configure DAQ
            try
            {
                if (u3 == null)
                    u3 = new U3(LJUD.CONNECTION.USB, "1", true); // Connection through USB
                ConfigLJ();
            }
            catch (LabJackUDException h)
            {
                MessageBox.Show("Error opening DAQ. "+ h.ToString(), "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                serialPort1.Close();
                if (u3 != null)
                    LJUD.ePut(u3.ljhandle, LJUD.IO.SWDT_CONFIG, LJUD.CHANNEL.SWDT_DISABLE, 0, 0); 
                return;
            }
            //Refresh the values of slopes
            label23.Text = String.Format("{0:0.0000}", pHCalSlope);
            label35.Text = String.Format("{0:0.0000}", SyrCalSlope);
            //Format the chart axis
            chart4.ChartAreas[0].AxisY.LabelStyle.Format = "#.###";
            chart4.ChartAreas[0].AxisX.LabelStyle.Format = "#";
            //Activate controls
            button13.Enabled = true;
            button15.Enabled = true;
            button14.Enabled = false;
            textBox2.Enabled = false;
            textBox5.Enabled = false;
            textBox6.Enabled = false;
            textBox4.Enabled = false;
            comboBox1.Enabled = false;
            checkBox5.Enabled = false;
            comboBox3.Enabled = false;
            //Clear Graph
            foreach (var series in chart4.Series)
                series.Points.Clear();
            //Set initial state
            ExpState = 0;
            // Create a timer with an interval.
            aTimer = new System.Timers.Timer(SubSampling);
            // Hook up the Elapsed event for the timer. 
            aTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent);
            //Start Export Data
            string _FirstLine = "PH Setp.:"+ textBox2.Text + "Max. Vol [ul]:" + textBox5.Text + "Mix Time [s]:" + textBox6.Text +",Time[s],Value,,VOLUME,Time[s],Volume[ul],,DEVIATION,Time[s],Value,\n";
            string[] Content = new string[chart4.Series["Series1"].Points.Count];
            string Path = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)+ @"\Calcification Experiments";
            try
            {
                if (!System.IO.Directory.Exists(Path))
                    System.IO.Directory.CreateDirectory(Path);
                swexp = new StreamWriter(Path+@"\"+DateTime.Now.ToString("yyyy-MM-dd HHmmss")+".csv");
                swexp.Write(_FirstLine);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error Writing File", MessageBoxButtons.OK, MessageBoxIcon.Error);
                LJUD.ePut(u3.ljhandle, LJUD.IO.SWDT_CONFIG, LJUD.CHANNEL.SWDT_DISABLE, 0, 0); 
                return;
            } 
            ExpStart = DateTime.Now;
            //Disable all other timers that might be enabled
            timer1.Enabled = false;
            timer2.Enabled = false;
            timer3.Enabled = false;
            textBox3.Enabled = true;
            //Disable other tabs
            EnableTab(tabPage1, false);
            EnableTab(tabPage2, false);
            EnableTab(tabPage3, false);
            //Enable Graph and reading of DAQ
            aTimer.Enabled = true;
            timer4.Enabled = true;
        }
        //Reading of DAQ
        private void timer4_Tick(object sender, EventArgs e)
        {
            timer4.Enabled = false;
            bool requestedExit = false;
            while (!requestedExit)
            {
                try
                {
                    LJUD.GoOne(u3.ljhandle);
                    LJUD.GetFirstResult(u3.ljhandle, ref ioType, ref channel, ref dblValue, ref dummyInt, ref dummyDouble);
                }
                catch (LabJackUDException h)
                {
                    //MessageBox.Show("Error getting the DAQ data. " + h.ToString(), "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UpdateStatStrip(h, true);
                    System.Threading.Thread.Sleep(10000);
                    ConfigLJ();

                }
                if (ioType == LJUD.IO.GET_AIN_DIFF)
                {
                    switch (ExpState)
                    { 
                        case 0:
                            //Initial state, accum the reading
                            ExpTicks++;
                            ExpAccVal += (dblValue - pHCalIntercept) / pHCalSlope;
                            dblValueAcc += dblValue;
                            //If the ticks are 10, change state
                            if (ExpTicks >= 10)
                                ExpState = 1;
                            break;
                        case 1:
                            //Calculate the avgs
                            ExpAvgVal = ExpAccVal / 10;
                            AvgVoltage = dblValueAcc / 10;
                            //Calculate the deviation
                            Deviation = ExpAvgVal - Convert.ToDouble(textBox2.Text);
                            //Reset Ticks and accum.
                            ExpTicks = 0;
                            ExpAccVal = 0;
                            dblValueAcc = 0;
                            //If the initial time is not reached go to state 0
                            if (DateTime.Now.AddSeconds(0 - Convert.ToDouble(textBox4.Text)) < ExpStart)
                            {
                                ExpState = 0;
                                break;
                            }
                            //If the pH value is over the setpoint and positive infusion was selected -> infuse
                            if (comboBox3.SelectedIndex == 0) // Positive Deviation Infusion
                            {
                                if (Deviation > 0 && InfStarted == false && TimeMix == false && TimeMix_Wdr == false)
                                {

                                    ExpState = 2;
                                    serialPort1.Write("CLD INF\r\n");
                                    System.Threading.Thread.Sleep(20);
                                    serialPort1.Write("CLD WDR\r\n");
                                    System.Threading.Thread.Sleep(20);
                                    break;
                                }
                            }
                            //If the pH value is under the setpoint and negative infusion was selected -> infuse
                            else //Negative Deviation
                            {
                                if (Deviation < 0 && InfStarted == false && TimeMix == false && TimeMix_Wdr == false)
                                {
                                    ExpState = 2;
                                    serialPort1.Write("CLD INF\r\n");
                                    System.Threading.Thread.Sleep(20);
                                    serialPort1.Write("CLD WDR\r\n");
                                    System.Threading.Thread.Sleep(20);
                                    break;
                                }
                            }
                            //Check if it is infusing
                            if (InfStarted)
                            {
                                ExpState = 3;
                                break;
                            }
                            //Timemix
                            if (TimeMix)
                            {
                                if (DateTime.Now.AddSeconds(-Convert.ToDouble(textBox6.Text)) > ExpWaitTime)
                                    TimeMix = false;
                                ExpState = 0;
                                break;
                            }
                            if (TimeMix_Wdr)
                            {
                                if (DateTime.Now.AddSeconds(-6) > ExpWdrTime)
                                    TimeMix_Wdr = false;
                                ExpState = 0;
                                break;
                            }
                            ExpState = 0;
                            break;
                            ////////////////////////////////////////
                        case 2:
                            //Infusion starts
                            InfStarted = true;
                            //Calculate the volume to infuse
                            CalcVol = Math.Abs(Deviation) * (Convert.ToDouble(textBox5.Text) - 20) + 20;
                            VoltoInf = (CalcVol - SyrCalIntercept) / SyrCalSlope;
                            //Minimum: 20ul
                            if (VoltoInf < 20)
                            {
                                VoltoInf = (20 - SyrCalIntercept) / SyrCalSlope;
                                CalcVol = 20;
                            }
                            //Max: Set
                            if (VoltoInf > Convert.ToDouble(textBox5.Text))
                            {
                                VoltoInf = (Convert.ToDouble(textBox5.Text) - SyrCalIntercept) / SyrCalSlope;
                                CalcVol = Convert.ToDouble(textBox5.Text);
                            }
                            //Set the pump
                            serialPort1.Write("DIR INF\r\n");
                            System.Threading.Thread.Sleep(20);
                            serialPort1.Write("VOL " + String.Format("{0:000.0}", VoltoInf) + "\r\n");
                            //Accum the vol
                            AccumVolInf += CalcVol;
                            //Increase the ticks
                            Inf_Ticks++;
                            System.Threading.Thread.Sleep(20);
                            //Set the timer to now to check if some secs passed
                            ExpWaitTime = DateTime.Now;
                            serialPort1.Write("RUN\r\n");
                            System.Threading.Thread.Sleep(20);
                            ExpState=0;
                            //Go back to 0 even it is infusing so it can read pH in the meanwhile
                            break;
                        case 3:
                            try
                            {
                                //Read the value infused
                                _Reading = serialPort1.ReadExisting();
                                serialPort1.Write("DIS\r\n");
                                System.Threading.Thread.Sleep(20);
                                _Reading = serialPort1.ReadExisting();
                                //Check if the read volume but corrected is bigger than the volume to infuse or 3 secs passed
                                ReadVol = Convert.ToDouble(_Reading.Substring(5, 5));
                                if (ReadVol * SyrCalSlope + SyrCalIntercept >= VoltoInf || DateTime.Now.AddSeconds(-3) > ExpWaitTime) //up to 3 senconds to reach
                                {
                                    InfStarted = false; //Infusion finished, check if 1ml was infused and recharge
                                    if (AccumVolInf > WdrVol && checkBox5.Checked == true)
                                    {
                                        //Increment the Wdr
                                        WdrVol = WdrVol + 1000;
                                        serialPort1.Write("DIR WDR\r\n");
                                        System.Threading.Thread.Sleep(30);
                                        //Write to the syringe

                                        //BUG********************************
                                        //serialPort1.Write("VOL " + String.Format("{0:0000}", (1000 - SyrCalIntercept) / SyrCalSlope) + "\r\n");
                                        serialPort1.Write("VOL " + String.Format("{0:0000}", 1000 / SyrCalSlope - Inf_Ticks * SyrCalIntercept / SyrCalSlope) + "\r\n");
                                        Inf_Ticks = 0;

                                        System.Threading.Thread.Sleep(30);
                                        serialPort1.Write("RUN\r\n");
                                        System.Threading.Thread.Sleep(30);
                                        //Wdr Mix time
                                        TimeMix_Wdr = true;
                                        ExpWdrTime = DateTime.Now;
                                    }
                                    //Mixing time
                                    TimeMix = true;
                                }
                                //Back to read pH
                                ExpState = 0;
                            }
                            catch (Exception h)
                            {
                                UpdateStatStrip(h, true);
                                try
                                {
                                    serialPort1.Close();
                                    serialPort1.PortName = "COM" + Convert.ToString(comboBox22.SelectedIndex);
                                    serialPort1.Open();
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.Message, "Port opening", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    if (u3 != null)
                                        LJUD.ePut(u3.ljhandle, LJUD.IO.SWDT_CONFIG, LJUD.CHANNEL.SWDT_DISABLE, 0, 0);
                                    return;
                                }
                                //Send commands to the pump
                                serialPort1.Write("STP\r\n");
                                System.Threading.Thread.Sleep(20);
                                serialPort1.Write("CLD INF\r\n");
                                System.Threading.Thread.Sleep(20);
                                serialPort1.Write("CLD WDR\r\n");
                                System.Threading.Thread.Sleep(20);
                                serialPort1.Write("DIR INF\r\n");
                                System.Threading.Thread.Sleep(20);
                                serialPort1.Write("DIA " + textBox21.Text + "\r\n");
                                System.Threading.Thread.Sleep(20);
                                serialPort1.Write("RAT 800 MH\r\n"); //Rate fixed
                                System.Threading.Thread.Sleep(20);
                                serialPort1.Write("VOL UL\r\n");
                                System.Threading.Thread.Sleep(20);
                                ExpState = 0;
                            }
                            break;
                    }
                }
                try
                {
                    LJUD.GetNextResult(u3.ljhandle, ref ioType, ref channel, ref dblValue, ref dummyInt, ref dummyDouble);
                }
                catch (LabJackUDException h)
                {
                    if (h.LJUDError == U3.LJUDERROR.NO_MORE_DATA_AVAILABLE)
                        requestedExit = true;//no more data to read
                    else
                    {
                        //MessageBox.Show("Error getting DAQ data. " + h.ToString(), "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        UpdateStatStrip(h, true);
                        System.Threading.Thread.Sleep(10000);
                        ConfigLJ();
                    }
                }
            }
            //Enable timer again
            timer4.Enabled = true;
        }

        public void ConfigLJ()
        {
            LJUD.AddRequest(u3.ljhandle, LJUD.IO.SWDT_CONFIG, LJUD.CHANNEL.SWDT_ENABLE, 10, 0, 0);
            LJUD.AddRequest(u3.ljhandle, LJUD.IO.PUT_CONFIG, LJUD.CHANNEL.SWDT_RESET_DEVICE, 1, 0, 0);
            LJUD.ePut(u3.ljhandle, LJUD.IO.PIN_CONFIGURATION_RESET, 0, 0, 0);
            LJUD.ePut(u3.ljhandle, LJUD.IO.PUT_ANALOG_ENABLE_PORT, 0, 31, 16);//first 5 FIO analog b0000000000011111
            LJUD.AddRequest(u3.ljhandle, LJUD.IO.GET_AIN_DIFF, 4, 0, 32, 0);//Request FIO4
            LJUD.GoOne(u3.ljhandle);
        }

        //Refresh Graph
        private void OnTimedEvent(Object source, ElapsedEventArgs e)
        {
            if (InvokeRequired)
            {
                Invoke(new MethodInvoker(delegate
                    {
                        //Calculate the time from the beggining of the experiment
                        TimeDif = (DateTime.Now - ExpStart).TotalSeconds;
                        chart4.Series["Series1"].Points.AddXY(TimeDif, ExpAvgVal);
                        chart4.Series["Series2"].Points.AddXY(TimeDif, AccumVolInf);
                        chart4.Series["Series3"].Points.AddXY(TimeDif, Deviation);
                        swexp.Write("," + TimeDif + "," + ExpAvgVal + ",,," + TimeDif + "," + AccumVolInf + ",,," + TimeDif + "," + Deviation + "\n");
                        swexp.Flush();
                        //Update if the pointer is divisible by the update
                        graphUpdate++;
                        if (graphUpdate % GraphPt == 0)
                        {
                            chart4.Series.ResumeUpdates();
                            chart4.Series.Invalidate();
                            chart4.Series.SuspendUpdates();
                            graphUpdate = 0;
                        }
                        //refresh the labels
                        label13.Text = String.Format("{0:0.0000 V}", AvgVoltage);
                        label16.Text = String.Format("{0:0.000}", ExpAvgVal);                     
                        label17.Text = String.Format("{0:0.000}", Deviation);
                        label21.Text = String.Format("{0}", (DateTime.Now - ExpStart).Days)+ " days " + String.Format("{0:00}", (DateTime.Now - ExpStart).Hours) + ":" + String.Format("{0:00}", (DateTime.Now - ExpStart).Minutes) + ":" + String.Format("{0:00}", (DateTime.Now - ExpStart).Seconds);
                        label19.Text = String.Format("{0:00000.00}", AccumVolInf) + " ul";
                    }));
            }
        }
        //Change Graph
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox2.SelectedIndex)
            {
                case -1:
                    chart4.Series["Series1"].Enabled = true;
                    chart4.Series["Series2"].Enabled = false;
                    chart4.Series["Series3"].Enabled = false;
                    chart4.ChartAreas[0].AxisY.Title = "pH";
                    chart4.Update();
                    break;
                case 0:
                    chart4.Series["Series1"].Enabled = true;
                    chart4.Series["Series2"].Enabled = false;
                    chart4.Series["Series3"].Enabled = false;
                    chart4.ChartAreas[0].AxisY.Title = "pH";
                    chart4.Update();
                    break;
                case 1:
                    chart4.Series["Series1"].Enabled = false;
                    chart4.Series["Series2"].Enabled = true;
                    chart4.Series["Series3"].Enabled = false;
                    chart4.ChartAreas[0].AxisY.Title = "Volume [ul]";
                    chart4.Update();
                    break;
                case 2:
                    chart4.Series["Series1"].Enabled = false;
                    chart4.Series["Series2"].Enabled = false;
                    chart4.Series["Series3"].Enabled = true;
                    chart4.ChartAreas[0].AxisY.Title = "pH Deviation";
                    chart4.Update();
                    break;
            }
        }
        //Pause
        private void button13_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure you want to Pause?", "Pause?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;
            Paused = true;
            //disable timers
            timer4.Enabled = false;
            ExpState = 0;
            aTimer.Enabled = false;
            //Disable the watchdog
            if (u3 != null)
                LJUD.ePut(u3.ljhandle, LJUD.IO.SWDT_CONFIG, LJUD.CHANNEL.SWDT_DISABLE, 0, 0); 
            //Change controls and enable tabs
            button14.Text = "Resume";
            button13.Enabled = false;
            button15.Enabled = false;
            button14.Enabled = true;
            textBox2.Enabled = false;
            textBox4.Enabled = false;
            panel14.Enabled = false;
            EnableTab(tabPage1, true);
            EnableTab(tabPage2, true);
            EnableTab(tabPage3, true);
            button6.Enabled = true;
            button7.Enabled = false;
            textBox3.Enabled = true;
            
        }

        //Disable the watchdog.

        //Stop the experiment
        private void button15_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure you want to stop?", "Stop?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;
            //Disable timers
            timer4.Enabled = false;
            aTimer.Enabled = false;
            //Close COM
            serialPort1.Close();
            swexp.Close();
            //Change controls
            button13.Enabled = false;
            button15.Enabled = false;
            button14.Enabled = true;
            textBox2.Enabled = true;
            textBox4.Enabled = true;
            comboBox1.Enabled = true;
            checkBox5.Enabled = true;
            comboBox3.Enabled = true;
            textBox5.Enabled = true;
            textBox6.Enabled = true;
            panel14.Enabled = true;
            Paused = false;
            EnableTab(tabPage1, true);
            EnableTab(tabPage2, true);
            EnableTab(tabPage3, true);
            //Set variables to 0
            ExpTicks = 0;
            ExpAvgVal = 0;
            ExpAccVal = 0;
            Deviation=0;
            VoltoInf=0;
            AvgVoltage = 0;
            ExpTicks = 0;
            dblValueAcc = 0;
            if (u3 != null)
                LJUD.ePut(u3.ljhandle, LJUD.IO.SWDT_CONFIG, LJUD.CHANNEL.SWDT_DISABLE, 0, 0); 
        }

        //*********************************************************************************
        //***********************************CLOSING***************************************

        private void TECAS_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (MessageBox.Show("Are you sure you want to close?", "Quit?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                e.Cancel = true;
                return;
            }
            if (serialPort1.IsOpen)
                serialPort1.Close();
            if (u3 != null)
                LJUD.ePut(u3.ljhandle, LJUD.IO.SWDT_CONFIG, LJUD.CHANNEL.SWDT_DISABLE, 0, 0);
            if(!newErrorLog)
                Error_SW.Close();
        }

        //*********************************************************************************
        //***********************************ZOOM******************************************

        private void chData_MouseWheel1(object sender, MouseEventArgs e)
        {
            try
            {
                xMin = chart4.ChartAreas[0].AxisX.ScaleView.ViewMinimum;
                xMax = chart4.ChartAreas[0].AxisX.ScaleView.ViewMaximum;
                yMin = chart4.ChartAreas[0].AxisY.ScaleView.ViewMinimum;
                yMax = chart4.ChartAreas[0].AxisY.ScaleView.ViewMaximum;

                if (e.Delta < 0)
                {
                    posXStart = chart4.ChartAreas[0].AxisX.PixelPositionToValue(e.Location.X) - (xMax - xMin) * 4;
                    posXFinish = chart4.ChartAreas[0].AxisX.PixelPositionToValue(e.Location.X) + (xMax - xMin) * 4;
                    posYStart = chart4.ChartAreas[0].AxisY.PixelPositionToValue(e.Location.Y) - (yMax - yMin) * 4;
                    posYFinish = chart4.ChartAreas[0].AxisY.PixelPositionToValue(e.Location.Y) + (yMax - yMin) * 4;
                    chart4.ChartAreas[0].AxisX.ScaleView.Zoom(posXStart, posXFinish);
                    chart4.ChartAreas[0].AxisY.ScaleView.Zoom(posYStart, posYFinish);
                }

                if (e.Delta > 0)
                {
                    posXStart = chart4.ChartAreas[0].AxisX.PixelPositionToValue(e.Location.X) - (xMax - xMin) / 4;
                    posXFinish = chart4.ChartAreas[0].AxisX.PixelPositionToValue(e.Location.X) + (xMax - xMin) / 4;
                    posYStart = chart4.ChartAreas[0].AxisY.PixelPositionToValue(e.Location.Y) - (yMax - yMin) / 4;
                    posYFinish = chart4.ChartAreas[0].AxisY.PixelPositionToValue(e.Location.Y) + (yMax - yMin) / 4;
                    chart4.ChartAreas[0].AxisX.ScaleView.Zoom(posXStart, posXFinish);
                    chart4.ChartAreas[0].AxisY.ScaleView.Zoom(posYStart, posYFinish);
                }
            }
            catch { }
        }

        private void chart4_MouseEnter(object sender, EventArgs e)
        {
            chart4.MouseWheel += new System.Windows.Forms.MouseEventHandler(this.chData_MouseWheel1);
            chart4.Focus();
        }

        private void chart4_MouseLeave(object sender, EventArgs e)
        {
            chart4.MouseWheel -= new System.Windows.Forms.MouseEventHandler(this.chData_MouseWheel1);
            chart4.Focus();
            chart4.ChartAreas[0].AxisX.ScaleView.ZoomReset(0);
            chart4.ChartAreas[0].AxisY.ScaleView.ZoomReset(0);
        }

        private void TECAS_Load(object sender, EventArgs e)
        {
            chart4.Series.SuspendUpdates();
        }
        //Enable tab function
        public static void EnableTab(TabPage page, bool enable)
        {
            foreach (Control ctl in page.Controls) ctl.Enabled = enable;
        }

        private void tabControl1_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (timer4.Enabled)
            {
                tabControl1.SelectTab(tabPage4);
            }
            return;
        }

        private void button21_Click(object sender, EventArgs e)
        {
            try
            {
                if (u3 == null)
                    u3 = new U3(LJUD.CONNECTION.USB, "1", true); // Connection through USB
                LJUD.ePut(u3.ljhandle, LJUD.IO.SWDT_CONFIG, LJUD.CHANNEL.SWDT_DISABLE, 0, 0);
                MessageBox.Show("WD Disabled!", "Watchdog", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            catch (Exception h)
            {
                MessageBox.Show("LJ not connected: " + h.Message, "Watchdog", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        //*********************************************************************************
        //***********************************END*******************************************
    }

}

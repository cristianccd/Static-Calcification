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
        bool pHCal = false;

        //pH measurement
        double pHMeasureAvg = 0, pHMeasureTicks = 0, pHMeasureVal = 0;
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
        double dblValue = 0,dummyDouble = 0;
        int dummyInt = 0;

        //Experiment
        double ExpTicks = 0, ExpAvgVal = 0, ExpAccVal = 0, Deviation=0, VoltoInf=0, TimeDif=0;
        DateTime ExpStart, ExpWaitTime;
        bool InfStarted = false;
        bool Paused = false;
        int ExpState = 0;
        double AccumVolInf = 0, SubSampling=0;
        static private System.Timers.Timer aTimer;

        //***************************

        public TECAS()
        {
            InitializeComponent();
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
        private void button10_Click(object sender, EventArgs e)
        {
            if(pHCalErr())
            {
                //Open labjack and catch exception if not able to open 
                try
                {
                    if (u3 == null)
                        u3 = new U3(LJUD.CONNECTION.USB, "1", true); // Connection through USB
                    LJUD.ePut(u3.ljhandle, LJUD.IO.PIN_CONFIGURATION_RESET, 0, 0, 0);
                    LJUD.ePut(u3.ljhandle, LJUD.IO.PUT_ANALOG_ENABLE_PORT, 0, 15, 16);//first 4 FIO analog b0000000000001111
                    LJUD.AddRequest(u3.ljhandle, LJUD.IO.GET_AIN, 0, 0, 0, 0);//Request AIN0
                }
                catch (LabJackUDException)
                {
                    MessageBox.Show("Error opening DAQ", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                foreach (var series in chart3.Series)
                    series.Points.Clear();
                //Block pH controls
                textBox1.Enabled = false;
                button8.Enabled = false;
                if (dataGridView2.RowCount <= 2)
                    button9.Enabled = false;
                button10.Enabled = false;
                button11.Enabled = false;
                button12.Enabled = false;
                label48.Visible = true;
                button19.Enabled = true;
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
                pHCal = false;
                panel1.Visible = false;
                for (int i = 0; i < dataGridView2.RowCount; i++)
                    dataGridView2[1, i].Value = "";
                timer3.Enabled = true;
            }
            else
                return;
        }

    private void timer3_Tick(object sender, EventArgs e)
        {
            timer3.Enabled = false; //first disable the timer to perform actions
            bool requestedExit = false;
            //blink label Please Wait...
            Waitlbltick++;
            if (Waitlbltick >= 5)
            {
                Waitlbltick = 0;
                label48.Visible = !label48.Visible;
            }
            //Start of calibration
            switch (pHCalState)
            {
                case 0://calibration started
                    pHQuestSample = MessageBox.Show("When you are ready to measure the Sample " + Convert.ToString(pHSampleNo + 1) + " press OK", "pH Meter Calibration", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                    //if the result is OK, change state and start counting time
                    if (pHQuestSample == DialogResult.OK)
                    {
                        pHCalState = 1;
                        pHstartTime = DateTime.Now;
                    }
                    if (pHQuestSample == DialogResult.Cancel)
                    {
                        //Enable pH controls
                        CancelpHTest();
                        return;
                    }
                    timer3.Enabled = true;
                    break;
                case 1://Start accumulating time and check if the target is reached
                    label12.Text = String.Format("{0:00}", (DateTime.Now - pHstartTime).Minutes) + ":" + String.Format("{0:00}", (DateTime.Now - pHstartTime).Seconds);
                    if (Convert.ToDouble((DateTime.Now - pHstartTime).TotalMinutes) >= Convert.ToDouble(textBox1.Text))
                    {
                        pHCalState = 2; //Time reached
                        timer3.Enabled = true;
                        break;
                    }
                    pHTicks++;
                    SumpHCal = SumpHCal + dblValue;
                    timer3.Enabled = true;
                    break;
                case 2://Time reached, change sample or exit
                    pHAvgVal = SumpHCal / pHTicks;
                    dataGridView2[1, pHSampleNo].Value = String.Format("{0:0.0000}", pHAvgVal); ;
                    //reset values
                    pHSampleNo++;
                    pHTicks = 0;
                    SumpHCal = 0;
                    pHCalState = 0;
                    if (pHSampleNo >= dataGridView2.RowCount)
                        //end of calibration
                        pHCalState = 10;
                    timer3.Enabled = true;
                    break;
                case 10://Exit: calculate the results
                    MessageBox.Show("Finished, please check the results...", "Finished", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //auxiliar variables to obtain results
                    double a = 0, b = 0, c = 0, d = 0, f = 0;
                    for (int i = 0; i < dataGridView2.RowCount; i++)
                    {
                        //sumxy
                        a= a + Convert.ToDouble(dataGridView2[0, i].Value) * Convert.ToDouble(dataGridView2[1, i].Value);
                        //sumx
                        b= b + Convert.ToDouble(dataGridView2[1, i].Value);
                        //sumy
                        c= c +Convert.ToDouble(dataGridView2[0, i].Value);
                        //sumx2
                        d= d + Math.Pow(Convert.ToDouble(dataGridView2[1, i].Value), 2);
                        //sumy2
                        f= f + Math.Pow(Convert.ToDouble(dataGridView2[0, i].Value),2);
                        chart3.Series["Series2"].Points.AddXY(Convert.ToDouble(dataGridView2[1,i].Value),dataGridView2[0, i].Value);
                    }
                    /*R2 = ((Nxysum - xsumysum)^2)/((Nx^2sum - xsum*xsum)*(Ny^2sum - ysum*ysum))*/
                    pHR2=(Math.Pow((dataGridView2.RowCount*a-b*c),2))/((dataGridView2.RowCount*d-Math.Pow(b,2))*(dataGridView2.RowCount*f-Math.Pow(c,2)));
                    /*Slope(b) = (NΣXY - (ΣX)(ΣY)) / (NΣX2 - (ΣX)2) Intercept(a) = (ΣY - b(ΣX)) / N */
                    pHCalSlope = (dataGridView2.RowCount * a - b * c) / (dataGridView2.RowCount*d - Math.Pow(b, 2));
                    pHCalIntercept=(c-pHCalSlope*b)/dataGridView2.RowCount;
                    if(pHCalIntercept>=0)
                        label6.Text = "y=" + String.Format("{0:0.0000}", pHCalSlope) + " x+" + String.Format("{0:0.0000}", pHCalIntercept);
                    else
                        label6.Text = "y=" + String.Format("{0:0.0000}", pHCalSlope) + " x" + String.Format("{0:0.0000}", pHCalIntercept);
                    label5.Text = "R =  " + String.Format("{0:0.0000}", pHR2);
                    //Graph
                    chart3.Series["Series1"].Points.AddXY(Convert.ToDouble(dataGridView2[1,0].Value),pHCalSlope*Convert.ToDouble(dataGridView2[1,0].Value)+pHCalIntercept);
                    chart3.Series["Series1"].Points.AddXY(Convert.ToDouble(dataGridView2[1, dataGridView2.RowCount-1].Value), pHCalSlope * Convert.ToDouble(dataGridView2[1, dataGridView2.RowCount - 1].Value) + pHCalIntercept);
                    panel1.Visible = true;
                    button8.Enabled = true;
                    if (dataGridView2.RowCount <= 2)
                        button9.Enabled = false;
                    button10.Enabled = true;
                    button11.Enabled = true;
                    button12.Enabled = true;
                    pHTicks = 0;
                    SumpHCal = 0;
                    pHSampleNo = 0;
                    pHCalState = 0;
                    pHCal = true;
                    button6.Enabled = true;
                    checkBox3.Checked = true;
                    checkBox2.Checked = true;
                    label48.Visible = false;
                    button19.Enabled = false;
                    return;
            }
            while (!requestedExit)
            {
                try
                {
                    LJUD.GoOne(u3.ljhandle);
                    LJUD.GetFirstResult(u3.ljhandle, ref ioType, ref channel, ref dblValue, ref dummyInt, ref dummyDouble);
                }
                catch (LabJackUDException)
                {
                    MessageBox.Show("Error getting the DAQ results", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                if (ioType == LJUD.IO.GET_AIN)
                    label10.Text = String.Format("{0:0.00000}", dblValue);
                try
                {
                    LJUD.GetNextResult(u3.ljhandle, ref ioType, ref channel, ref dblValue, ref dummyInt, ref dummyDouble);
                }
                catch (LabJackUDException h)
                {
                    if (h.LJUDError == U3.LJUDERROR.NO_MORE_DATA_AVAILABLE)
                        requestedExit = true;//no more data to read
                    else
                        MessageBox.Show("Error getting DAQ data", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

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
                        label6.Text = "y="+pHCalSlope.ToString()+" x+"+pHCalIntercept.ToString();
                    else
                        label6.Text = "y=" + pHCalSlope.ToString() + " x" + pHCalIntercept.ToString();
                    label5.Text = "R  =" + pHR2;
                    panel1.Visible = true;
                    chart3.Series["Series1"].Points.AddXY(0, pHCalIntercept);
                    chart3.Series["Series1"].Points.AddXY(14, pHCalSlope * 14 + pHCalIntercept);
                    button12.Enabled = true;
                    pHCal = true;
                    button6.Enabled = true;
                    checkBox3.Checked = true;
                    checkBox2.Checked = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Calibration", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            try
            {
                if (saveFileDialog2.ShowDialog() == DialogResult.OK)
                {
                    StreamWriter sw = new StreamWriter(saveFileDialog2.FileName);
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
        }

        private void CancelpHTest()
        {
            timer3.Enabled = false;
            textBox1.Enabled = true;
            button8.Enabled = true;
            button9.Enabled = false;
            button10.Enabled = true;
            button11.Enabled = true;
            button12.Enabled = false;
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
            label10.Text = "0.00";
            label12.Text = "00:00";
            pHCal = false;
            dataGridView2.Rows.Clear();
            dataGridView2.Rows.Add(2);
        }
        //*********************************************************************************
        //***********************************END*******************************************

        //*********************************************************************************
        //******************************PH MEASUREMENT*************************************

        private void button6_Click(object sender, EventArgs e)
        {
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
            //Configure DAQ
            try
            {
                if (u3 == null)
                    u3 = new U3(LJUD.CONNECTION.USB, "1", true); // Connection through USB
                LJUD.ePut(u3.ljhandle, LJUD.IO.PIN_CONFIGURATION_RESET, 0, 0, 0);
                LJUD.ePut(u3.ljhandle, LJUD.IO.PUT_ANALOG_ENABLE_PORT, 0, 15, 16);//first 4 FIO analog b0000000000001111
                LJUD.AddRequest(u3.ljhandle, LJUD.IO.GET_AIN, 0, 0, 0, 0);//Request AIN0
            }
            catch (LabJackUDException)
            {
                MessageBox.Show("Error opening DAQ", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (!pHCal)
            {
                MessageBox.Show("pH Calibration not done!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            pHMeasureStart = DateTime.Now;
            chart2.ChartAreas[0].AxisY.LabelStyle.Format = "#.###";
            chart2.ChartAreas[0].AxisX.LabelStyle.Format = "#.###";
            chart2.ChartAreas[0].AxisY.ScrollBar.Enabled = false;
            chart2.ChartAreas[0].AxisX.ScrollBar.Enabled = false;
            foreach (var series in chart2.Series)
                series.Points.Clear();
            timer2.Enabled = true;
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            timer2.Enabled = false;
            bool requestedExit = false;
            while (!requestedExit)
            {
                try
                {
                    LJUD.GoOne(u3.ljhandle);
                    LJUD.GetFirstResult(u3.ljhandle, ref ioType, ref channel, ref dblValue, ref dummyInt, ref dummyDouble);
                }
                catch (LabJackUDException)
                {
                    MessageBox.Show("Error getting the DAQ results", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                if (ioType == LJUD.IO.GET_AIN)
                {
                    pHMeasureVal = pHCalSlope * dblValue + pHCalIntercept;
                    pHMeasureAvg = pHMeasureAvg + pHMeasureVal;
                    label42.Text = String.Format("{0:0.000000000 V}", dblValue);
                    pHMeasureTicks++;
                    if (pHMeasureTicks >= 30)//average between N samples
                    {
                        label3.Text = String.Format("{0:0.000000}", pHMeasureAvg/30);
                        if (Convert.ToDouble(textBox3.Text) == 0)
                            label37.Text = "0.000000";
                        else
                            label37.Text = String.Format("{0:0.000000}", Convert.ToString(Convert.ToDouble(textBox3.Text) - pHMeasureVal));
                        chart2.Series["Series1"].Points.AddXY((DateTime.Now-pHMeasureStart).TotalSeconds,pHMeasureVal);
                        chart2.ChartAreas[0].AxisY.ScaleView.Position = pHMeasureAvg/30 - 0.25;
                        chart2.ChartAreas[0].AxisY.ScaleView.Size = 0.5;

                        if ((DateTime.Now - pHMeasureStart).TotalSeconds>30)
                            chart2.ChartAreas[0].AxisX.ScaleView.Position = (DateTime.Now-pHMeasureStart).TotalSeconds - 30;
                        pHMeasureAvg = 0;
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
                        MessageBox.Show("Error getting DAQ data", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            timer2.Enabled = true;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            timer2.Enabled = false;
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

                chart1.Series["Series1"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Point;
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
                SampleNo = 0;
                State = 0;
                serialPort1.Write("STP\r\n");
                System.Threading.Thread.Sleep(20);
                serialPort1.Write("CLD INF\r\n");
                System.Threading.Thread.Sleep(20);
                serialPort1.Write("CLD WDR\r\n");
                System.Threading.Thread.Sleep(20);
                serialPort1.Write("DIA " + textBox21.Text + "\r\n");
                System.Threading.Thread.Sleep(20);
                serialPort1.Write("RAT 500 MH\r\n");
                System.Threading.Thread.Sleep(20);
                foreach (var series in chart1.Series)
                    series.Points.Clear();
                //Block controls
                dataGridView1.Enabled = false;
                comboBox22.Enabled = false;
                textBox21.Enabled = false;
                textBox22.Enabled = false;
                checkBox4.Enabled = false;
                button16.Enabled = false;
                button17.Enabled = false;
                button4.Enabled = false;
                button5.Enabled = false;
                button1.Enabled = false;
                button2.Enabled = false;
                button3.Enabled = false;
                button20.Enabled = true;
                timer1.Enabled = true;
            }  
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            switch (State)
            { 
                case 0://sampling started
                        if (SampleNo == dataGridView1.Rows.Count) //end of samples
                        {
                            label1.Text = "Sample No.: " + Convert.ToString(SampleNo);
                            State = 10; //quit
                            timer1.Enabled = true;
                            break;
                        }
                    
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
                    if (Convert.ToDouble(_Reading.Substring(5, 5)) >= Convert.ToDouble(dataGridView1[0, SampleNo-1].Value.ToString()))
                    {
                        State = 2;//input volume form2
                        label33.Text = "Current Volume: " + _Reading.Substring(5, 5) + _Reading.Substring(16, 2).ToLower();
                    }
                    timer1.Enabled = true;
                    break;
                case 2:
                    _SampleUnits = dataGridView1[1, SampleNo-1].Value.ToString();
                    SyrVolForm = new SyrSamInp(_SampleUnits);
                    SyrVolForm.Show();
                    //this.Hide();
                    this.Enabled = false;
                    State = 3;
                    timer1.Enabled = true;
                    break;
                case 3:
                    timer1.Enabled = true;
                    break;
                case 4:
                    this.Enabled = true;
                    dataGridView1[2, SampleNo - 1].Value = SampleVolume;
                    _SampleUnits = Convert.ToString(dataGridView1[1, SampleNo - 1].Value);
                    dataGridView1[3, SampleNo - 1].Value = _SampleUnits;
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
                    SyrR2=(Math.Pow((dataGridView1.RowCount*a-b*c),2))/((dataGridView1.RowCount*d-Math.Pow(b,2))*(dataGridView1.RowCount*f-Math.Pow(c,2)));
                    /*Slope(b) = (NΣXY - (ΣX)(ΣY)) / (NΣX2 - (ΣX)2) Intercept(a) = (ΣY - b(ΣX)) / N */
                    SyrCalSlope = (dataGridView1.RowCount * a - b * c) / (dataGridView1.RowCount*d - Math.Pow(b, 2));
                    SyrCalIntercept=(c-SyrCalSlope*b)/dataGridView1.RowCount;

                    label30.Text = "R =  " + String.Format("{0:0.0000}", SyrR2);
                    if (SyrCalIntercept>=0)
                        label29.Text = "y=" + String.Format("{0:0.0000}", SyrCalSlope) + " x+" + String.Format("{0:0.0000}", SyrCalIntercept);
                    else
                        label29.Text = "y=" + String.Format("{0:0.0000}", SyrCalSlope) + " x" + String.Format("{0:0.0000}", SyrCalIntercept);
                    serialPort1.Close();
                    checkBox1.Checked = true;
                    panel5.Visible = true;
                    dataGridView1.Enabled = true;
                    comboBox22.Enabled = true;
                    textBox21.Enabled = true;
                    textBox22.Enabled = true;
                    checkBox4.Enabled = true;
                    checkBox4.Checked = false;
                    button16.Enabled = false;
                    button17.Enabled = false;
                    button4.Enabled = true;
                    button5.Enabled = true;
                    button1.Enabled = true;
                    if(dataGridView1.RowCount>2)
                        button2.Enabled = true;
                    button3.Enabled = true;
                    button20.Enabled = false;
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
            button16.Enabled = false;
            button17.Enabled = false;
            button4.Enabled = true;
            button5.Enabled = false;
            button1.Enabled = true;
            if (dataGridView1.RowCount > 2)
                button2.Enabled = true;
            button3.Enabled = true;
            button20.Enabled = false;
            serialPort1.Close();
        }

        //Open
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
                    panel5.Visible = true;
                    chart1.Series["Series2"].Points.AddXY(Convert.ToDouble(_LoadCal.Split('#')[3]), Convert.ToDouble(_LoadCal.Split('#')[4]));
                    chart1.Series["Series2"].Points.AddXY(Convert.ToDouble(_LoadCal.Split('#')[5]), Convert.ToDouble(_LoadCal.Split('#')[6]));
                    checkBox1.Checked = true;
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
                    StreamWriter sw = new StreamWriter(saveFileDialog1.FileName);
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

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            
            if (checkBox4.Checked == true)
            {
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
                    serialPort1.Write("VOL 0\r\n");
                    System.Threading.Thread.Sleep(20);
                    serialPort1.Write("CLD INF\r\n");
                    System.Threading.Thread.Sleep(20);
                    serialPort1.Write("CLD WDR\r\n");
                    System.Threading.Thread.Sleep(20);
                    serialPort1.Write("DIA " + textBox21.Text + "\r\n");
                    System.Threading.Thread.Sleep(20);
                    serialPort1.Write("RAT 500 MH\r\n");
                    System.Threading.Thread.Sleep(20);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Port opening", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    checkBox4.Checked = false;
                    return;
                }
                label44.Text = "0 ml";
                label46.Text = "0 ml";
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
            serialPort1.Write("RAT 500 MH\r\n"); //Rate fixed
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
            serialPort1.Write("RAT 500 MH\r\n"); //Rate fixed
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
            label46.Text = String.Format("{0:0.0000}",(Convert.ToDouble(_Reading.Substring(11, 5)))) + " " + _Reading.Substring(16, 2).ToLower();//_Reading.Substring(11, 5).ToLower() + " " + _Reading.Substring(16, 2).ToLower();
            serialPort1.Write("STP\r\n");
            System.Threading.Thread.Sleep(20);
            serialPort1.Close();
        }

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
            label44.Text="0 ml";
            label46.Text = "0 ml";
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

        private bool CheckExpErr()
        {
            serialPort1.Close();
            try
            {
                Convert.ToDouble(textBox2.Text);
            }
            catch (Exception)
            {
                MessageBox.Show("Setpoint is not in a valid format","Error!",MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            if (checkBox1.Checked != true || checkBox2.Checked != true)
            {
                MessageBox.Show("Calibrations not done, please load the calibration files!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            try
            {
                Convert.ToDouble(textBox4.Text);
            }
            catch (Exception)
            {
                MessageBox.Show("Mixing time is not in a valid format", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            if(Convert.ToDouble(textBox4.Text) < 0 || Convert.ToDouble(textBox4.Text) > 60)
            {
                MessageBox.Show("Please choose a mixing time between 0 and 60 seconds!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            if (comboBox1.SelectedIndex==-1)
            {
                MessageBox.Show("Please choose a sampling rate!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            return true;
        }
        private void button14_Click(object sender, EventArgs e)
        {
            if (Paused)
            {
                button13.Enabled = true;
                button15.Enabled = true;
                button14.Enabled = false;
                Paused = false;
                timer4.Enabled = true;
                textBox2.Enabled = false;                
                return;
            }
            if (!CheckExpErr())
                return;

            try
            {
                serialPort1.PortName = "COM" + Convert.ToString(comboBox22.SelectedIndex);
                serialPort1.Open();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Port opening", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
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
            AccumVolInf = 0;
            try
            {
                switch (comboBox1.SelectedIndex)
                {
                    case 0:
                        SubSampling = 500;
                        break;
                    case 1:
                        SubSampling = 1000;
                        break;
                    case 2:
                        SubSampling = 2000;
                        break;
                    case 3:
                        SubSampling = 3000;
                        break;
                    case 4:
                        SubSampling = 4000;
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Sumbsampling rate", MessageBoxButtons.OK, MessageBoxIcon.Error);
                serialPort1.Close();
                return;
            }
            //Configure DAQ
            try
            {
                if (u3 == null)
                    u3 = new U3(LJUD.CONNECTION.USB, "1", true); // Connection through USB
                LJUD.ePut(u3.ljhandle, LJUD.IO.PIN_CONFIGURATION_RESET, 0, 0, 0);
                LJUD.ePut(u3.ljhandle, LJUD.IO.PUT_ANALOG_ENABLE_PORT, 0, 15, 16);//first 4 FIO analog b0000000000001111
                LJUD.AddRequest(u3.ljhandle, LJUD.IO.GET_AIN, 0, 0, 0, 0);//Request AIN0
            }
            catch (LabJackUDException)
            {
                MessageBox.Show("Error opening DAQ", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                serialPort1.Close();
                return;
            }
            label23.Text = String.Format("{0:0.0000}", pHCalSlope);
            label35.Text = String.Format("{0:0.0000}", SyrCalSlope);
            chart4.ChartAreas[0].AxisY.LabelStyle.Format = "#.###";
            chart4.ChartAreas[0].AxisX.LabelStyle.Format = "#";
            ExpStart = DateTime.Now;
            button13.Enabled = true;
            button15.Enabled = true;
            button14.Enabled = false;
            foreach (var series in chart4.Series)
                series.Points.Clear();
            ExpState = 0;
            textBox2.Enabled = false;
            textBox4.Enabled = false;
            comboBox1.Enabled = false;
           
            // Create a timer with a two second interval.
            aTimer = new System.Timers.Timer(SubSampling);
            // Hook up the Elapsed event for the timer. 
            aTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent);
            aTimer.Enabled = true;
            timer4.Enabled = true;
        }

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
                catch (LabJackUDException)
                {
                    MessageBox.Show("Error getting the DAQ results", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                if (ioType == LJUD.IO.GET_AIN)
                {
                    switch (ExpState)
                    { 
                        case 0:
                            ExpTicks++;
                            ExpAccVal = ExpAccVal + (dblValue * pHCalSlope + pHCalIntercept);
                            if (ExpTicks >= 30)
                                ExpState = 1;
                            break;
                        case 1:
                            ExpAvgVal = ExpAccVal / 30;
                            label13.Text = String.Format("{0:0.000000000 V}", dblValue);
                            label16.Text = String.Format("{0:0.0000000}", ExpAvgVal);
                            //******************************************************
                            Deviation = Convert.ToDouble(textBox2.Text) - ExpAvgVal;
                            //******************************************************                        
                            label17.Text = String.Format("{0:0.0000000}", Deviation);
                            label21.Text = String.Format("{0:00}", (DateTime.Now - ExpStart).Hours) + ":" + String.Format("{0:00}", (DateTime.Now - ExpStart).Minutes) + ":" + String.Format("{0:00}", (DateTime.Now - ExpStart).Seconds);
                            //******************************************************
                            label19.Text = String.Format("{0:00000.00}", AccumVolInf) + " ul";

                            ExpTicks = 0;
                            ExpAccVal = 0;
                            //if it is paused, do not infuse... just read ph
                            if (Paused)
                            {
                                ExpState = 0;
                                break;
                            }                            
                            if (Deviation > 0.001 && InfStarted == false)
                            {
                                ExpState = 2;
                                serialPort1.Write("CLD INF\r\n");
                                System.Threading.Thread.Sleep(20);
                                break;
                            }
                            if (InfStarted && DateTime.Now.AddSeconds(-10) > ExpWaitTime)
                            {
                                ExpState = 3;
                                break;
                            }
                            ExpState = 0;
                            break;
                        case 2:
                            InfStarted = true;
                            ExpWaitTime=DateTime.Now;
                            VoltoInf = ((Deviation * (50 / 0.03)) - SyrCalIntercept) / SyrCalSlope;
                            if (VoltoInf < 10)
                                VoltoInf = 10;
                            if (VoltoInf > 50)
                                VoltoInf = 50;
                            serialPort1.Write("VOL " + String.Format("{0:000.0}", VoltoInf) + "\r\n");
                            AccumVolInf = AccumVolInf + VoltoInf;
                            System.Threading.Thread.Sleep(20);
                            serialPort1.Write("RUN\r\n");
                            System.Threading.Thread.Sleep(20);
                            ExpState=0;
                            break;
                        case 3:
                            InfStarted = false;
                            serialPort1.Write("STP\r\n");
                            System.Threading.Thread.Sleep(20);
                            ExpState=0;
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
                        MessageBox.Show("Error getting DAQ data", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            timer4.Enabled = true;
        }

        private void OnTimedEvent(Object source, ElapsedEventArgs e)
        {
            TimeDif = (DateTime.Now - ExpStart).TotalSeconds;
            chart4.Series["Series1"].Points.AddXY(TimeDif, ExpAvgVal);
            chart4.Series["Series2"].Points.AddXY(TimeDif, AccumVolInf);
            chart4.Series["Series3"].Points.AddXY(TimeDif, Deviation);
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox2.SelectedIndex)
            {
                case -1:
                    chart4.Series["Series1"].Enabled = true;
                    chart4.Series["Series2"].Enabled = false;
                    chart4.Series["Series3"].Enabled = false;
                    break;
                case 0:
                    chart4.Series["Series1"].Enabled = true;
                    chart4.Series["Series2"].Enabled = false;
                    chart4.Series["Series3"].Enabled = false;
                    break;
                case 1:
                    chart4.Series["Series1"].Enabled = false;
                    chart4.Series["Series2"].Enabled = true;
                    chart4.Series["Series3"].Enabled = false;
                    break;
                case 2:
                    chart4.Series["Series1"].Enabled = false;
                    chart4.Series["Series2"].Enabled = false;
                    chart4.Series["Series3"].Enabled = true;
                    break;
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            Paused = true;
            button13.Enabled = false;
            button15.Enabled = false;
            button14.Enabled = true;
            textBox2.Enabled = false;
        }

        private void button15_Click(object sender, EventArgs e)
        {
            serialPort1.Close();
            timer4.Enabled = false;
            aTimer.Enabled = false;
            aTimer.Enabled = false;
            button13.Enabled = false;
            button15.Enabled = false;
            button14.Enabled = true;
            textBox2.Enabled = true;
            textBox4.Enabled = true;
            comboBox1.Enabled = true;
            //Export Data
            string _FirstLine = "PH,X,Y,,VOLUME,X,Y,,DEVIATION,X,Y\n";
            string[] Content = new string[chart4.Series["Series1"].Points.Count];
            string Path = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)+ @"\Calcification Experiments";
            try
            {
                if (!System.IO.Directory.Exists(Path))
                    System.IO.Directory.CreateDirectory(Path);
                StreamWriter sw = new StreamWriter(Path+@"\"+DateTime.Now.ToString("yyyy-MM-dd HHmmss")+".csv");
                sw.Write(_FirstLine);
                for (int i = 0; i < chart4.Series["Series1"].Points.Count; i++)
                {
                    sw.Write("," + chart4.Series["Series1"].Points[i].XValue + "," + chart4.Series["Series1"].Points[i].YValues[0] + ",,," + chart4.Series["Series2"].Points[i].XValue + "," + chart4.Series["Series2"].Points[i].YValues[0] + ",,," + chart4.Series["Series3"].Points[i].XValue + "," + chart4.Series["Series3"].Points[i].YValues[0] + "\n");
                }
                sw.Close();
                Paused = false;
                ExpTicks = 0;
                ExpAvgVal = 0;
                ExpAccVal = 0;
                Deviation=0;
                VoltoInf=0;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error Writing File", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } 
            //End Export Data
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
        }


        //*********************************************************************************
        //***********************************END*******************************************
    }

}

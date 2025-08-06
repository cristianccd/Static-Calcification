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

namespace TECAS_Static_Calcification
{
    public partial class TECAS : Form
    {
        //Variables
        //***************************
        //pH Calibration
        DateTime pHstartTime;
        double pHCalSlope = 0, pHCalIntercept = 0, pHR2 = 0, SumpHCal = 0, pHAccumTime = 0, pHTicks = 0, pHAvgVal = 0;
        int pHSampleNo = 0, pHCalState = 0;
        bool pHCal = false;

        //pH measurement
        double pHMeasureAvg = 0, pHMeasureTicks = 0, pHMeasureVal = 0;
        DateTime pHMeasureStart;
        DialogResult pHQuestSample;

        //Syringe Calibration
        double SyrR2=0, SyrCalIntercept=0, SyrCalSlope=0;
        bool SyrCal = false;
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
        double ExpTicks = 0, ExpAvgVal = 0, ExpAccVal = 0, ExpAvgSyr = 0, ExpCurrVol=0, Deviation=0, VoltoInf=0, AccVol=0;
        DateTime ExpStart;
        bool Paused = false;


        //***************************

        public TECAS()
        {
            InitializeComponent();
            //pH calibration initialization
            dataGridView2.Rows.Add();
            dataGridView2.Rows.Add();
            

            //add 2 rows for each table (minimum for calibration)
            dataGridView1.Rows.Add();
            dataGridView1.Rows.Add();
            //disable units in grid1
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
                    MessageBox.Show("Time must be possitive!");
                    return false;
                }
            }
            catch (Exception h)
            {
                MessageBox.Show("Time format not valid");
                return false;
            }
            if (dataGridView2.RowCount < 2)
            {
                MessageBox.Show("Not eneough Samples");
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
                        MessageBox.Show("pH sample No."+Convert.ToString(i+1)+" not correct!");
                        return false;
                    }
                }
                catch (Exception h)
                {
                    MessageBox.Show("pH format not valid");
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
                catch (LabJackUDException h)
                {
                    MessageBox.Show("Error opening DAQ");
                    return;
                }
                foreach (var series in chart3.Series)
                    series.Points.Clear();
                timer3.Enabled = true;
            }
            else
                return;
        }

        private void timer3_Tick(object sender, EventArgs e)
        {
            timer3.Enabled = false; //first disable the timer to perform actions
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
                    MessageBox.Show("Error getting the DAQ results");
                }
                if(ioType==LJUD.IO.GET_AIN)
                    label10.Text = String.Format("{0:0.000000}", dblValue);
                try 
                {
                    LJUD.GetNextResult(u3.ljhandle, ref ioType, ref channel, ref dblValue, ref dummyInt, ref dummyDouble); 
                }
                catch (LabJackUDException h)
                {
                    if (h.LJUDError == U3.LJUDERROR.NO_MORE_DATA_AVAILABLE)
                        requestedExit = true;//no more data to read
                    else
                        MessageBox.Show("Error getting DAQ data");
                }
            }
            //Start of calibration
            switch (pHCalState)
            {
                case 0://calibration started
                    pHQuestSample = MessageBox.Show("When you are ready to measure the sample " + Convert.ToString(pHSampleNo + 1) + " press OK", MessageBoxButtons.OK.ToString());
                    //if the result is OK, change state and start counting time
                    if (pHQuestSample == DialogResult.OK)
                    {
                        pHCalState = 1;
                        pHAccumTime = 0;
                        pHstartTime = DateTime.Now;
                    }
                    //otherwise just wait until is OK
                    else
                        return;
                    timer3.Enabled = true;
                    break;
                case 1://Start accumulating time and check if the target is reached
                    pHAccumTime = pHAccumTime + Convert.ToDouble(timer3.Interval) / 1000 / 60;
                    label12.Text = (DateTime.Now - pHstartTime).Minutes + ":" + (DateTime.Now - pHstartTime).Seconds + " mins";
                    if (Convert.ToDouble((DateTime.Now - pHstartTime).TotalMinutes) >= Convert.ToDouble(textBox1.Text))//pHAccumTime >= Convert.ToDouble(textBox1.Text))
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
                    MessageBox.Show("Finished, please check the results...","Finished");
                    //auxiliar variables to obtain results
                    double a = 0, b = 0, c = 0, d = 0, f = 0;
                    for (int i = 0; i < dataGridView2.RowCount; i++)
                    {
                        //sumxy
                        a= a + Convert.ToDouble(dataGridView2[0, i].Value) * Convert.ToDouble(dataGridView2[1, i].Value);
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
                    /*R2 = ((Nxysum - xsumysum)^2)/(Nx^2sum - xsum*xsum)*(Ny^2sum - ysum*ysum)*/
                    pHR2=(Math.Pow((dataGridView2.RowCount*a-b*c),2))/((dataGridView2.RowCount*d-Math.Pow(b,2))*(dataGridView2.RowCount*f-Math.Pow(c,2)));
                    /*Slope(b) = (NΣXY - (ΣX)(ΣY)) / (NΣX2 - (ΣX)2) Intercept(a) = (ΣY - b(ΣX)) / N */
                    pHCalSlope = (dataGridView2.RowCount * a - b * c) / (dataGridView2.RowCount*d - Math.Pow(b, 2));
                    pHCalIntercept=(c-pHCalSlope*b)/dataGridView2.RowCount;
                    if(pHCalIntercept>=0)
                        label6.Text = "y=" + String.Format("{0:0.0000}", pHCalSlope) + " x+" + String.Format("{0:0.0000}", pHCalIntercept);
                    else
                        label6.Text = "y=" + String.Format("{0:0.0000}", pHCalSlope) + " x" + String.Format("{0:0.0000}", pHCalIntercept);
                    label5.Text = "R =  " + String.Format("{0:0.0000}", pHR2);
                    chart3.Series["Series1"].Points.AddXY(Convert.ToDouble(dataGridView2[0,0].Value),pHCalSlope*Convert.ToDouble(dataGridView2[0,0].Value)+pHCalIntercept);
                    chart3.Series["Series1"].Points.AddXY(Convert.ToDouble(dataGridView2[0, dataGridView2.RowCount-1].Value), pHCalSlope * Convert.ToDouble(dataGridView2[0, dataGridView2.RowCount - 1].Value) + pHCalIntercept);
                    panel1.Visible = true;
                    button12.Enabled = true;
                    pHTicks = 0;
                    SumpHCal = 0;
                    pHSampleNo = 0;
                    pHCalState = 0;
                    pHCal = true;
                    button6.Enabled = true;
                    checkBox3.Checked = true;
                    checkBox2.Checked = true;
                    return;
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
                catch (Exception h)
                {
                    MessageBox.Show("Invalid Set Point Format!");
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
            catch (LabJackUDException h)
            {
                MessageBox.Show("Error opening DAQ");
                return;
            }
            if (!pHCal)
            {
                MessageBox.Show("pH Calibration not done!");
                return;
            }
            pHMeasureStart = DateTime.Now;
            chart2.ChartAreas[0].AxisY.LabelStyle.Format = "#.###";
            chart2.ChartAreas[0].AxisX.LabelStyle.Format = "#.###";
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
                catch (LabJackUDException h)
                {
                    MessageBox.Show("Error getting the DAQ results");
                }
                if (ioType == LJUD.IO.GET_AIN)
                {
                    pHMeasureVal = pHCalSlope * dblValue + pHCalIntercept;
                    pHMeasureAvg = pHMeasureAvg + pHMeasureVal;
                    label42.Text = String.Format("{0:0.000000000 mV}", dblValue);
                    pHMeasureTicks++;
                    if (pHMeasureTicks >= 50)//average between N samples
                    {
                        
                        label3.Text = String.Format("{0:0.000000}", pHMeasureVal);
                        if (Convert.ToDouble(textBox3.Text) == 0)
                            label37.Text = "0.000000";
                        else
                            label37.Text = Convert.ToString(Convert.ToDouble(textBox3.Text) - pHMeasureVal);
                        chart2.Series["Series1"].Points.AddXY((DateTime.Now-pHMeasureStart).TotalSeconds,pHMeasureVal);
                        if ((DateTime.Now - pHMeasureStart).TotalSeconds>30)
                            chart2.ChartAreas[0].AxisX.ScaleView.Position = (DateTime.Now-pHMeasureStart).TotalSeconds - 30;
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
                        MessageBox.Show("Error getting DAQ data");
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
                MessageBox.Show("COM port not selected!");
                return false;
            }
            try//if diameter and capacity are numbers
            {
                Convert.ToDouble(textBox21.Text);
                Convert.ToDouble(textBox22.Text);
            }
            catch (Exception) //Diameter or capacity not double
            {
                MessageBox.Show("Wrong input parameters!");
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
                    MessageBox.Show("Please check the sample's volumes");
                    return false;
                }
                //now if indexes are right
                if (dataGridView1[1, i].Value.ToString() != "ml" && dataGridView1[1, i].Value.ToString() != "ul")
                {
                    MessageBox.Show("Units are not properly selected");
                    return false;
                }
            }
            if (dataGridView1.RowCount < 2)
            {
                MessageBox.Show("Not enough samples for calibration!");
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
                    
                        SyrQuestSample=MessageBox.Show("When you are ready to measure the sample " + Convert.ToString(SampleNo+1) + " press OK", MessageBoxButtons.OK.ToString());
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
                        else
                            return;
                        timer1.Enabled = true;    
                        break;
                    
                case 1:
                    serialPort1.ReadExisting();
                    serialPort1.Write("DIS\r\n");
                    System.Threading.Thread.Sleep(20);
                    _Reading=serialPort1.ReadExisting();
                    label33.Text = "Current Volume: " + _Reading.Substring(5, 5) + _Reading.Substring(16, 2).ToLower();
                    if (Convert.ToDouble(_Reading.Substring(5, 5)) == Convert.ToDouble(dataGridView1[0, SampleNo-1].Value.ToString()))
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
                    this.Hide();
                    State = 3;
                    timer1.Enabled = true;
                    break;
                case 3:
                    timer1.Enabled = true;
                    break;
                case 4:
                    this.Show();
                    dataGridView1[2, SampleNo - 1].Value = SampleVolume;
                    _SampleUnits = Convert.ToString(dataGridView1[1, SampleNo - 1].Value);
                    dataGridView1[3, SampleNo - 1].Value = _SampleUnits;
                    State = 0;
                    timer1.Enabled = true;
                    break;
                case 10:
                    MessageBox.Show("Finished, please check the results...",MessageBoxButtons.OK.ToString());
                    double a = 0, b = 0, c = 0, d = 0, f = 0;
                    for (int i = 0; i < dataGridView1.RowCount; i++)
                    {
                        //sumxy
                        a= a + Convert.ToDouble(dataGridView1[0, i].Value) * Convert.ToDouble(dataGridView1[2, i].Value);
                        //sumx
                        b= b + Convert.ToDouble(dataGridView1[0, i].Value);
                        //sumy
                        c= c +Convert.ToDouble(dataGridView1[2, i].Value);
                        //sumx2
                        d= d + Math.Pow(Convert.ToDouble(dataGridView1[0, i].Value), 2);
                        //sumy2
                        f= f + Math.Pow(Convert.ToDouble(dataGridView1[2, i].Value),2);
                        chart3.Series["Series2"].Points.AddXY(Convert.ToDouble(dataGridView1[0,i].Value),dataGridView1[2, i].Value);
                    }
                    /*R2 = ((Nxysum - xsumysum)^2)/(Nx^2sum - xsum*xsum)*(Ny^2sum - ysum*ysum)*/
                    SyrR2=(Math.Pow((dataGridView1.RowCount*a-b*c),2))/((dataGridView1.RowCount*d-Math.Pow(b,2))*(dataGridView1.RowCount*f-Math.Pow(c,2)));
                    /*Slope(b) = (NΣXY - (ΣX)(ΣY)) / (NΣX2 - (ΣX)2) Intercept(a) = (ΣY - b(ΣX)) / N */
                    SyrCalSlope = (dataGridView1.RowCount * a - b * c) / (dataGridView1.RowCount*d - Math.Pow(b, 2));
                    SyrCalIntercept=(c-SyrCalSlope*b)/dataGridView1.RowCount;
                    chart1.Series["Series2"].Points.AddXY(Convert.ToDouble(dataGridView1[0, 0].Value), SyrCalSlope * Convert.ToDouble(dataGridView1[0, 0].Value) + SyrCalIntercept);
                    chart1.Series["Series2"].Points.AddXY(Convert.ToDouble(dataGridView1[0, dataGridView1.RowCount - 1].Value), SyrCalSlope * Convert.ToDouble(dataGridView1[0, dataGridView1.RowCount - 1].Value) + SyrCalIntercept);
                    label30.Text = "R =  " + String.Format("{0:0.0000}", SyrR2);
                    if (SyrCalIntercept>=0)
                        label29.Text = "y=" + String.Format("{0:0.0000}", SyrCalSlope) + " x+" + String.Format("{0:0.0000}", SyrCalIntercept);
                    else
                        label29.Text = "y=" + String.Format("{0:0.0000}", SyrCalSlope) + " x" + String.Format("{0:0.0000}", SyrCalIntercept);
                    serialPort1.Close();
                    SyrCal = true;
                    checkBox1.Checked = true;
                    panel5.Visible = true;
                    button5.Enabled = true;
                    return;
            }
        }

        private void TECAS_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (serialPort1.IsOpen)
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
                    SyrCal = true;
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
                button16.Enabled = true;
                button17.Enabled = true;
                try
                {
                    serialPort1.PortName = "COM" + Convert.ToString(comboBox22.SelectedIndex);
                    serialPort1.Open();
                    serialPort1.Write("VOL 0\r\n");
                    System.Threading.Thread.Sleep(20);
                    serialPort1.Write("CLD INF\r\n");
                    System.Threading.Thread.Sleep(20);
                    serialPort1.Write("CLD WDR\r\n");
                    System.Threading.Thread.Sleep(20);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Port opening", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                
            }
            else
            {
                button16.Enabled = false;
                button17.Enabled = false;
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
            serialPort1.Write("DIA " + textBox21.Text + "\r\n");
            System.Threading.Thread.Sleep(20);
            serialPort1.Write("DIR INF\r\n");
            System.Threading.Thread.Sleep(20);
            serialPort1.Write("RAT 100 MH\r\n"); //Rate fixed
            System.Threading.Thread.Sleep(20);
            System.Threading.Thread.Sleep(20);
            serialPort1.Write("RUN\r\n");
            System.Threading.Thread.Sleep(20);
        }

        private void button17_MouseUp(object sender, MouseEventArgs e)
        {
            _Reading = serialPort1.ReadExisting();
            serialPort1.Write("DIS\r\n");
            System.Threading.Thread.Sleep(20);
            _Reading = serialPort1.ReadExisting();
            label44.Text = _Reading.Substring(5, 5).ToLower() +" "+_Reading.Substring(16, 2).ToLower();
            serialPort1.Write("STP\r\n");
            System.Threading.Thread.Sleep(20);
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
            serialPort1.Write("DIA " + textBox21.Text + "\r\n");
            System.Threading.Thread.Sleep(20);
            serialPort1.Write("DIR WDR\r\n");
            System.Threading.Thread.Sleep(20);
            serialPort1.Write("RAT 50 MH\r\n"); //Rate fixed
            System.Threading.Thread.Sleep(20);
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
            label46.Text = _Reading.Substring(11, 5).ToLower() + " " + _Reading.Substring(16, 2).ToLower();
            serialPort1.Write("STP\r\n");
            System.Threading.Thread.Sleep(20);
            serialPort1.Write("DIR INF\r\n");
            System.Threading.Thread.Sleep(20);
            serialPort1.Close();
        }

        //*********************************************************************************
        //***********************************END*******************************************

        //*********************************************************************************
        //******************************STATIC EXPERIMENT**********************************

        private bool CheckExpErr()
        {
            try
            {
                Convert.ToDouble(textBox2.Text);
            }
            catch (Exception h)
            {
                MessageBox.Show("Setpoint is not in a valid format");
                return false;
            }
            if (checkBox1.Checked != true || checkBox2.Checked != true)
            {
                MessageBox.Show("Calibrations not done, please load the calibration files!");
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
            serialPort1.Write("DIA " + textBox21.Text + "\r\n");
            System.Threading.Thread.Sleep(20);
            serialPort1.Write("RAT 500 MH\r\n"); //Rate fixed
            System.Threading.Thread.Sleep(20);
            serialPort1.Write("VOL UL\r\n"); //Rate fixed
            System.Threading.Thread.Sleep(20);

            //Configure DAQ
            try
            {
                if (u3 == null)
                    u3 = new U3(LJUD.CONNECTION.USB, "1", true); // Connection through USB
                LJUD.ePut(u3.ljhandle, LJUD.IO.PIN_CONFIGURATION_RESET, 0, 0, 0);
                LJUD.ePut(u3.ljhandle, LJUD.IO.PUT_ANALOG_ENABLE_PORT, 0, 15, 16);//first 4 FIO analog b0000000000001111
                LJUD.AddRequest(u3.ljhandle, LJUD.IO.GET_AIN, 0, 0, 0, 0);//Request AIN0
            }
            catch (LabJackUDException h)
            {
                MessageBox.Show("Error opening DAQ");
                return;
            }
            label23.Text = String.Format("{0:0.000000}", pHCalSlope);
            label35.Text = String.Format("{0:0.000000}", SyrCalSlope);
            chart4.ChartAreas[0].AxisY.LabelStyle.Format = "#.###";
            chart4.ChartAreas[0].AxisX.LabelStyle.Format = "#.###";
            ExpStart = DateTime.Now;
            button13.Enabled = true;
            button15.Enabled = true;
            button14.Enabled = false;
            foreach (var series in chart4.Series)
                series.Points.Clear();
            timer4.Enabled = true;
            textBox2.Enabled = false;
        }

        private void timer4_Tick(object sender, EventArgs e)
        {
            timer4.Enabled = false;
            if (Paused)
            {
                timer4.Enabled = false;
                return;
            }

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
                    MessageBox.Show("Error getting the DAQ results");
                }
                if (ioType == LJUD.IO.GET_AIN)
                {
                    ExpTicks++;
                    ExpAccVal = ExpAccVal + (dblValue * pHCalSlope + pHCalIntercept);
                    label13.Text = String.Format("{0:0.000000000 mV}", dblValue);
                    if (ExpTicks >= 10)//average between N samples
                    {
                        ExpAvgVal = ExpAccVal / 50;
                        label16.Text = String.Format("{0:0.0000000}", ExpAvgVal);
                        chart4.Series["Series1"].Points.AddXY((DateTime.Now - ExpStart).TotalSeconds, ExpAvgVal);
                        serialPort1.ReadExisting();
                        serialPort1.Write("DIS\r\n");
                        System.Threading.Thread.Sleep(20);
                        _Reading = serialPort1.ReadExisting();
                        ExpCurrVol = Convert.ToDouble(_Reading.Substring(5, 5));
                        ExpCurrVol = ExpCurrVol*SyrCalSlope+SyrCalIntercept;
                        Deviation = Convert.ToDouble(textBox2.Text) - ExpAvgVal;
                        chart4.Series["Series2"].Points.AddXY((DateTime.Now - ExpStart).TotalSeconds, ExpCurrVol);
                        label17.Text = String.Format("{0:0.0000000}", Deviation);
                        label21.Text = String.Format("{0:0}", (DateTime.Now - ExpStart).TotalDays) + " days " + String.Format("{0:00}", (DateTime.Now - ExpStart).Hours) + ":" + String.Format("{0:00}", (DateTime.Now - ExpStart).Minutes) + ":" + String.Format("{0:00}", (DateTime.Now - ExpStart).Seconds);
                        //Start Control Loop
                        if (Deviation > 0.0005)
                        {
                            //proportional control VoltoInf=Kp*e(t)+p0 -- p0: set point; Kp: Gain; e(t)=error
                            VoltoInf = Deviation/0.005; //1ul per 0.005
                            serialPort1.Write("VOL " + Convert.ToString(VoltoInf*SyrCalSlope+SyrCalIntercept) + "\r\n");
                            System.Threading.Thread.Sleep(20);
                            //serialPort1.Write("VOL UL\r\n");
                            System.Threading.Thread.Sleep(20);
                            serialPort1.Write("RUN\r\n");
                            AccVol = AccVol + ExpCurrVol;
                        }
                        else
                            serialPort1.Write("STP\r\n");
                        System.Threading.Thread.Sleep(20);
                        //End Control Loop
                        label19.Text = String.Format("{0:0.000}", AccVol)+_Reading.Substring(16, 2).ToLower();
                        chart4.Series["Series3"].Points.AddXY((DateTime.Now - ExpStart).TotalSeconds, Deviation);
                        ExpTicks = 0;
                        ExpAccVal = 0;
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
                        MessageBox.Show("Error getting DAQ data");
                }
            }
            timer4.Enabled = true;
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
            button13.Enabled = false;
            button15.Enabled = false;
            button14.Enabled = true;
            textBox2.Enabled = true;
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
                    sw.Write(","+chart4.Series["Series1"].Points[i].XValue + "," + chart4.Series["Series1"].Points[i].YValues[0] + ",,," + chart4.Series["Series2"].Points[i].XValue + "," + chart4.Series["Series2"].Points[i].YValues[0] + ",,," + chart4.Series["Series3"].Points[i].XValue + "," + chart4.Series["Series3"].Points[i].YValues[0] + "\n");
                sw.Close();
                Paused = false;
                ExpTicks = 0;
                ExpAvgVal = 0;
                ExpAccVal = 0;
                ExpAvgSyr = 0;
                ExpCurrVol=0;
                Deviation=0;
                VoltoInf=0;
                AccVol=0;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error Writing File", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } 
            //End Export Data
        }

        //*********************************************************************************
        //***********************************END*******************************************
    }

}

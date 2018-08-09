using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.IO;
using ExcelDataReader;
using System.Data.OleDb;
using GemBox.Spreadsheet;


namespace TEST
{
    public partial class Form1 : Form
    {

        String Send_data;
        String Receice_data;
        string[] time;
        bool isClose = false;
        
        public Form1()
        {
            InitializeComponent();
            
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            string[] ports = SerialPort.GetPortNames();
            cBox_Port.Items.AddRange(ports);
            disable_control();
            timer1.Interval = 1000;
            timer1.Start();
            
        }

        private void btn_Connect_Click(object sender, EventArgs e)
        {
            try
            {
                serialPort1.PortName = cBox_Port.Text;
                serialPort1.BaudRate = Convert.ToInt32(cBox_Baudrate.Text);
                serialPort1.DataBits = Convert.ToInt32(cBox_Data.Text);
                serialPort1.StopBits = (StopBits)Enum.Parse(typeof(StopBits), cBox_Stop.Text);
                serialPort1.Parity = (Parity)Enum.Parse(typeof(Parity), cBox_Parity.Text);

                serialPort1.Open();
                progressBar1.Value = 100;
                btn_Connect.Visible = false;
                btn_Connect.Enabled = false;
                btn_Disconnect.Enabled = true;
                btn_Disconnect.Visible = true;
                Enable_control();
                //timer1.Start();

            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Enable_control()
        {
            txt_Send.Enabled = true;
            btn_Send.Enabled = true;
            button1.Enabled = true;
            txt_Read.Enabled = true;
            txt_Read.Enabled = true;
            TEST.Enabled = true;
        }

        private void disable_control()
        {
            txt_Send.Enabled = false;
            btn_Send.Enabled = false;
            button1.Enabled = false;
            txt_Read.Enabled = false;
            txt_Read.Enabled = false;
            TEST.Enabled = false;
        }

        private void btn_Disconnect_Click(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen && isClose==true)
            {
                serialPort1.Close();
                progressBar1.Value = 0;
                btn_Connect.Enabled = true;
                btn_Connect.Visible = true;
                btn_Disconnect.Enabled = false;
                btn_Disconnect.Visible = false;
                disable_control();
            }
            if (serialPort1.IsOpen && isClose == false)
            {
                MessageBox.Show("VUI LONG THỬ LẠI", "LỖI NGẮT KẾT NỐI", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btn_Send_Click(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                Send_data = txt_Send.Text;
                serialPort1.Write(Send_data + "\r\n");  
            }
        }
       
        private void serialPort1_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            isClose = false;
            Thread.Sleep(500);
            Receice_data = serialPort1.ReadExisting();
            this.Invoke(new EventHandler(Showdata));
            isClose = true;
        }
     
        private void Showdata(object sender, EventArgs e)
        {
            if(TEST.SelectedTab==tabPage1)
            {
                txt_Read.Text += Receice_data;
            }
            //Led_display();
            RTC_display();
        }

        private void RTC_display()
        {
            string[] rtc_data = Receice_data.Split('*');
            for (UInt32 i = 0; i < rtc_data.Length; i++)
            {
                time = rtc_data[i].Split(',');
                if (time.Length == 21 && time[0] == "T")
                {
                    label15.Text = time[1] + "/" + time[2] + "/" + time[3] + "  " + time[4] + ":" + time[5] + ":" + time[6];
                    label24.Text = time[7] + "/" + time[8] + "/" + time[9] + "  " + time[10] + ":" + time[11] + ":" + time[12];
                    if (time[13] == "1") led1.On = true;
                    if (time[13] == "0") led1.On = false;
                    if (time[14] == "1") led2.On = true;
                    if (time[14] == "0") led2.On = false;
                    if (time[15] == "1") led3.On = true;
                    if (time[15] == "0") led3.On = false;
                    if (time[16] == "1") led4.On = true;
                    if (time[16] == "0") led4.On = false;
                    label37.Text = time[17];
                    label9.Text = time[18];
                    label38.Text = time[19];
                }
            }
        }

        private void btn_Read_Click(object sender, EventArgs e)
        {
            if (txt_Read.Text != "")
            {
                txt_Read.Text = "";
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (txt_Send.Text != "")
            {
                txt_Send.Text = "";
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                serialPort1.Write("<\r\n");
            }
            else
            {
                serialPort1.Write(">\r\n");
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                serialPort1.Write("?\r\n");
            }
            else
            {
                serialPort1.Write("/\r\n");
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked)
            {
                serialPort1.Write(";\r\n");
            }
            else
            {
                serialPort1.Write(":\r\n");
            }
        }

      

        private void btn_Time_Click(object sender, EventArgs e)
        {
            
            if (cBox_h1.Text != "" && cBox_m1.Text != "" && cBox_h2.Text != "" && cBox_m2.Text != "")
            {
                serialPort1.Write("S1," + cBox_h1.Text + "," + cBox_m1.Text + "," + cBox_h2.Text + "," + cBox_m2.Text + "\r\n");
            }
            else
            {
                MessageBox.Show("VUI LONG CAI DAT LAI THOI GIAN", "LỖI CÀI ĐẶT THỜI GIAN",  MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {   
            serialPort1.Write("~\r\n");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            serialPort1.Write("!\r\n");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            serialPort1.Write("@\r\n");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (comboBox4.Text != "" && comboBox3.Text != "" && comboBox2.Text != "" && comboBox1.Text != "")
            {
                serialPort1.Write("S2," + comboBox4.Text + "," + comboBox3.Text + "," + comboBox2.Text + "," + comboBox1.Text + "\r\n");
            }
            else
            {
                MessageBox.Show("VUI LONG CAI DAT LAI THOI GIAN", "LỖI CÀI ĐẶT THỜI GIAN", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            //RTC_display();
            check_COM();
        }

        int slBanDau = 0;
        string comBanDau = string.Empty;

        private void check_COM()
        {
            string selected = RefreshComPortList(cBox_Port.Items.Cast<string>(), cBox_Port.SelectedItem as string, serialPort1.IsOpen);
            if (!String.IsNullOrEmpty(selected))
            {
                if (cBox_Port.Items.Count < slBanDau && selected != comBanDau)
                {
                    btn_Connect.PerformClick();
                }
                cBox_Port.Items.Clear();
                cBox_Port.Items.AddRange(OrderedPortNames());
                cBox_Port.SelectedItem = selected;

                slBanDau = cBox_Port.Items.Count;
                comBanDau = selected;
            }
        }

        private string[] OrderedPortNames()
        {
            // Just a placeholder for a successful parsing of a string to an integer
            int num;

            // Order the serial port names in numberic order (if possible)
            return SerialPort.GetPortNames().OrderBy(a => a.Length > 3 && int.TryParse(a.Substring(3), out num) ? num : 0).ToArray();
        }

        private string RefreshComPortList(IEnumerable<string> PreviousPortNames, string CurrentSelection, bool PortOpen)
        {
            // Create a new return report to populate
            string selected = null;

            // Retrieve the list of ports currently mounted by the operating system (sorted by name)
            string[] ports = SerialPort.GetPortNames();

            // First determain if there was a change (any additions or removals)
            bool updated = PreviousPortNames.Except(ports).Count() > 0 || ports.Except(PreviousPortNames).Count() > 0;

            // If there was a change, then select an appropriate default port
            if (updated)
            {
                // Use the correctly ordered set of port names
                ports = OrderedPortNames();

                // Find newest port if one or more were added
                string newest = SerialPort.GetPortNames().Except(PreviousPortNames).OrderBy(a => a).LastOrDefault();

                // If the port was already open... (see logic notes and reasoning in Notes.txt)
                if (PortOpen)
                {
                    if (ports.Contains(CurrentSelection)) selected = CurrentSelection;
                    else if (!String.IsNullOrEmpty(newest)) selected = newest;
                    else selected = ports.LastOrDefault();
                }
                else
                {
                    if (!String.IsNullOrEmpty(newest)) selected = newest;
                    else if (ports.Contains(CurrentSelection)) selected = CurrentSelection;
                    else selected = ports.LastOrDefault();
                }
            }

            // If there was a change to the port list, return the recommended default selection
            return selected;
        }

        private void BTN_SETTIME_Click(object sender, EventArgs e)
        {
            if (comboBox6.Text != "" && comboBox5.Text != "" && comboBox8.Text != "" && comboBox7.Text != "" && comboBox10.Text!="" && comboBox9.Text!="")
            {
                serialPort1.Write("ST," + comboBox6.Text + "," + comboBox5.Text + "," + comboBox8.Text + "," + comboBox7.Text + "," + comboBox10.Text + "," +comboBox9.Text + "\r\n");
            }
            else
            {
                MessageBox.Show("VUI LONG CAI DAT LAI THOI GIAN", "LỖI CÀI ĐẶT THỜI GIAN", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (comboBox11.Text.CompareTo("KIEM TRA SX") == 0)
            {
                saveFileDialog1.InitialDirectory = "D:";
                saveFileDialog1.Title = "Save Excel File";
                saveFileDialog1.FileName = "";
                saveFileDialog1.Filter = "Excel File|*.xlsx";
                if (saveFileDialog1.ShowDialog() != DialogResult.Cancel)
                {
                    Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                    ExcelApp.Application.Workbooks.Add(Type.Missing);
                    //DATA
                    for (int i = 1; i <= dataGridView1.Columns.Count; i++)
                    {
                        ExcelApp.Cells[6, i] = dataGridView1.Columns[i - 1].HeaderText;
                    }
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataGridView1.Columns.Count; j++)
                        {
                            ExcelApp.Cells[i + 7, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                        }
                    }
                    //TIEU DE
                    ExcelApp.Range["A1","H55"].Style.Font.Name = "Times New Roman";
                    ExcelApp.Range["A2", "H55"].Style.Font.Size = 11;
                    ExcelApp.Range["A1"].Style.Font.Size = 13;
                    ExcelApp.Range[ExcelApp.Cells[1, 1], ExcelApp.Cells[1, 8]].Merge();
                    ExcelApp.Cells[1, 1] = "KẾT QUẢ KIỂM TRA BỘ ĐIỀU KHIỂN ICENTER - TRONG QUÁ TRÌNH SẢN XUẤT";
                    ExcelApp.Cells[2, 2] = "Số phiếu :";
                    ExcelApp.Cells[2, 4] = "/SX";
                    ExcelApp.Cells[3, 2] = "Id/Serial của Icenter :";
                    ExcelApp.Cells[4, 2] = "Ngày kiểm tra :";
                    ExcelApp.Cells[5, 2] = "Người kiểm tra :";
                    ExcelApp.Cells[51, 2] = "Người kiểm tra";
                    ExcelApp.Cells[55, 2] = "Đinh Văn Công";
                    ExcelApp.Cells[51, 6] = "Duyệt";
                    ExcelApp.Cells[55, 6] = "Chu Hải";
                    ExcelApp.Range[ExcelApp.Cells[6, 3], ExcelApp.Cells[6, 7]].Merge();
                    ExcelApp.Range[ExcelApp.Cells[7, 3], ExcelApp.Cells[7, 7]].Merge();
                    ExcelApp.Range[ExcelApp.Cells[13, 3], ExcelApp.Cells[13, 7]].Merge();
                    ExcelApp.Range[ExcelApp.Cells[16, 3], ExcelApp.Cells[16, 7]].Merge();
                    ExcelApp.Range[ExcelApp.Cells[19, 3], ExcelApp.Cells[19, 7]].Merge();
                    ExcelApp.Range[ExcelApp.Cells[22, 3], ExcelApp.Cells[22, 7]].Merge();
                    ExcelApp.Range[ExcelApp.Cells[45, 3], ExcelApp.Cells[45, 7]].Merge();
                    ExcelApp.Range[ExcelApp.Cells[48, 3], ExcelApp.Cells[48, 7]].Merge();
                    ExcelApp.ActiveWorkbook.SaveCopyAs(saveFileDialog1.FileName.ToString());
                    ExcelApp.ActiveWorkbook.Saved = true;
                    ExcelApp.Quit();
                    MessageBox.Show("Lưu thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }
        
        private void button8_Click(object sender, EventArgs e)
        {
            System.Data.DataTable table = new System.Data.DataTable();
            //Phieu kiem tra sau SX
            if (comboBox11.Text.CompareTo("KIEM TRA SX") == 0 )
            {
                table.Columns.Add("STT", typeof(string));
                table.Columns.Add("NỘI DUNG KIỂM TRA", typeof(string));
                table.Columns.Add("GIÁ TRỊ", typeof(string));
                table.Columns.Add(null, typeof(string));
                
                table.Columns.Add(null,typeof(string));
                table.Columns.Add(null, typeof(string));
                table.Columns.Add(null, typeof(string));
                table.Columns.Add("ĐÁNH GIÁ", typeof(string));

                table.Rows.Add("1", "Kiểm tra khối nguồn", "", "");
                table.Rows.Add("", "V_Shunt =", "", "");
                table.Rows.Add("", "V_12 = ", "", "");
                table.Rows.Add("", "V_5 = ", "", "");
                table.Rows.Add("", "V_3 = ", "", "");
                table.Rows.Add("", "LED P_L1 =", "", "");

                table.Rows.Add("2", "Kiểm tra khối MCU", "", "");
                table.Rows.Add("", "Kiểm tra TAG:", "", "");
                table.Rows.Add("", "Kiểm tra Flash loader:", "", "");

                table.Rows.Add("3", "Kiểm tra khối truyền thông RS232", "", "");
                table.Rows.Add("", "Kiểm tra cổng RS232_2 :", "", "");
                table.Rows.Add("", "Kiểm tra cổng RS232_1 :", "", "");

                table.Rows.Add("4", "Kiểm tra khối truyền thông RS485", "", "");
                table.Rows.Add("", "Kiểm tra cổng RS485_2 :", "", "");
                table.Rows.Add("", "Kiểm tra cổng RS485_1 :", "", "");

                table.Rows.Add("5", "Kiểm tra khối ngõ vào - ra Relay", "", "");
                table.Rows.Add("", "Thực hiện đóng/ngắt Relay 1 :", "", "");
                table.Rows.Add("", "_Trạng thái Contactor:", "", "");
                table.Rows.Add("", "_Trạng thái Relay ngõ vào :", "", "");
                table.Rows.Add("", "_Trạng thái LED chỉ thị ngõ ra :", "", "");
                table.Rows.Add("", "_Trạng thái LED chỉ thị ngõ vào :", "", "");
                table.Rows.Add("", "_Trạng thái dữ liệu trên terminal :", "", "");
                table.Rows.Add("", "Thực hiện đóng/ngắt Relay 2 :", "", "");
                table.Rows.Add("", "_Trạng thái Contactor:", "", "");
                table.Rows.Add("", "_Trạng thái Relay ngõ vào", "", "");
                table.Rows.Add("", "_Trạng thái LED chỉ thị ngõ ra.", "", "");
                table.Rows.Add("", "_Trạng thái LED chỉ thị ngõ vào.", "", "");
                table.Rows.Add("", "_Trạng thái dữ liệu trên terminal.", "", "");
                table.Rows.Add("", "Thực hiện đóng/ngắt Relay 3 :", "", "");
                table.Rows.Add("", "_Trạng thái Contactor:", "", "");
                table.Rows.Add("", "_Trạng thái Relay ngõ vào", "", "");
                table.Rows.Add("", "_Trạng thái LED chỉ thị ngõ ra.", "", "");
                table.Rows.Add("", "_Trạng thái LED chỉ thị ngõ vào.", "", "");
                table.Rows.Add("", "_Trạng thái dữ liệu trên terminal.", "", "");
                table.Rows.Add("", "Đóng Switch ngõ vào Relay 4 :", "", "");
                table.Rows.Add("", "_Trạng thái Relay ngõ vào", "", "");
                table.Rows.Add("", "_Trạng thái LED chỉ thị ngõ vào.", "", "");
                table.Rows.Add("", "_Trạng thái dữ liệu trên terminal.", "", "");

                table.Rows.Add("6", "Kiểm tra khối RTC", "", "");
                table.Rows.Add("", "Kết quả hoạt động của bộ RTC thứ 1", "", "");
                table.Rows.Add("", "Kết quả hoạt động của bộ RTC thứ 2", "", "");

                table.Rows.Add("7", "Kiểm tra khối HMI", "", "");
                table.Rows.Add("", "", "", "");

                dataGridView1.DataSource = table;
                dataGridView1.Columns[1].Width = 180;
                dataGridView1.Columns[2].Width = 30;
                dataGridView1.Columns[3].Width = 30;
                dataGridView1.Columns[4].Width = 30;
                dataGridView1.Columns[5].Width = 30;
                dataGridView1.Columns[6].Width = 30;
            }
        }
    }
}

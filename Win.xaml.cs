using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Net.Mail;
using System.IO;
using System.Data.SQLite;
using System.Drawing;
using ClosedXML.Excel;
using System.Collections.ObjectModel;
using System.Threading;

namespace DLMonitor
{

    /// <summary>
    /// MainWindow.xaml 的互動邏輯
    /// </summary>
    public partial class Win : Window
    {

        System.Windows.Forms.NotifyIcon notifyIcon;
        public static string strHost = String.Empty;
        public static string strAccount = String.Empty;
        public static string strPwd = String.Empty;
        public static string strFrom = String.Empty;
        public static string strTo = String.Empty;
        public string Check_Oem;
        public int Check_Decimal;

        private List<int> GetData()
        {
            List<int> list = new List<int>
            {
                1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24
            };
            return list;
        }
 
        private List<string> GetData2()
        {
            List<string> list = new List<string>
            {
                "全部","紀錄","異常"
            };
            return list;
        }

        public string  Celsius_Mail, Humidit_Mail,Day_Mail,DL_Open, User, Subject, Body, Mail1, Mail2, Mail3, Mail4, Mail5,Re_Conn,Hour_Log;
        public int Hour,Time,Time_Out,H_Start, H_End, C_Start, C_End,Mode,R_Start,R_End,DL_Start,DL_End,Hour_send,Hour_Time;
        SQLiteConnection DL_conn;
        SQLiteCommand DL_cmd;
        SQLiteDataReader DL_dr;
        ObservableCollection<Excel> data;
        ObservableCollection<Setting> data1;
        public System.Windows.Forms.Timer timer;
        EasyModbus.ModbusClient modbusClient;
        public Error newWindow = new Error("");

        public class Excel
        {
            public string TIME { get; set; }
            public string SITE { get; set; }
            public string NAME { get; set; }
            public string Humidity { get; set; }
            public string Celsius { get; set; }
            public string Message { get; set; }
        }

        public class Setting
        {
            public string TIME { get; set; }
            public string IpAddress { get; set; }
            public string NAME { get; set; }
            public string Mail { get; set; }
            public string Humidity { get; set; }
            public string Celsius { get; set; }
            public string Hour { get; set; }
            public string ReSend { get; set; }
            public string Path { get; set; }
            public string ReConn { get; set; }
        }

        //關閉
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            timer.Stop();
            if (newWindow.Activate() == true)
            {
                newWindow.Close();
            }
            if (File.Exists(TB_Path.Text))
            {
                Data_Insert(TB_NAME.Text.ToString(), "程式關閉", 0, 0);
            }
            LB_Message.Items.Add(DateTime.Now + " [Message] " + "正常關閉");
            //Log存成文字檔
            StreamWriter sw = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + "\\"+ CB_IP.SelectedItem.ToString() + DateTime.Now.ToString("yyyyMMddHHmmss") + ".txt");
            foreach (object item in LB_Message.Items)
            {
                sw.WriteLine(item.ToString());
            }
            sw.Close();
            base.OnClosed(e);
            notifyIcon.Visible = false;
            Application.Current.Shutdown();
        }

        //初始化
        public Win(string title)
        {
            InitializeComponent();
            WindowStyle = WindowStyle.ToolWindow;
            Icon_Set();
            var today = DateTime.Today;
            var tomorrow = today.AddDays(1);
            DP_Start.Text = today.ToLongDateString();
            DP_End.Text = tomorrow.ToLongDateString();
            DP_Start1.Text = today.ToLongDateString();
            DP_End1.Text = tomorrow.ToLongDateString();
            List<int> itemNames = GetData();
            CB_Hour.ItemsSource = itemNames;
            List<string> itemNames2 = GetData2();
            CB_SITE.ItemsSource = itemNames2;
            timer = new System.Windows.Forms.Timer
            {
                Interval = 1000
            };
            timer.Tick += new EventHandler(Timer_Tick);
            Data_Set();
            //隱藏一些判斷功能
            GB_Hidden.Visibility = Visibility.Hidden;
            newWindow = new Error("設備異常 時間:" + DateTime.Now.ToString() + "名稱:" + TB_NAME.Text.ToString());
            newWindow.Show();
            newWindow.Visibility = Visibility.Hidden;
            notifyIcon.Visible = true;
            //連線重寄
            Re_Conn = "OFF";
            //多開讀取
            this.Title = title;
            DL_conn = new SQLiteConnection(@"Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\Setting.db" + "");
            DL_conn.Open();
            DL_cmd = DL_conn.CreateCommand();
            DL_cmd.CommandText = " SELECT * FROM Setting WHERE NAME='" + title + "'";
            DL_dr = DL_cmd.ExecuteReader();
            while (DL_dr.Read())
            {
                TB_IP.Text = DL_dr[1].ToString();
                string[] words = TB_IP.Text.ToString().Split('.');
                NUP_IP1.Value = int.Parse(words[0]);
                NUP_IP2.Value = int.Parse(words[1]);
                NUP_IP3.Value = int.Parse(words[2]);
                NUP_IP4.Value = int.Parse(words[3]);
                TB_Port.Text = DL_dr[2].ToString();
                TB_NAME.Text = DL_dr[3].ToString();
                TB_Mail1.Text = DL_dr[4].ToString();
                TB_Mail2.Text = DL_dr[5].ToString();
                TB_Mail3.Text = DL_dr[6].ToString();
                TB_Mail4.Text = DL_dr[7].ToString();
                TB_Mail5.Text = DL_dr[8].ToString();
                NUP_Humidity1.Value = int.Parse(DL_dr[9].ToString());
                NUP_Humidity2.Value = int.Parse(DL_dr[10].ToString());
                NUP_Celsius1.Value = int.Parse(DL_dr[11].ToString());
                NUP_Celsius2.Value = int.Parse(DL_dr[12].ToString());
                Hour = int.Parse(DL_dr[13].ToString());
                foreach (var word in CB_Hour.Items)
                {
                    if (Hour == int.Parse(word.ToString()))
                    {
                        CB_Hour.SelectedItem = word;
                    }
                }
                TB_User.Text = DL_dr[14].ToString();
                PB_Pass.Password = DL_dr[15].ToString();
                TB_Path.Text = DL_dr[16].ToString();
                NUP_ReSend.Value = int.Parse(DL_dr[17].ToString());
                NUP_ReConn.Value = int.Parse(DL_dr[18].ToString());
            }
            DL_dr.Close();
            DL_conn.Close();
            //寄信flag
            Celsius_Mail = "OFF";
            Humidit_Mail = "OFF";
            //每天二次
            Day_Mail = "OFF";
            //每?小時Log
            Time = DateTime.Now.Minute * 60 + DateTime.Now.Second;
            Time_Out = Hour * 60 * 60;
            PBar_Time.Value = Time;
            PBar_Time.Maximum = Time_Out;
            Data_Insert(TB_NAME.Text.ToString(), "程式開始", 0, 0);
            Day_Mail = "ON";
            //設備未連線
            DL_Open = "OFF";
            //執行緒寄信
            User = TB_User.Text.ToString();
            Mail1 = TB_Mail1.Text.ToString();
            Mail2 = TB_Mail2.Text.ToString();
            Mail3 = TB_Mail3.Text.ToString();
            Mail4 = TB_Mail4.Text.ToString();
            Mail5 = TB_Mail5.Text.ToString();
            Subject = "測試主題";
            Body = "測試內容";
            Mode = 0;
            this.notifyIcon.BalloonTipText = title + "溫濕度計";
            this.notifyIcon.Text = title + "溫濕度計";
            CB_IP.SelectedItem = title;
            //Log整點記
            Hour_Log = "OFF";
            Hour_Time = DateTime.Now.Hour + int.Parse(CB_Hour.SelectedItem.ToString());
            if (Hour_Time > 24 )
            {
                Hour_Time -= 24;
            }
            if (Hour_Time == 24)
            {
                Hour_Time = 0;
            }
            timer.Start();
        }

        //隱藏下方功能
        private void Btn_Win_Click(object sender, RoutedEventArgs e)
        {
            if (Btn_Win.Content.ToString() == "-")
            {
                this.Width = 530;
                this.Height = 180;
                Btn_Win.Content = "O";
            }
            else
            {
                this.Width = 1024;
                this.Height = 768;
                Btn_Win.Content = "-";
            }
        }

        //選擇設備
        private void CB_IP_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (CB_IP.SelectedItem.ToString() != this.Title)
            {
                if (CB_IP.SelectedItem.ToString() != "---請選擇設備名稱---")
                {
                    string Show_Win;
                    Show_Win = "OFF";
                    foreach (Window window in Application.Current.Windows)
                    {
                        if (window.Title == CB_IP.SelectedItem.ToString())
                        {
                            Show_Win = "ON";
                        }
                    }
                    if (Show_Win == "OFF")
                    {
                        Win newWindow = new Win(CB_IP.SelectedItem.ToString());
                        newWindow.Show();
                        this.Visibility = Visibility.Hidden;
                    }
                    else
                    {
                        if (CB_IP.SelectedItem.ToString() != this.Title)
                        {
                            foreach (Window window in Application.Current.Windows)
                            {
                                if (window.Title == CB_IP.SelectedItem.ToString())
                                {
                                    window.Visibility = Visibility.Visible;
                                    CB_IP.SelectedItem = CB_IP.SelectedItem.ToString();
                                    this.Visibility = Visibility.Hidden;
                                    return;
                                }
                            }
                        }
                    }
                }
            }
        }

        //設定Icon
        public void Icon_Set()
        {
            this.notifyIcon = new System.Windows.Forms.NotifyIcon
            {
                BalloonTipText = "溫濕度計警訊系統",
                Text = "溫濕度計警訊系統",
                Icon = new Icon("temperature.ico")
            };
            notifyIcon.MouseDoubleClick += OnNotifyIconDoubleClick;
            this.notifyIcon.ShowBalloonTip(1000);
        }

        //顯示
        private void OnNotifyIconDoubleClick(object sender, EventArgs e)
        {
            this.Show();
        }

        //隱藏
        private void Btn_A_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
        }

        //設定
        public void Data_Set()
        {

            try
            {
                //讀資料庫的設定
                if (File.Exists(AppDomain.CurrentDomain.BaseDirectory + "\\Setting.db"))
                {
                    CB_IP.Items.Add("---請選擇設備名稱---");
                    DL_conn = new SQLiteConnection(@"Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\Setting.db" + "");
                    DL_conn.Open();
                    DL_cmd = DL_conn.CreateCommand();
                    DL_cmd.CommandText = " SELECT DISTINCT NAME FROM Setting ";
                    DL_dr = DL_cmd.ExecuteReader();
                    while (DL_dr.Read())
                    {
                        CB_IP.Items.Add(DL_dr[0].ToString());
                    }
                    DL_dr.Close();
                    DL_conn.Close();
                }


            }
            catch (Exception ex)
            {
                LB_Message.Items.Add(DateTime.Now + " [Alarm] " + ex.Message);
            }
        }

        //多執行緒寄信
        private void ModifyUI()
        {
            // 模擬一些工作正在進行
            Thread.Sleep(TimeSpan.FromSeconds(2));
            if (User == "")
            {
                this.LB_Message.Items.Add(DateTime.Now + " [Warn] " + "無寄信者");
            }
            try
            {
                strHost = "ibase.com.tw";   //STMP服务器地址
                strAccount = User + "@ibase.com.tw";       //SMTP服务帐号（自己邮箱）
                strPwd = PB_Pass.Password; //SMTP服务密码（自己邮箱密码）
                strFrom = strAccount; //寄信人
                strTo = strFrom;
                SmtpClient SmtpClient = new SmtpClient
                {
                    DeliveryMethod = SmtpDeliveryMethod.Network,//指定电子邮件发送方式
                    Host = strHost,//指定SMTP服务器
                    Credentials = new System.Net.NetworkCredential(strAccount, strPwd)//用户名和密码
                };
                MailMessage MailMessage = new MailMessage(strFrom, strTo)
                {
                    Subject = Subject,//主题
                    Body = Body,//内容
                    BodyEncoding = Encoding.UTF8,//正文编码
                    IsBodyHtml = true,//设置为HTML格式
                    Priority = MailPriority.High//优先级
                };
                if (Mode != 0)
                {
                    //收件信箱
                    if (Mail1 != "")
                    {
                        MailMessage.To.Add(Mail1);
                    }
                    if (Mail2 != "")
                    {
                        MailMessage.To.Add(Mail2);
                    }
                    if (Mail3 != "")
                    {
                        MailMessage.To.Add(Mail3);
                    }
                    if (Mail4 != "")
                    {
                        MailMessage.To.Add(Mail4);
                    }
                    if (Mail5 != "")
                    {
                        MailMessage.To.Add(Mail5);
                    }
                }
                SmtpClient.Send(MailMessage);
                if (Mode == 0)
                {
                    this.LB_Message.Items.Add(DateTime.Now + " [Message] " + "測試寄信成功");
                }
            }
            catch
            {
                //this.LB_Message.Items.Add(DateTime.Now + " [Alarm] " + ex.Message);
            }

        }

        //計時
        void Timer_Tick(object sender, EventArgs e)
        {
            try
            {
                timer.Stop();
                LB_Time.Content = DateTime.Now.ToString();
                TB_Time.Text = Time.ToString();
                CB_IP.SelectedItem = this.Title;
                // 依設定時間紀錄
                if (DateTime.Now.Hour == Hour_Time)
                {
                    if (Hour_Log == "OFF")
                    {
                        if (TB_Humidity_Read.Text != "")
                        {
                            Data_Insert(TB_NAME.Text.ToString(), "Log", double.Parse(TB_Humidity_Read.Text.ToString()), double.Parse(TB_Celsius_Read.Text.ToString()));
                        }
                        Hour_Log = "ON";
                        Hour_Time = DateTime.Now.Hour + int.Parse(CB_Hour.SelectedItem.ToString());
                        if (Hour_Time > 24)
                        {
                            Hour_Time -= 24;
                        }
                        if (Hour_Time == 24)
                        {
                            Hour_Time = 0;
                        }
                    }
                }
                else
                {
                    Hour_Log = "OFF";
                }
                //每天二次寄信確認程式執行中
                if (DateTime.Now.Hour == 10 || DateTime.Now.Hour == 22)
                {
                    if (Day_Mail == "OFF")
                    {
                        Day_Mail = "ON";
                        Mode = 1;
                        Subject = "程式確認";
                        Body = "時間" + DateTime.Now.ToString() + TB_NAME.Text.ToString() + "濕度目前值：" + TB_Humidity_Read.Text.ToString() + "溫度目前值：" + TB_Celsius_Read.Text.ToString();
                        Thread thread = new Thread(ModifyUI);
                        thread.Start();
                    }
                }
                else
                {
                    Day_Mail = "OFF";
                }

                //取得Mobus
                Mobus(TB_IP.Text.ToString(), int.Parse(TB_Port.Text.ToString()));

                if (TB_Humidity_Read.Text != "")
                {
                    //濕度異常
                    if (double.Parse(TB_Humidity_Read.Text.ToString()) > double.Parse(NUP_Humidity2.Value.ToString()) || double.Parse(TB_Humidity_Read.Text.ToString()) < double.Parse(NUP_Humidity1.Value.ToString()))
                    {
                        TB_Humidity_Read.Background = System.Windows.Media.Brushes.Red;
                        if (H_End > 0)
                        {
                            H_Start += 1;
                            TB_H.Text = H_Start.ToString();
                            if (H_Start >= H_End)
                            {
                                Mode = 1;
                                Subject = "濕度異常";
                                Body = "時間" + DateTime.Now.ToString() + TB_NAME.Text.ToString() + "濕度目前值：" + TB_Humidity_Read.Text.ToString() + "設定值從" + NUP_Humidity1.Value.ToString() + "到" + NUP_Humidity2.Value.ToString();
                                Thread thread = new Thread(ModifyUI);
                                thread.Start();
                                H_Start = 0;
                                H_End = int.Parse(NUP_ReSend.Value.ToString()) * 60;
                            }
                        }
                        if (Humidit_Mail == "OFF")
                        {
                            Data_Insert(TB_NAME.Text.ToString(), "濕度異常", double.Parse(TB_Humidity_Read.Text.ToString()), double.Parse(TB_Celsius_Read.Text.ToString()));
                            Mode = 1;
                            Subject = "濕度異常";
                            Body = "時間" + DateTime.Now.ToString() + TB_NAME.Text.ToString() + "濕度目前值：" + TB_Humidity_Read.Text.ToString() + "設定值從" + NUP_Humidity1.Value.ToString() + "到" + NUP_Humidity2.Value.ToString();
                            Thread thread = new Thread(ModifyUI);
                            thread.Start();
                            Humidit_Mail = "ON";
                            H_Start = 0;
                            H_End = int.Parse(NUP_ReSend.Value.ToString()) * 60;
                        }
                    }
                    else
                    {
                        TB_Humidity_Read.Background = System.Windows.Media.Brushes.LightYellow;
                        Humidit_Mail = "OFF";

                    }
                }

                if (TB_Celsius_Read.Text != "" )
                {
                    //溫度異常
                    if (double.Parse(TB_Celsius_Read.Text.ToString()) > double.Parse(NUP_Celsius2.Value.ToString()) || double.Parse(TB_Celsius_Read.Text.ToString()) < double.Parse(NUP_Celsius1.Value.ToString()))
                    {
                        TB_Celsius_Read.Background = System.Windows.Media.Brushes.Red;
                        if (C_End > 0)
                        {
                            C_Start += 1;
                            TB_C.Text = C_Start.ToString();
                            if (C_Start >= C_End)
                            {
                                Mode = 1;
                                Subject = "溫度異常";
                                Body = "時間" + DateTime.Now.ToString() + TB_NAME.Text.ToString() + "溫度目前值：" + TB_Celsius_Read.Text.ToString() + "設定值：從" + NUP_Celsius1.Value.ToString() + "到" + NUP_Celsius2.Value.ToString();
                                Thread thread = new Thread(ModifyUI);
                                thread.Start();
                                C_Start = 0;
                                C_End = int.Parse(NUP_ReSend.Value.ToString()) * 60;
                            }
                        }
                        if (Celsius_Mail == "OFF")
                        {
                            Data_Insert(TB_NAME.Text.ToString(), "溫度異常", double.Parse(TB_Humidity_Read.Text.ToString()), double.Parse(TB_Celsius_Read.Text.ToString()));
                            Mode = 1;
                            Subject = "溫度異常";
                            Body = "時間" + DateTime.Now.ToString() + TB_NAME.Text.ToString() + "溫度目前值：" + TB_Celsius_Read.Text.ToString() + "設定值：從" + NUP_Celsius1.Value.ToString() + "到" + NUP_Celsius2.Value.ToString();
                            Thread thread = new Thread(ModifyUI);
                            thread.Start();
                            Celsius_Mail = "ON";
                            C_Start = 0;
                            C_End = int.Parse(NUP_ReSend.Value.ToString()) * 60;
                        }
                    }
                    else
                    {
                        TB_Celsius_Read.Background = System.Windows.Media.Brushes.LightYellow;
                        Celsius_Mail = "OFF";
                    }
                }
                timer.Start();
            }
            catch (Exception ex)
            {
                LB_Message.Items.Add(DateTime.Now + " [Alarm] " + ex.Message);
            }
        }

        //連線Mobus
        public void Mobus(String IP, int Port)
        {
            try
            {
                modbusClient = new EasyModbus.ModbusClient
                {
                    IPAddress = IP,
                    Port = Port
                };
                modbusClient.Connect();
                if (modbusClient.Connected == true)
                {
                    R_Start = 0;
                    DL_Start = 0;
                    DL_Open = "OFF";
                    newWindow.Visibility = Visibility.Hidden;
                    IM_DL.Opacity = 1;
                    TB_Status.Background = System.Windows.Media.Brushes.Green;
                    int[] serverResponse = modbusClient.ReadHoldingRegisters(0, 3);
                    double dHumidity = serverResponse[0] / 100.0;
                    double dCelsius = serverResponse[1] / 100.0;
                    double dFahrenheit = serverResponse[2] / 100.0;
                    if (DateTime.Now.Second == 0)
                    {
                        TB_Humidity_Read.Text = dHumidity.ToString();
                        TB_Celsius_Read.Text = dCelsius.ToString();
                        TB_Fahrenheit_Read.Text = dFahrenheit.ToString();
                    }
                }
                modbusClient.Disconnect();
            }
            catch
            {
                //異常發生持續1分才發
                DL_End = 30;
                R_Start += 1;
                if (DL_Open == "OFF")
                {
                    DL_Start += 1;
                    IM_DL.Opacity = 0.1;
                    if (DL_Start >= DL_End)
                    {
                        newWindow.Visibility = Visibility.Visible;
                        TB_Status.Background = System.Windows.Media.Brushes.Red;
                        newWindow.LB_Error.Content = "設備異常 時間:" + DateTime.Now.ToString() + "名稱:" + TB_NAME.Text.ToString();
                        Mode = 1;
                        Subject = "設備異常";
                        Body = "時間" + DateTime.Now.ToString() + TB_NAME.Text.ToString();
                        Thread thread2 = new Thread(ModifyUI);
                        thread2.Start();
                        Data_Insert(TB_NAME.Text.ToString(), "設備異常", 0, 0);
                        R_Start = 0;
                        R_End = int.Parse(NUP_ReConn.Value.ToString()) * 30;
                        DL_Open = "ON";
                        DL_Start = 0;
                    }
                }
                else
                {
                    if (R_End > 0)
                    {
                        if (R_Start >= R_End)
                        {
                            Mode = 1;
                            Subject = "設備異常";
                            Body = newWindow.LB_Error.Content.ToString();
                            Thread thread = new Thread(ModifyUI);
                            thread.Start();
                            R_Start = 0;
                            R_End = int.Parse(NUP_ReConn.Value.ToString()) * 30;
                        }
                    }
                }
            }

        }


        //寄信測試
        private void Btn_Mail_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Mode = 1;
                Subject = "測試";
                Body = "時間" + DateTime.Now.ToString() + TB_NAME.Text.ToString()+ "測試內容";
                Thread thread = new Thread(ModifyUI);
                thread.Start();
            }
            catch (Exception ex)
            {
                LB_Message.Items.Add(DateTime.Now + " [Alarm] " + ex.Message);
            }
           
        }
            
        //儲存設定到Setting資料庫
        private void Btn_Save_Click(object sender, RoutedEventArgs e)
        {
            double Celsius,Fahrenheit;
            TB_IP.Text = NUP_IP1.Value.ToString()+"."+ NUP_IP2.Value.ToString() + "."+ NUP_IP3.Value.ToString() + "."+ NUP_IP4.Value.ToString();
            TB_Humidity.Text = NUP_Humidity1.Value.ToString() + "." + NUP_Humidity2.Value.ToString();
            TB_Celsius.Text = NUP_Celsius1.Value.ToString() + "." + NUP_Celsius2.Value.ToString();
            Celsius = double.Parse(TB_Celsius.Text);
            Fahrenheit = Celsius * (9 / 5) + 32;
            TB_Fahrenheit_Set.Text = Fahrenheit.ToString();

            try
            {
                //資料庫
                DL_conn = new SQLiteConnection(@"Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\Setting.db" + "");
                DL_conn.Open();
                DL_cmd = DL_conn.CreateCommand();
                string Time = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                DL_cmd.CommandText = " INSERT INTO Setting VALUES(datetime('" + Time + "'),'" + TB_IP.Text + "','" + TB_Port.Text + "','" + TB_NAME.Text + "','" + TB_Mail1.Text + "','" + TB_Mail2.Text + "','" + TB_Mail3.Text + "','" + TB_Mail4.Text + "','" + TB_Mail5.Text + "','" + NUP_Humidity1.Value.ToString() + "','" + NUP_Humidity2.Value.ToString() + "','" + NUP_Celsius1.Value.ToString() + "','" + NUP_Celsius2.Value.ToString() + "','" + CB_Hour.SelectedItem.ToString() + "','" + TB_User.Text + "','" + PB_Pass.Password + "','" + TB_Path.Text + "','" + NUP_ReSend.Value + "','" + NUP_ReConn.Value + "')";
                DL_cmd.ExecuteNonQuery();
                DL_conn.Close();
                LB_Message.Items.Add(DateTime.Now + " [Message] " + "已儲存");

            }
            catch (Exception ex)
            {
                LB_Message.Items.Add(DateTime.Now + " [Alarm] " + ex.Message);
            }
            
        }

        //DB的路徑
        private void Btn_Path_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog ofd = new Microsoft.Win32.OpenFileDialog
            {
                DefaultExt = ".db",
                Filter = "db file|*.db"
            };
            if (ofd.ShowDialog() == true)
            {
                TB_Path.Text=ofd.FileName; 
            }
        }

        //設定Excel匯出
        private void Btn_Excel1_Click(object sender, RoutedEventArgs e)
        {
            if (DG_SET.Items.Count == 0)
            {
                LB_Message.Items.Add(DateTime.Now + " [Warn] " + "無資料可匯出");
            }
            else
            {
                try
                {
                    System.Windows.Forms.FolderBrowserDialog path = new System.Windows.Forms.FolderBrowserDialog();
                    path.ShowDialog();
                    if (path.SelectedPath != "")
                    {
                        string filepath = path.SelectedPath + @"\Setting.xlsx";
                        XLWorkbook workbook = new XLWorkbook();
                        var sheet = workbook.Worksheets.Add("Setting");
                        int rowIdx = 2;
                        sheet.Cell(1, 1).Value = "設定日期";
                        sheet.Cell(1, 2).Value = "設備IP";
                        sheet.Cell(1, 3).Value = "設備名稱";
                        sheet.Cell(1, 4).Value = "收件者清單";
                        sheet.Cell(1, 5).Value = "濕度設定";
                        sheet.Cell(1, 6).Value = "溫度設定";
                        sheet.Cell(1, 7).Value = "儲存間隔(時)";
                        sheet.Cell(1, 8).Value = "異常重寄間隔(分)";
                        sheet.Cell(1, 10).Value = "連線間隔(分)";
                        sheet.Cell(1, 9).Value = "資料庫路徑";
                        foreach (var item in data1)
                        {
                            int conlumnIndex = 1;
                            foreach (var jtem in item.GetType().GetProperties())
                            {
                                sheet.Cell(rowIdx, conlumnIndex).Value = string.Concat("'", Convert.ToString(jtem.GetValue(item, null)));
                                conlumnIndex++;
                            }
                            rowIdx++;
                        }
                        workbook.SaveAs(filepath);
                        LB_Message.Items.Add(DateTime.Now + " [Message] " + "已匯出");
                    }

                }
                catch (Exception ex)
                {
                    LB_Message.Items.Add(DateTime.Now + " [Alarm] " + ex.Message);
                }
            }
        }

        //設定查詢
        private void Btn_Query1_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (File.Exists(AppDomain.CurrentDomain.BaseDirectory + "\\Setting.db"))
                {
                    string Start = DP_Start1.SelectedDate.Value.ToString("yyyy-MM-dd");
                    string End = DP_End1.SelectedDate.Value.ToString("yyyy-MM-dd");
                    //資料庫

                    DL_conn = new SQLiteConnection(@"Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "\\Setting.db" + "");
                    DL_conn.Open();
                    DL_cmd = DL_conn.CreateCommand();
                    DL_cmd.CommandText = " SELECT *  FROM Setting WHERE TIME BETWEEN date('" + Start + "') AND  date('" + End + "')";
                    if (TB_NAME_Search.Text.ToString() != "")
                    {
                        DL_cmd.CommandText = DL_cmd.CommandText + " AND NAME LIKE '%" + TB_NAME_Search1.Text.ToString() + "%'";
                    }
                    DL_dr = DL_cmd.ExecuteReader();
                    data1 = new ObservableCollection<Setting>();
                    while (DL_dr.Read())
                    {
                        Setting item = new Setting
                        {
                            TIME = DL_dr[0].ToString(),
                            IpAddress = DL_dr[1].ToString(),
                            NAME = DL_dr[3].ToString(),
                            Mail = "收件者1：" + DL_dr[4].ToString()+ "收件者2：" + DL_dr[5].ToString() + "收件者3：" + DL_dr[6].ToString()+ "收件者4：" + DL_dr[7].ToString() + "收件者5：" + DL_dr[8].ToString(),
                            Humidity = "濕度從：" + DL_dr[9].ToString() + "到" + DL_dr[10].ToString(),
                            Celsius = "溫度從：" + DL_dr[11].ToString() + "到" + DL_dr[12].ToString(),
                            Hour = DL_dr[13].ToString(),
                            ReSend = DL_dr[17].ToString(),
                            ReConn = DL_dr[18].ToString(),
                            Path = DL_dr[16].ToString(),
                        };
                        data1.Add(item);
                    }
                    DL_dr.Close();
                    DL_conn.Close();
                    DG_SET.ItemsSource = data1;
                    if (DG_SET.Items.Count == 0)
                    {
                        LB_Message.Items.Add(DateTime.Now + " [Warn] " + "查無資料");
                    }
                }
                else
                {
                    LB_Message.Items.Add(DateTime.Now + " [Warn] " + "無資料庫");
                }
            }
            catch (Exception ex)
            {
                LB_Message.Items.Add(DateTime.Now + " [Alarm] " + ex.Message);
            }
        }

        //紀錄Excel匯出
        private void Btn_Excel_Click(object sender, RoutedEventArgs e)
        {
            if (DG_LOG.Items.Count == 0)
            {
                //MessageBox.Show("無資料可匯出", "訊息", MessageBoxButton.OK, MessageBoxImage.Warning);
                LB_Message.Items.Add(DateTime.Now + " [Warn] " + "無資料可匯出");
            }
            else
            {
                try
                {
                    System.Windows.Forms.FolderBrowserDialog path = new System.Windows.Forms.FolderBrowserDialog();
                    path.ShowDialog();
                    if (path.SelectedPath != "")
                    {
                        string filepath = path.SelectedPath + @"\Report.xlsx";
                        XLWorkbook workbook = new XLWorkbook();
                        var sheet = workbook.Worksheets.Add("Report");
                        int rowIdx = 2;
                        sheet.Cell(1, 1).Value = "紀錄日期";
                        sheet.Cell(1, 2).Value = "紀錄類型";
                        sheet.Cell(1, 3).Value = "設備名稱";
                        sheet.Cell(1, 4).Value = "濕度";
                        sheet.Cell(1, 5).Value = "溫度攝氏";
                        foreach (var item in data)
                        {
                            int conlumnIndex = 1;
                            foreach (var jtem in item.GetType().GetProperties())
                            {
                                sheet.Cell(rowIdx, conlumnIndex).Value = string.Concat("'", Convert.ToString(jtem.GetValue(item, null)));
                                conlumnIndex++;
                            }
                            rowIdx++;
                        }
                        workbook.SaveAs(filepath);
                        LB_Message.Items.Add(DateTime.Now + " [Message] " + "已匯出");
                    }

                }
                catch (Exception ex)
                {
                    LB_Message.Items.Add(DateTime.Now + " [Alarm] " + ex.Message);
                }
            }
        }

        //紀錄查詢
        private void Btn_Query_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (File.Exists(TB_Path.Text))
                {
                    string Start = DP_Start.SelectedDate.Value.ToString("yyyy-MM-dd");
                    string End = DP_End.SelectedDate.Value.ToString("yyyy-MM-dd");
                    //資料庫

                    DL_conn = new SQLiteConnection(@"Data Source=" + TB_Path.Text + "");
                    DL_conn.Open();
                    DL_cmd = DL_conn.CreateCommand();
                    DL_cmd.CommandText = " SELECT datetime(TIME) as 紀錄日期,SITE as 紀錄類型,NAME as 設備名稱,Humidity as 濕度,Celsius as 溫度攝氏,Message as 設定範圍  FROM Thermometer_Log WHERE TIME BETWEEN date('" + Start + "') AND  date('" + End + "')";
                    if (TB_NAME_Search.Text.ToString() != "")
                    {
                        DL_cmd.CommandText = DL_cmd.CommandText + " AND NAME LIKE '%" + TB_NAME_Search.Text.ToString() + "%'";
                    }
                    if (CB_SITE.SelectedItem.ToString() != "全部")
                    {
                        if (CB_SITE.SelectedItem.ToString() == "紀錄")
                        {
                            DL_cmd.CommandText += " AND SITE = 'Log' ";
                        }
                        else
                        {
                            DL_cmd.CommandText += " AND SITE LIKE '%異常' ";
                        }
                    }
                    DL_cmd.CommandText += " ORDER BY datetime(TIME),NAME ";
                    DL_dr = DL_cmd.ExecuteReader();
                    data = new ObservableCollection<Excel>();
                    while (DL_dr.Read())
                    {
                        Excel item = new Excel
                        {
                            TIME = DL_dr[0].ToString(),
                            SITE = DL_dr[1].ToString(),
                            NAME = DL_dr[2].ToString(),
                            Humidity = DL_dr[3].ToString(),
                            Celsius = DL_dr[4].ToString(),
                            Message = DL_dr[5].ToString(),
                        };
                        data.Add(item);
                    }
                    DL_dr.Close();
                    DL_conn.Close();
                    DG_LOG.ItemsSource = data;
                    if (DG_LOG.Items.Count == 0)
                    {
                        LB_Message.Items.Add(DateTime.Now + " [Warn] " + "查無資料");
                    }
                }
                else
                {
                    LB_Message.Items.Add(DateTime.Now + " [Warn] " + "無資料");
                }
            }
            catch (Exception ex)
            {
                LB_Message.Items.Add(DateTime.Now + " [Alarm] " + ex.Message);
            }

        }

        //Insert資料庫
        void Data_Insert(string NAME,string SITE,double Humidity,double Celsius)
        {
            try
            {
                if (File.Exists(TB_Path.Text))
                {
                    string Setting;
                    Setting = "";
                    if (NUP_Humidity1.Value > 0)
                    {
                        Setting = "濕度設定值從" + NUP_Humidity1.Value.ToString() + "到" + NUP_Humidity2.Value.ToString() + ",溫度設定值從" + NUP_Celsius1.Value.ToString() + "到" + NUP_Celsius2.Value.ToString();
                    }
                    string Time = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    DL_conn = new SQLiteConnection(@"Data Source=" + TB_Path.Text + "");
                    DL_conn.Open();
                    DL_cmd = DL_conn.CreateCommand();
                    DL_cmd.CommandText = " INSERT INTO Thermometer_Log VALUES(datetime('" + Time + "'),'" + NAME + "','" + SITE + "','" + Humidity + "','" + Celsius + "','" + Setting + "')";
                    DL_cmd.ExecuteNonQuery();
                    DL_conn.Close();
                }
                else
                {
                    LB_Message.Items.Add(DateTime.Now + " [Warn] " + "無資料庫");
                }

            }
            catch (Exception ex)
            {
                LB_Message.Items.Add(DateTime.Now + " [Alarm] " + ex.Message);
            }
            finally
            {
                DL_cmd.Dispose();
                DL_conn.Close();
            }
        }

    }
}

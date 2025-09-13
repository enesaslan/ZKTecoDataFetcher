using System;
using System.Collections.Generic;
using System.Windows.Forms;
using zkemkeeper;
using ClosedXML.Excel;
using System.IO;

namespace zktestconnection
{
    public partial class Form1 : Form
    {
        private CZKEM cZKEM;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.FormBorderStyle = FormBorderStyle.FixedSingle; // pencere sabit boyut
            this.MaximizeBox = false;
            richTextBox1.Clear();
            dataGridView1.Columns.Clear();
            try
            {
                cZKEM = new CZKEM();
                string sdkVersion = "";

                if (cZKEM.GetSDKVersion(ref sdkVersion))
                {
                    label1.Text = "SDK Version: " + sdkVersion;
                }
                else
                {
                    label1.Text = "SDK Version could not read";
                }
            }
            catch (Exception ex)
            {
                label1.Text = "Error: " + ex.Message;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string ip = textBox1.Text.Trim();
            if (!int.TryParse(textBox2.Text.Trim(), out int port))
            {
                MessageBox.Show("Please enter a valid port number");
                return;
            }

            LogEkle($"Connecting... IP: {ip}, Port: {port}");

            try
            {
                cZKEM = (CZKEM)Activator.CreateInstance(
                    Type.GetTypeFromCLSID(new Guid("00853A19-BD51-419B-9269-2DABE57EB61F")));

                if (!cZKEM.Connect_Net(ip, port))
                {
                    LogEkle("Oppps! Something Went Wrong");
                    MessageBox.Show("Could not connect to the device!");
                    return;
                }

                LogEkle("Connection successful");
                int machineNumber = 1;

                dataGridView1.Columns.Clear();
                dataGridView1.Columns.Add("EnrollNumber", "User ID");
                dataGridView1.Columns.Add("VerifyMode", "VerifyMode");
                dataGridView1.Columns.Add("InOutMode", "InOutMode");
                dataGridView1.Columns.Add("DateTime", "DateTime");
                dataGridView1.Columns.Add("WorkCode", "WorkCode");
                dataGridView1.Columns.Add("Name", "Name");
                dataGridView1.Columns.Add("Password", "Password");
                dataGridView1.Columns.Add("Privilege", "Privilege");
                dataGridView1.Columns.Add("Enabled", "EnabledDisabled");
                dataGridView1.Columns.Add("FingerIndex", "FingerIndex");
                dataGridView1.Columns.Add("TmpLength", "TmpLength");
                dataGridView1.Columns.Add("Flag", "Flag");

                // Kullanıcı bilgilerini dictionary içine al
                Dictionary<string, (string name, string password, int privilege, bool enabled)> userInfoDict
                    = new Dictionary<string, (string, string, int, bool)>();

                string enrollNumber, name, password;
                int privilege;
                bool enabled;

                if (cZKEM.ReadAllUserID(machineNumber))
                {
                    while (cZKEM.SSR_GetAllUserInfo(machineNumber, out enrollNumber, out name, out password, out privilege, out enabled))
                    {
                        if (!userInfoDict.ContainsKey(enrollNumber))
                        {
                            userInfoDict[enrollNumber] = (name, password, privilege, enabled);
                        }
                    }
                }

                // Logları oku
                if (cZKEM.ReadAllGLogData(machineNumber))
                {
                    dataGridView1.Rows.Clear();
                    int workCode = 0;
                    int verifyMode, inOutMode;
                    int year, month, day, hour, minute, second;
                    int satirSayisi = 0;

                    while (cZKEM.SSR_GetGeneralLogData(
                        machineNumber,
                        out enrollNumber,
                        out verifyMode,
                        out inOutMode,
                        out year, out month, out day,
                        out hour, out minute, out second,
                        ref workCode))
                    {
                        string tarihSaat = $"{year}-{month:D2}-{day:D2} {hour:D2}:{minute:D2}:{second:D2}";

                        // Kullanıcı bilgilerini dictionary’den al
                        string uName = "", uPass = "";
                        string uPriv = "", uEnabled = "";

                        if (userInfoDict.ContainsKey(enrollNumber))
                        {
                            var info = userInfoDict[enrollNumber];
                            uName = info.name;
                            uPass = info.password;
                            uPriv = info.privilege.ToString();
                            uEnabled = info.enabled ? "Yes" : "No";
                        }

                        int fingerIndex = -1;
                        int flag = 0;
                        int tmpLength = 0;
                        string tmpData;
                        if (cZKEM.GetUserTmpExStr(machineNumber, enrollNumber, 0, out flag, out tmpData, out tmpLength))
                        {
                            fingerIndex = 0;
                        }

                        // Satırı ekle
                        dataGridView1.Rows.Add(
                            enrollNumber,
                            verifyMode,
                            inOutMode,
                            tarihSaat,
                            workCode,
                            uName,
                            uPass,
                            uPriv,
                            uEnabled,
                            fingerIndex == -1 ? "" : fingerIndex.ToString(),
                            tmpLength == 0 ? "" : tmpLength.ToString(),
                            flag.ToString()
                        );

                        satirSayisi++;
                    }

                    LogEkle($"Total of {satirSayisi} log rows were retrieved.");
                }
                else
                {
                    LogEkle("Log data could not be read.");
                    MessageBox.Show("Log data could not be read!");
                }
            }
            catch (Exception ex)
            {
                LogEkle("Error: " + ex.Message);
                MessageBox.Show("An error occurred: " + ex.Message);
            }
        }

        private void LogEkle(string mesaj)
        {
            richTextBox1.AppendText($"[{DateTime.Now:HH:mm:ss}] {mesaj}\n");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("No data found to export.");
                return;
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel |*.xlsx";
            saveFileDialog.Title = "Excel";
            saveFileDialog.FileName = "log.xlsx";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Logs");

                        // Başlıklar
                        for (int i = 0; i < dataGridView1.Columns.Count; i++)
                        {
                            worksheet.Cell(1, i + 1).Value = dataGridView1.Columns[i].HeaderText;
                        }

                        // Satırlar
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            for (int j = 0; j < dataGridView1.Columns.Count; j++)
                            {
                                worksheet.Cell(i + 2, j + 1).Value = dataGridView1.Rows[i].Cells[j].Value?.ToString();
                            }
                        }

                        workbook.SaveAs(saveFileDialog.FileName);
                    }

                    MessageBox.Show("Excel file has been successfully created!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            MessageBox.Show("• It can read data from ZKTeco branded devices and export it\n\n• It works with TCP protocol\n\n\n Created By Enes Aslan | 2025", "ZktEco Data Fetcher");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
        }

    }
}

using NLog;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using VehicleRegistrationReporter.DataApi;
using Newtonsoft.Json;

namespace VehicleRegistrationReporter
{
    public partial class MainForm : Form
    {
        // 界面信息显示的最大行数设置，默认设置为： 1000.
        private const int MAX_ROWS_SHOW_SCREEN = 100;
        // 服务间隔运行时间的单位，默认会设置为 1000 * 60，表示设置的数字为分钟。
        private const int INTERVAL_UNIT = 1000 * 60;
        private Logger logger = LogManager.GetCurrentClassLogger();
        public MainForm()
        {
            InitializeComponent();
            // Load the appsettings from App.config file.
            logger.Log(LogLevel.Info, "程序启动。");
            logger.Log(LogLevel.Info, "加载服务设置信息。");
        }

        private void btnInvoke_Click(object sender, EventArgs e)
        {
            //var text = "http://127.0.0.1:8080/projectName/services/writeObjectOut";
            //var crc = ApiHelper.GenerateClearCRC(text);
            //rtx.Text = crc;

            var testDataStr = rtxData.Text;
            WeightData testData;

            try
            {
                testData = JsonConvert.DeserializeObject<WeightData>(testDataStr);
            }
            catch (Exception ex)
            {
                logger.Log(LogLevel.Error, ex.Message);
                MessageBox.Show("请检查提交的【测试数据】格式是否正确！");
                return;
            }

            var url = txtUrl.Text;
            if(string.IsNullOrEmpty(url) ) {
                MessageBox.Show("请输入正确的【接口地址】！");
                return;
            }
            try
            {
                var uri = new Uri(url);
            }catch (Exception ex)
            {
                MessageBox.Show("请输入正确的【接口地址】！");
                return;
            }

            using (var apiHelper = new ApiHelper(url, WriteLog))
            {
                var jkId = txtId.Text;
                var jkYhm = txtName.Text;
                var jkSqm = txtAuthCode.Text;

                try
                {

                    apiHelper.WriteObjectOut(jkId, jkYhm, jkSqm, testData);
                }
                catch(Exception ex)
                {
                    WriteLog(ex.Message);
                }
            }

        }

        private void WriteLog(string msg)
        {
            // Update the info text to screen.
            Invoke(new Action(() =>
            {
                WriteLogToScreen(msg);
            }));
        }

        // 显示运行的数据到屏幕
        private void WriteLogToScreen(string value)
        {
            // 如果界面显示的信息数太多，例如超过了1000行，那么就从前面移除多出的行数。
            if (rtx.Lines.Length > MAX_ROWS_SHOW_SCREEN)
            {
                string text = rtx.Text;
                string[] lines = rtx.Lines;

                int removeCount = rtx.Lines.Length - MAX_ROWS_SHOW_SCREEN;

                if (removeCount > 0)
                {
                    int newLength = lines.Length - removeCount;
                    string[] newLines = new string[newLength];

                    Array.Copy(lines, removeCount, newLines, 0, newLength);

                    rtx.Lines = newLines;
                }


            }

            // set the rich text box focus. 
            rtx.Focus();
            // set the cursor to the end of the rtx text. 
            rtx.Select(rtx.TextLength, 0);
            // Scroll the control to the cursor postion.
            rtx.ScrollToCaret();
            rtx.AppendText(string.Format("[{0:HH:mm:ss.fff}] {1}\r\n", DateTime.Now, value));
            logger.Log(LogLevel.Info, value);

        }
    }
}

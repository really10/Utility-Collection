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
            //var str = "ABCDEFGH0099";
            //var key = "0123456789ABCDEF";

            //var chiper = AesEncryption.AesEncrypt(str, key);
            //this.WriteLog(chiper);

            //var text = "http://127.0.0.1:8080/projectName/services/writeObjectOut";
            //var crc = ApiHelper.GenerateClearCRC(text);
            //rtx.Text = crc;

            var aesKey = "0123456789ABCDEF";

            var testDataStr = rtxData.Text;
            WeightData testData;

            try
            {
                //这里是为了方便测试，所以才直接从界面输入的JSON反序列化回来。
                testData = JsonConvert.DeserializeObject<WeightData>(testDataStr);
                //正式使用，可以参考下面的代码。按真实的对象数据进行赋值。
                //======================================
                WeightData data = new WeightData();
                data.InOutType = "in"; //Varchar(10)，进出类型 in-进，out-出
                data.CardNumber = "粤CYX010"; //Varchar(12) 
                data.CardColor = "蓝色"; //Varchar(2)，示例中没有描述，假定是两个汉字
                data.EnterTime = DateTime.Now;
                data.ItemName = "1"; //Varchar(2),货物名称:1-产品; 2-副产品; 3-原辅材料; 4-燃料; 5-其他
                data.NetWeight = 34.56M; //Double(16,2)
                data.GrossWeight = 32.12M; //Double(16,2)
                data.TareWeight = 30.08M; //Double(16,2)
                data.ItemDetail = "货物明细"; //Varchar
                data.CompanyNumber = "企业编号"; //Varchar
                //delayMinute 延迟推送分钟 Int 无须填报 默认值为60，特殊情况可以传该参数，正常情况无需上传改字段
                //======================================

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
            //初始化API对象。如果对应的环境中，没有WriteLog的这种日志打印委托，可以使用 new ApiHelper(url)进行初始化。
            using (var apiHelper = new ApiHelper(url, aesKey, WriteLog))
            {
                var jkId = txtId.Text; //接口标识
                var jkYhm = txtName.Text; //接口用户名
                var jkSqm = txtAuthCode.Text; //接口授权码

                try
                {
                    ResponseData result = apiHelper.WriteObjectOut(jkId, jkYhm, jkSqm, testData);
                    if(result != null)
                    {
                        // 表示调用成功，处理 result 返回的内容。比如 result.Code 等。
                    }
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

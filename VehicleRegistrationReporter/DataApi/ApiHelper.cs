using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VehicleRegistrationReporter.DataApi
{
    public class ApiHelper : IDisposable
    {
        private RestClient client;
        private string apiUrl;
        private string aesKey;
        private Action<string> WriteLog;

        public ApiHelper(string apiUrl, string aesKey, Action<string> writeLog)
        {
            this.client = new RestClient();
            this.apiUrl = apiUrl;
            this.aesKey = aesKey;
            if (writeLog != null)
            {
                this.WriteLog = writeLog;
            }
            else
            { this.WriteLog = DefaultWriteLog; }
        }

        public ApiHelper(string apiUrl, string aesKey) : this(apiUrl, aesKey, null)
        {
        }

        public ResponseData WriteObjectOut(string id, string name, string authCode, WeightData data)
        {
            var jsonData = JsonConvert.SerializeObject(data);
            var crcCode = GenerateClearCRC(jsonData);

            var request = new RestRequest(apiUrl);
            request.Method = Method.Post;
            request.RequestFormat = DataFormat.None;

            request.AddParameter("jkId", id);
            request.AddParameter("jkYhm", name);
            request.AddParameter("jkSqm", authCode);
            request.AddParameter("crcCode", crcCode);

            var jsonDataCipher = AesEncryption.AesEncrypt(jsonData, aesKey);

            request.AddParameter("jsonData", jsonDataCipher);

            WriteLog("--------准备发送数据--------");
            WriteLog($"校验码：{crcCode}");
            WriteLog($"数据：{jsonData}");
            WriteLog($"数据密文：{jsonDataCipher}");

            var response = client.Execute(request);
            var content = response.Content;

            WriteLog($"返回状态：{response.StatusCode}");
            WriteLog($"返回数据：{content}");

            if(!response.IsSuccessful)
            {
                var errors = GetErrorMessages(response.ErrorException);
                foreach( var error in errors )
                {
                    WriteLog($"错误：{error}");
                }
            }
            else
            {
                var result = JsonConvert.DeserializeObject<ResponseData>(content);
                return result;
            }

            return null;
        }


        public void Dispose()
        {
            if (client != null) { client.Dispose(); }
        }

        private List<string> GetErrorMessages(Exception ex)
        {
            var result = new List<string>();
            if (ex != null) 
            { 
                result.Add(ex.Message); 
                ex = ex.InnerException;
                if(ex != null)
                {
                    result.Add(ex.Message);
                    ex = ex.InnerException;
                    if (ex != null)
                    {
                        result.Add(ex.Message);
                    }
                }
            }
            return result;
        }


        // ===============================================================================
        // CRC 校验生成代码。

        public static string GenerateClearCRC(String puchMsg)
        {
            if (string.IsNullOrEmpty(puchMsg))
            {
                return "";
            }
            puchMsg = puchMsg.Replace("\n", "").Replace("\r", "");
            return PaddingLeft(Convert.ToInt32(CheckCRC(puchMsg)).ToString("X"), 4, '0');
        }

        protected static string PaddingLeft(string str, int left, char ch)
        {
            return Padding(str, left, ch, true);
        }

        protected static string Padding(string str, int length, char ch, bool isLeft)
        {
            if (str.Length < length)
            {
                string pad = "";
                for (int i = str.Length; i < length; i++)
                {
                    pad += ch;
                }
                if (isLeft)
                {
                    str = pad + str;
                }
                else
                {
                    str = str + pad;
                }
            }
            return str;
        }

        public static int CheckCRC(string puchMsg)
        {
            int uchCRC = 0xFFFF;
            int usDataLen = puchMsg.Length;
            int i = 0;
            int cursor = 0;
            while (usDataLen > 0)
            {
                uchCRC = uchCRC ^ (puchMsg[cursor++] << 8);
                for (i = 0; i < 8; i++)
                {
                    if (uchCRC % 2 != 0)
                    {
                        uchCRC >>= 1;
                        uchCRC = uchCRC ^ (0xA001);
                    }
                    else
                    {
                        uchCRC >>= 1;
                    }
                }
                usDataLen--;
            }
            return uchCRC;
        }

        protected void DefaultWriteLog(string msg)
        {

        }
    }
}

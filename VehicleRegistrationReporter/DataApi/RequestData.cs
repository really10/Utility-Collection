using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace VehicleRegistrationReporter.DataApi
{
    public class RequestData
    {
        /// <summary>
        /// 接口标识 
        /// </summary>
        [JsonProperty("jkId")]
        public string Id { get; set; }

        /// <summary>
        /// 接口用户名 
        /// </summary>
        [JsonProperty("jkYhm")]
        public string Name { get; set; }

        /// <summary>
        /// 接口授权码
        /// </summary>
        [JsonProperty("jkSqm")]
        public string AuthCode { get; set; }

        /// <summary>
        /// 交换验证码
        /// </summary>
        [JsonProperty("crcCode")]
        public string CRCCode { get; set; }

        /// <summary>
        /// 查询条件
        /// </summary>
        [JsonProperty("jsonData")]
        public string JsonData { get; set; }
    }
}

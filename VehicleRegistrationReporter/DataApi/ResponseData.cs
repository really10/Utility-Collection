using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VehicleRegistrationReporter.DataApi
{
    public class ResponseData
    {
        /// <summary>
        /// 交易类型
        /// </summary>
        [JsonProperty("exchangeType")]
        public string ExchangeType { get; set; }

        /// <summary>
        /// 交换标识码
        /// </summary>
        [JsonProperty("exchangeCode")]
        public string ExcahnageCode { get; set; }

        /// <summary>
        /// 请求时间
        /// </summary>
        [JsonProperty("responseTime")]
        public string ResponseTime { get; set; }

        /// <summary>
        /// 接口标识
        /// </summary>
        [JsonProperty("jkId")]
        public string JkId { get; set; }

        /// <summary>
        /// 回执结果
        /// </summary>
        [JsonProperty("code")]
        public string Code { get; set; }

        /// <summary>
        /// 结果说明
        /// </summary>
        [JsonProperty("message")]
        public string Message { get; set; }

        /// <summary>
        /// 交换版本
        /// </summary>
        [JsonProperty("version")]
        public string Version { get; set; }
    }
}

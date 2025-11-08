using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VehicleRegistrationReporter.DataApi
{
    public class WeightData
    {
        /// <summary>
        /// 进出类型
        /// </summary>
        [JsonProperty("jclx")]
        public string InOutType { get; set; }

        /// <summary>
        /// 车牌颜色
        /// </summary>
        [JsonProperty("hpys")]
        public string CardColor { get; set;}

        /// <summary>
        /// 车牌号码
        /// </summary>
        [JsonProperty("hphm")]

        public string CardNumber { get; set; }

        /// <summary>
        /// 进场时间
        /// </summary>
        [JsonProperty("jcsj")]
        [JsonConverter(typeof(DateTimeJsonConverter))]

        public DateTime EnterTime { get; set; }

        /// <summary>
        /// 货物名称
        /// </summary>
        [JsonProperty("hwmc")]

        public string ItemName { get; set; }


        /// <summary>
        /// 货物重量
        /// </summary>
        [JsonProperty("ysljz")]

        public decimal NetWeight { get; set; }


        /// <summary>
        /// 货物毛重
        /// </summary>
        [JsonProperty("yslmz")]

        public decimal GrossWeight { get; set; }

        /// <summary>
        /// 货物皮重
        /// </summary>
        [JsonProperty("yslpz")]

        public decimal TareWeight { get; set; }

        /// <summary>
        /// 货物明细
        /// </summary>
        [JsonProperty("hwmx")]

        public string ItemDetail { get; set; }

        /// <summary>
        /// 企业编号
        /// </summary>
        [JsonProperty("qybh")]

        public string CompanyNumber { get; set; }
    }
}

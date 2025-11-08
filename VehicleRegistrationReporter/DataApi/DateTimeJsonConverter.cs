using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VehicleRegistrationReporter.DataApi
{
    public class DateTimeJsonConverter : JsonConverter
    {
        private const string DATETIME_FORMATE = "yyyy-MM-dd HH:mm:ss";
        public override bool CanConvert(Type objectType)
        {
            return objectType == typeof(DateTime);
        }

        public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
        {
            if (reader.Value == null)
            {
                return default(DateTime);
            }
            DateTime value = DateTime.ParseExact(reader.Value.ToString(), DATETIME_FORMATE, CultureInfo.InvariantCulture);
            return value;
        }

        public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
        {
            DateTime? dateTime = (DateTime?)value;
            if (dateTime == null)
            {
                writer.WriteNull();
            }
            else
            {
                writer.WriteValue(dateTime.Value.ToString(DATETIME_FORMATE));
            }
        }

    }
}

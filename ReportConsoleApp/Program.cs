using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Utility;

namespace ReportConsoleApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            ReportUtil reportUtil = new ReportUtil();

            string title1 = "镇燕子浜河道整治工程汇总"; //报表标题
            string filePath1 = "D:\\t1.xlsx"; //报表保存路径
            //报表测试数据
            var dt1 = BuildTotalReportData();

            reportUtil.BuildTotalReport(filePath1, title1, dt1);

            Console.WriteLine("Report Start =========>");
        }

        static DataTable BuildTotalReportData()
        {
            DataTable dt1 = new DataTable();
            //标题
            dt1.Columns.Add("日期", typeof(string));
            dt1.Columns.Add("毛重（公斤）", typeof(int));
            dt1.Columns.Add("皮重（公斤）", typeof(int));
            dt1.Columns.Add("净重（公斤）", typeof(int));
            dt1.Columns.Add("土方数（米3）", typeof(int));
            dt1.Columns.Add("车数", typeof(int));
            dt1.Columns.Add("备注", typeof(string));

            //数据
            dt1.Rows.Add(GetRow(dt1, "2025年8月6日", 1444540, 492180, 952360, 528, 44, null));
            dt1.Rows.Add(GetRow(dt1, "2025年8月7日", 427940, 145720, 282220, 156, 13, null));
            dt1.Rows.Add(GetRow(dt1, "2025年8月8日", 643580, 222000, 421580, 235, 20, null));
            dt1.Rows.Add(GetRow(dt1, "2025年8月9日", 817420, 274240, 543180, 302, 25, null));
            dt1.Rows.Add(GetRow(dt1, "2025年8月10日", 1021440, 330120, 691320, 384, 30, null));

            return dt1;

            DataRow GetRow(DataTable dt, string data, int gross, int tare, int net, int cubic, int num, string remark )
            {
               var row = dt.NewRow();
                row[0] = data;
                row[1] = gross; 
                row[2] = tare;
                row[3] = net;
                row[4] = cubic;
                row[5] = num;
                row[6] = remark;
                return row;
            }
        }
    }
}

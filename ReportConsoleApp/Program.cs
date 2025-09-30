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

            string title2s1 = "8月6日镇燕子浜河道整治工程汇总"; //报表标题-第一个Sheet
            string title2s2 = "镇燕子浜河道整治工程8月6日土方数量统计明细表"; //报表标题-第二个Sheet
            string filePath2 = "D:\\t2.xlsx"; //报表保存路径
            //报表测试数据
            var dt2 = BuildDayReportData();

            reportUtil.BuildDayReport(filePath2, title2s1, title2s2, dt2);

            Console.WriteLine("Report Start =========>");
        }

        //手动创建测试数据，用于TotalReport
        static DataTable BuildTotalReportData()
        {
            DataTable dt1 = new DataTable();
            //标题
            dt1.Columns.Add("日期", typeof(string));
            dt1.Columns.Add("毛重（公斤）", typeof(int));
            dt1.Columns.Add("皮重（公斤）", typeof(int));
            dt1.Columns.Add("净重（公斤）", typeof(int));
            dt1.Columns.Add("土方数（立方米）", typeof(decimal));
            dt1.Columns.Add("车数", typeof(int));
            dt1.Columns.Add("备注", typeof(string));

            //数据
            dt1.Rows.Add(GetRow(dt1, "2025年8月6日", 1444540, 492180, 952360, 528.5M, 44, null));
            dt1.Rows.Add(GetRow(dt1, "2025年8月7日", 427940, 145720, 282220, 156.0M, 13, null));
            dt1.Rows.Add(GetRow(dt1, "2025年8月8日", 643580, 222000, 421580, 235.0M, 20, null));
            dt1.Rows.Add(GetRow(dt1, "2025年8月9日", 817420, 274240, 543180, 302.3M, 25, null));
            dt1.Rows.Add(GetRow(dt1, "2025年8月10日", 1021440, 330120, 691320, 384.0M, 30, null));

            return dt1;

            DataRow GetRow(DataTable dt, string data, int gross, int tare, int net, decimal cubic, int num, string remark)
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

        //手动创建测试数据，用于DayReport
        static DataTable BuildDayReportData()
        {
            DataTable dt1 = new DataTable();
            //标题
            dt1.Columns.Add("序号", typeof(string));
            dt1.Columns.Add("车号", typeof(string));
            dt1.Columns.Add("毛重", typeof(int));
            dt1.Columns.Add("皮重", typeof(int));
            dt1.Columns.Add("净重", typeof(int));
            dt1.Columns.Add("毛重时间", typeof(DateTime));
            dt1.Columns.Add("皮重时间", typeof(DateTime));
            dt1.Columns.Add("土方数", typeof(decimal));
            dt1.Columns.Add("发货单位", typeof(string));

            //数据
            dt1.Rows.Add(GetRow(dt1, 73100, "浙F80961", 27800, 11120, 16680, new DateTime(2025, 8, 6, 16, 45, 0), new DateTime(2025, 8, 6, 16, 55, 0), 3, "镇燕子浜河道整治工程"));
            dt1.Rows.Add(GetRow(dt1, 73092, "浙F80961", 28340, 11160, 17180, new DateTime(2025, 8, 6, 15, 29, 0), new DateTime(2025, 8, 6, 15, 39, 0), 2, "镇燕子浜河道整治工程"));
            dt1.Rows.Add(GetRow(dt1, 73084, "浙F80961", 29520, 11140, 18380, new DateTime(2025, 8, 6, 14, 14, 0), new DateTime(2025, 8, 6, 14, 18, 0), 3, "镇燕子浜河道整治工程"));
            dt1.Rows.Add(GetRow(dt1, 73070, "浙F80961", 28300, 11140, 17160, new DateTime(2025, 8, 6, 12, 57, 0), new DateTime(2025, 8, 6, 13, 12, 0), 4, "镇燕子浜河道整治工程"));
            dt1.Rows.Add(GetRow(dt1, 73056, "浙F80961", 30040, 11200, 18840, new DateTime(2025, 8, 6, 10, 55, 0), new DateTime(2025, 8, 6, 11, 04, 0), 5, "镇燕子浜河道整治工程"));
            dt1.Rows.Add(GetRow(dt1, 73040, "浙F80961", 34600, 11180, 23420, new DateTime(2025, 8, 6, 9, 03, 0), new DateTime(2025, 8, 6, 9, 13, 0), 1, "镇燕子浜河道整治工程"));

            dt1.Rows.Add(GetRow(dt1, 73101, "浙FF5168", 22880, 11140, 11740, new DateTime(2025, 8, 6, 17, 31, 0), new DateTime(2025, 8, 6, 17, 35, 0), 2, "镇燕子浜河道整治工程"));
            dt1.Rows.Add(GetRow(dt1, 73096, "浙FF5168", 28220, 11200, 17020, new DateTime(2025, 8, 6, 16, 09, 0), new DateTime(2025, 8, 6, 16, 12, 0), 3, "镇燕子浜河道整治工程"));
            dt1.Rows.Add(GetRow(dt1, 73088, "浙FF5168", 28460, 11200, 17260, new DateTime(2025, 8, 6, 14, 48, 0), new DateTime(2025, 8, 6, 14, 51, 0), 4, "镇燕子浜河道整治工程"));
            dt1.Rows.Add(GetRow(dt1, 73075, "浙FF5168", 25320, 11180, 14140, new DateTime(2025, 8, 6, 13, 18, 0), new DateTime(2025, 8, 6, 13, 27, 0), 3, "镇燕子浜河道整治工程"));
            dt1.Rows.Add(GetRow(dt1, 73057, "浙FF5168", 28740, 11200, 17540, new DateTime(2025, 8, 6, 11, 24, 0), new DateTime(2025, 8, 6, 11, 41, 0), 2, "镇燕子浜河道整治工程"));
            dt1.Rows.Add(GetRow(dt1, 73042, "浙FF5168", 34160, 11320, 22840, new DateTime(2025, 8, 6, 9, 09, 0), new DateTime(2025, 8, 6, 9, 16, 0), 3, "镇燕子浜河道整治工程"));

            dt1.Rows.Add(GetRow(dt1, 73098, "浙FF6692", 33380, 11120, 22260, new DateTime(2025, 8, 6, 16, 29, 0), new DateTime(2025, 8, 6, 16, 34, 0), 4, "镇燕子浜河道整治工程"));
            dt1.Rows.Add(GetRow(dt1, 73091, "浙FF6692", 32540, 11120, 21420, new DateTime(2025, 8, 6, 15, 21, 0), new DateTime(2025, 8, 6, 15, 27, 0), 5, "镇燕子浜河道整治工程"));
            dt1.Rows.Add(GetRow(dt1, 73074, "浙FF6692", 34480, 11140, 23340, new DateTime(2025, 8, 6, 13, 17, 0), new DateTime(2025, 8, 6, 14, 23, 0), 3, "镇燕子浜河道整治工程"));
            dt1.Rows.Add(GetRow(dt1, 73062, "浙FF6692", 33380, 11240, 22140, new DateTime(2025, 8, 6, 12, 13, 0), new DateTime(2025, 8, 6, 12, 18, 0), 2, "镇燕子浜河道整治工程"));
            dt1.Rows.Add(GetRow(dt1, 73055, "浙FF6692", 34000, 11200, 22800, new DateTime(2025, 8, 6, 10, 43, 0), new DateTime(2025, 8, 6, 10, 48, 0), 4, "镇燕子浜河道整治工程"));
            dt1.Rows.Add(GetRow(dt1, 73039, "浙FF6692", 36780, 11240, 25540, new DateTime(2025, 8, 6, 9, 00, 0), new DateTime(2025, 8, 6, 9, 06, 0), 5, "镇燕子浜河道整治工程"));

            dt1.Rows.Add(GetRow(dt1, 73097, "浙FF8969", 34640, 11160, 23480, new DateTime(2025, 8, 6, 16, 15, 0), new DateTime(2025, 8, 6, 16, 21, 0), 3, "镇燕子浜河道整治工程"));
            dt1.Rows.Add(GetRow(dt1, 73090, "浙FF8969", 32120, 11080, 21040, new DateTime(2025, 8, 6, 15, 01, 0), new DateTime(2025, 8, 6, 15, 06, 0), 2, "镇燕子浜河道整治工程"));
            dt1.Rows.Add(GetRow(dt1, 73081, "浙FF8969", 34160, 11120, 23040, new DateTime(2025, 8, 6, 13, 59, 0), new DateTime(2025, 8, 6, 14, 06, 0), 5, "镇燕子浜河道整治工程"));
            dt1.Rows.Add(GetRow(dt1, 73071, "浙FF8969", 34260, 11160, 23100, new DateTime(2025, 8, 6, 13, 08, 0), new DateTime(2025, 8, 6, 13, 18, 0), 3, "镇燕子浜河道整治工程"));
            dt1.Rows.Add(GetRow(dt1, 73061, "浙FF8969", 34140, 11120, 23020, new DateTime(2025, 8, 6, 11, 55, 0), new DateTime(2025, 8, 6, 12, 03, 0), 2, "镇燕子浜河道整治工程"));
            dt1.Rows.Add(GetRow(dt1, 73053, "浙FF8969", 34020, 11160, 22860, new DateTime(2025, 8, 6, 10, 32, 0), new DateTime(2025, 8, 6, 10, 38, 0), 3, "镇燕子浜河道整治工程"));
            dt1.Rows.Add(GetRow(dt1, 73036, "浙FF8969", 36480, 11180, 25300, new DateTime(2025, 8, 6, 8, 39, 0), new DateTime(2025, 8, 6, 8, 49, 0), 5, "镇燕子浜河道整治工程"));

            dt1.Rows.Add(GetRow(dt1, 73094, "浙FF9060", 35960, 11000, 24960, new DateTime(2025, 8, 6, 15, 55, 0), new DateTime(2025, 8, 6, 15, 59, 0), 4, "镇燕子浜河道整治工程"));
            dt1.Rows.Add(GetRow(dt1, 73087, "浙FF9060", 35000, 11340, 23660, new DateTime(2025, 8, 6, 14, 33, 0), new DateTime(2025, 8, 6, 14, 37, 0), 3, "镇燕子浜河道整治工程"));
            dt1.Rows.Add(GetRow(dt1, 73079, "浙FF9060", 33500, 11060, 22440, new DateTime(2025, 8, 6, 13, 33, 0), new DateTime(2025, 8, 6, 13, 39, 0), 2, "镇燕子浜河道整治工程"));
            dt1.Rows.Add(GetRow(dt1, 73068, "浙FF9060", 34720, 11120, 23600, new DateTime(2025, 8, 6, 12, 37, 0), new DateTime(2025, 8, 6, 12, 41, 0), 3, "镇燕子浜河道整治工程"));
            dt1.Rows.Add(GetRow(dt1, 73060, "浙FF9060", 33820, 11020, 22800, new DateTime(2025, 8, 6, 11, 52, 0), new DateTime(2025, 8, 6, 11, 56, 0), 4, "镇燕子浜河道整治工程"));
            dt1.Rows.Add(GetRow(dt1, 73051, "浙FF9060", 36500, 11100, 25400, new DateTime(2025, 8, 6, 10, 24, 0), new DateTime(2025, 8, 6, 10, 29, 0), 3, "镇燕子浜河道整治工程"));

            dt1.Rows.Add(GetRow(dt1, 73095, "浙FH6265", 35200, 11220, 23980, new DateTime(2025, 8, 6, 16, 03, 0), new DateTime(2025, 8, 6, 16, 07, 0), 5, "镇燕子浜河道整治工程"));
            dt1.Rows.Add(GetRow(dt1, 73089, "浙FH6265", 33820, 11320, 22500, new DateTime(2025, 8, 6, 14, 48, 0), new DateTime(2025, 8, 6, 14, 56, 0), 6, "镇燕子浜河道整治工程"));
            dt1.Rows.Add(GetRow(dt1, 73080, "浙FH6265", 34060, 11460, 22600, new DateTime(2025, 8, 6, 13, 41, 0), new DateTime(2025, 8, 6, 13, 51, 0), 6, "镇燕子浜河道整治工程"));
            dt1.Rows.Add(GetRow(dt1, 73069, "浙FH6265", 33820, 11600, 22220, new DateTime(2025, 8, 6, 12, 54, 0), new DateTime(2025, 8, 6, 12, 59, 0), 5, "镇燕子浜河道整治工程"));
            dt1.Rows.Add(GetRow(dt1, 73059, "浙FH6265", 34200, 11300, 22900, new DateTime(2025, 8, 6, 11, 43, 0), new DateTime(2025, 8, 6, 11, 47, 0), 4, "镇燕子浜河道整治工程"));
            dt1.Rows.Add(GetRow(dt1, 73048, "浙FH6265", 34840, 11460, 23380, new DateTime(2025, 8, 6, 10, 06, 0), new DateTime(2025, 8, 6, 10, 12, 0), 4, "镇燕子浜河道整治工程"));

            dt1.Rows.Add(GetRow(dt1, 73099, "浙FK0021", 32660, 11060, 21600, new DateTime(2025, 8, 6, 16, 35, 0), new DateTime(2025, 8, 6, 16, 41, 0), 5, "镇燕子浜河道整治工程"));
            dt1.Rows.Add(GetRow(dt1, 73093, "浙FK0021", 33900, 11060, 22840, new DateTime(2025, 8, 6, 15, 31, 0), new DateTime(2025, 8, 6, 15, 39, 0), 6, "镇燕子浜河道整治工程"));
            dt1.Rows.Add(GetRow(dt1, 73086, "浙FK0021", 34960, 11080, 23880, new DateTime(2025, 8, 6, 14, 22, 0), new DateTime(2025, 8, 6, 14, 35, 0), 7, "镇燕子浜河道整治工程"));
            dt1.Rows.Add(GetRow(dt1, 73077, "浙FK0021", 32720, 11340, 21380, new DateTime(2025, 8, 6, 13, 26, 0), new DateTime(2025, 8, 6, 13, 35, 0), 4, "镇燕子浜河道整治工程"));
            dt1.Rows.Add(GetRow(dt1, 73064, "浙FK0021", 33720, 11140, 22580, new DateTime(2025, 8, 6, 12, 22, 0), new DateTime(2025, 8, 6, 12, 34, 0), 5, "镇燕子浜河道整治工程"));
            dt1.Rows.Add(GetRow(dt1, 73058, "浙FK0021", 33580, 11100, 22480, new DateTime(2025, 8, 6, 11, 35, 0), new DateTime(2025, 8, 6, 11, 40, 0), 4, "镇燕子浜河道整治工程"));
            dt1.Rows.Add(GetRow(dt1, 73045, "浙FK0021", 36800, 11180, 25620, new DateTime(2025, 8, 6, 9, 32, 0), new DateTime(2025, 8, 6, 9, 38, 0), 4, "镇燕子浜河道整治工程"));
            return dt1;

            DataRow GetRow(DataTable dt, int seq, string carNo, int gross, int tare, int net, DateTime grossTime, DateTime tareTime, int cubic, string unit)
            {
                var row = dt.NewRow();
                row[0] = seq;
                row[1] = carNo;
                row[2] = gross;
                row[3] = tare;
                row[4] = net;
                row[5] = grossTime;
                row[6] = tareTime;
                row[7] = cubic;
                row[8] = unit;
                return row;
            }
        }
    }
}

using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Utility
{
    public class ReportUtil
    {

        public void BuildDayReport(string filePath, string title1, string title2, DataTable dt)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                throw new ArgumentNullException("The filePath cannot be null.");
            }
            if (dt == null)
            {
                throw new ArgumentNullException("The dt cannot be null.");
            }

            //先对DataTable做好排序，按车牌和时间拍好序
            dt.DefaultView.Sort = "车号 ASC, 毛重时间 ASC";
            dt = dt.DefaultView.ToTable();

            var data = ProcessData(dt);


            using (var workbook = new XLWorkbook())
            {
                var ws1 = workbook.AddWorksheet("汇总");

                var summeries = new List<Summary>();
                foreach (var d in data)
                {
                    summeries.Add(d.Summary);
                }

                //设置列宽
                ws1.Column(1).Width = 15;
                ws1.Column(2).Width = 15;
                ws1.Column(3).Width = 15;
                ws1.Column(4).Width = 15;
                ws1.Column(5).Width = 20;
                ws1.Column(6).Width = 15;
                ws1.Column(7).Width = 20;

                var row = 1;
                ws1.Row(row).Height = 40;
                ws1.Range(row, 1, row, 7).Merge();
                var titleCell = ws1.Cell(row, 1);
                titleCell.Value = title1;
                titleCell.Style.Font.SetBold(true);
                titleCell.Style.Font.SetFontSize(16);
                SetHVCenter(titleCell);

                List<String> cn1 = new List<String>() { "车号", "毛重（公斤）", "皮重（公斤）", "净重（公斤）", "土方数（立方米）", "车数", "备注" };

                row = 2;
                ws1.Row(row).Height = 25;
                for (int i = 0; i < 7; i++)
                {

                    var cellName = cn1[i];
                    IXLCell cell = ws1.Cell(row, i + 1);
                    cell.Value = cellName;
                    SetCellBolder(cell);
                    SetFontBold(cell);
                    SetHVCenter(cell);
                }
                //合计记录
                long grossWeights = 0;
                long tareWeights = 0;
                long netWeights = 0;
                decimal cubics = 0;
                long nums = 0;

                row = 3;
                foreach (var s in summeries)
                {
                    ws1.Row(row).Height = 25;

                    RenderSummaryCell(ws1, row, 1, s.CarNumber);
                    RenderSummaryCell(ws1, row, 2, s.GrossWeight);
                    grossWeights += s.GrossWeight;

                    RenderSummaryCell(ws1, row, 3, s.TareWeight);
                    tareWeights += s.TareWeight;

                    RenderSummaryCell(ws1, row, 4, s.NetWeight);
                    netWeights += s.NetWeight;

                    RenderSummaryCell(ws1, row, 5, s.Cubic);
                    cubics += s.Cubic;

                    RenderSummaryCell(ws1, row, 6, s.Num);
                    nums += s.Num;

                    RenderSummaryCell(ws1, row, 7, string.Empty);
                    row += 1;
                }



                ws1.Row(row).Height = 25;
                RenderSummarySum(ws1, row, 1, "合计");
                RenderSummarySum(ws1, row, 2, grossWeights);
                RenderSummarySum(ws1, row, 3, tareWeights);
                RenderSummarySum(ws1, row, 4, netWeights);
                RenderSummarySum(ws1, row, 5, cubics);
                RenderSummarySum(ws1, row, 6, nums);
                RenderSummarySum(ws1, row, 7, string.Empty);

                row += 1;
                // =============
                ws1.Row(row).Height = 25;
                ws1.Range(row, 1, row, 4).Merge();
                var footer1 = ws1.Cell(row, 1);
                footer1.Value = "管理单位：";
                SetFooterCell(footer1);

                var footer2 = ws1.Cell(row, 5);
                footer2.Value = "记录人签字：";
                SetFooterCell(footer2);

                row += 1;
                ws1.Row(row).Height = 25;
                var footer3 = ws1.Cell(row, 5);
                footer3.Value = "审核人签字：";
                SetFooterCell(footer3);

                row += 1;
                ws1.Row(row).Height = 25;
                var footer4 = ws1.Cell(row, 5);
                footer4.Value = "日      期：";
                SetFooterCell(footer4);

                //=================================第二页面==========================
                var ws2 = workbook.AddWorksheet("明细");

                //设置列宽
                ws2.Column(1).Width = 13;
                ws2.Column(2).Width = 13;
                ws2.Column(3).Width = 13;
                ws2.Column(4).Width = 13;
                ws2.Column(5).Width = 13;
                ws2.Column(6).Width = 20;
                ws2.Column(7).Width = 20;
                ws2.Column(8).Width = 13;
                ws2.Column(9).Width = 30;
                ws2.Column(10).Width = 13;

                var r = 1;
                ws2.Row(r).Height = 20;
                ws2.Range(r, 1, r, 10).Merge();
                var titleCell2 = ws2.Cell(r, 1);
                titleCell2.Value = title2;
                titleCell2.Style.Font.SetBold(true);
                SetHVCenter(titleCell2);

                List<String> cn2 = new List<String>() { "序号", "车号", "毛重", "皮重", "净重", "毛重时间", "皮重时间", "土方数", "发货单位", "车数" };

                r = 2;
                ws2.Row(r).Height = 20;
                for (int i = 0; i < 10; i++)
                {

                    var cellName = cn2[i];
                    IXLCell cell = ws2.Cell(r, i + 1);
                    cell.Value = cellName;
                    SetCellBolder(cell);
                    SetFontBold(cell);
                    SetHVCenter(cell);
                }

                r = 3;
                long sumGross = 0;
                long sumTare = 0;
                long sumNet = 0;
                decimal sumCubic = 0;
                long sumNum = 0;
                foreach (var d in data)
                {
                    if (d.Items != null)
                    {
                        foreach (var item in d.Items)
                        {
                            RenderDetailCell(ws2, r, 1, item.Seq);
                            RenderDetailCell(ws2, r, 2, item.CarNumber);
                            RenderDetailCell(ws2, r, 3, item.GrossWeight);
                            RenderDetailCell(ws2, r, 4, item.TareWeight);
                            RenderDetailCell(ws2, r, 5, item.NetWeight);
                            RenderDetailCell(ws2, r, 6, item.GrossWeightTime);
                            RenderDetailCell(ws2, r, 7, item.TareWeightTime);
                            RenderDetailCell(ws2, r, 8, item.Cubic);
                            RenderDetailCell(ws2, r, 9, item.Unit);
                            RenderDetailCell(ws2, r, 10, string.Empty);
                            r++;
                        }
                        ws2.Range(r, 1, r, 2).Merge();
                        RenderSubSum(ws2, r, 1, d.Summary.CarNumber);
                        RenderSubSum(ws2, r, 2, string.Empty);
                        RenderSubSum(ws2, r, 3, d.Summary.GrossWeight);
                        RenderSubSum(ws2, r, 4, d.Summary.TareWeight);
                        RenderSubSum(ws2, r, 5, d.Summary.NetWeight);
                        RenderSubSum(ws2, r, 6, string.Empty);
                        RenderSubSum(ws2, r, 7, string.Empty);
                        RenderSubSum(ws2, r, 8, d.Summary.Cubic);
                        RenderSubSum(ws2, r, 9, string.Empty);
                        RenderSubSum(ws2, r, 10, d.Summary.Num);

                        sumGross += d.Summary.GrossWeight;
                        sumTare += d.Summary.TareWeight;
                        sumCubic += d.Summary.Cubic;
                        sumNet += d.Summary.NetWeight;
                        sumNum += d.Summary.Num;
                        r++;
                    }
                }
                ws2.Range(r, 1, r, 2).Merge();
                RenderSummarySum(ws2, r, 1, "合计：");
                RenderSummarySum(ws2, r, 2, string.Empty);
                RenderSummarySum(ws2, r, 3, sumGross);
                RenderSummarySum(ws2, r, 4, sumTare);
                RenderSummarySum(ws2, r, 5, sumNet);
                RenderSummarySum(ws2, r, 6, string.Empty);
                RenderSummarySum(ws2, r, 7, string.Empty);
                RenderSummarySum(ws2, r, 8, sumCubic);
                RenderSummarySum(ws2, r, 9, string.Empty);
                RenderSummarySum(ws2, r, 10, sumNum);



                workbook.SaveAs(filePath);
            }
        }

        private void RenderSummaryCell(IXLWorksheet ws, int rowNum, int colNum, object value)
        {
            IXLCell cell1 = ws.Cell(rowNum, colNum);
            cell1.Value = XLCellValue.FromObject(value);
            SetCellBolder(cell1);
            SetHVCenter(cell1);
        }

        private void RenderSummarySum(IXLWorksheet ws, int rowNum, int colNum, object value)
        {
            var sum1 = ws.Cell(rowNum, colNum);
            sum1.Value = XLCellValue.FromObject(value);
            SetHVCenter(sum1);
            SetFontBold(sum1);
            SetCellBolder(sum1);
        }

        private void RenderDetailCell(IXLWorksheet ws, int rowNum, int colNum, object value)
        {
            IXLCell cell1 = ws.Cell(rowNum, colNum);
            cell1.Value = XLCellValue.FromObject(value);
            SetCellBolder(cell1);
            SetHVCenter(cell1);
        }

        private void RenderSubSum(IXLWorksheet ws, int rowNum, int colNum, object value)
        {
            IXLCell cell1 = ws.Cell(rowNum, colNum);
            cell1.Value = XLCellValue.FromObject(value);
            SetCellBolder(cell1);
            SetHVCenter(cell1);
            cell1.Style.Fill.SetBackgroundColor(XLColor.Yellow);
        }

        private List<Detail> ProcessData(DataTable dt)
        {
            var results = new List<Detail>();
            string tmpCarNumber = "|"; //为了避免重复
            Detail tmpDetail = null;
            List<Item> tmpItems = null;
            //统计汇总
            Summary tmpSummary = null;
            foreach (DataRow row in dt.Rows)
            {
                var seq = ProcessRowInt(row["序号"]);
                var carNumber = ProcessRowString(row["车号"]);
                var grossWeight = ProcessRowInt(row["毛重"]);
                var tareWeight = ProcessRowInt(row["皮重"]);
                var netWeight = ProcessRowInt(row["净重"]);
                var grossWeightTime = ProcessRowDateTime(row["毛重时间"]);
                var tareWeightTime = ProcessRowDateTime(row["皮重时间"]);
                var unit = ProcessRowString(row["发货单位"]);
                var cubic = ProcessRowDecimal(row["土方数"]);

                //如果是个新车牌，则把上一辆车的数据保存起来。
                if (carNumber != tmpCarNumber)
                {
                    if (tmpDetail != null && tmpItems != null)
                    {
                        //处理汇总
                        tmpDetail.Summary = tmpSummary;
                        //添加一个车辆的对象到列表
                        tmpDetail.Items = tmpItems;
                        results.Add(tmpDetail);
                    }

                    tmpCarNumber = carNumber;
                    tmpSummary = new Summary() { CarNumber = carNumber };
                    tmpDetail = new Detail();
                    tmpItems = new List<Item>();
                }
                var item = new Item();
                item.CarNumber = carNumber;
                item.Seq = seq;
                item.GrossWeight = grossWeight;
                item.TareWeight = tareWeight;
                item.NetWeight = netWeight;
                item.GrossWeightTime = grossWeightTime;
                item.TareWeightTime = tareWeightTime;
                item.Cubic = cubic;
                item.Unit = unit;

                tmpSummary.GrossWeight += grossWeight;
                tmpSummary.TareWeight += tareWeight;
                tmpSummary.NetWeight += netWeight;
                tmpSummary.Cubic += cubic;
                tmpSummary.Num += 1;

                tmpItems.Add(item);
            }

            //添加最后一个对象。
            if (tmpDetail != null && tmpItems != null)
            {
                //处理汇总
                tmpDetail.Summary = tmpSummary;
                //添加一个车辆的对象到列表
                tmpDetail.Items = tmpItems;
                results.Add(tmpDetail);
            }

            return results;
        }

        private int ProcessRowInt(object value)
        {
            if (value != null)
            {
                if (int.TryParse(value.ToString(), out int result))
                {
                    return result;
                }
            }
            return 0;
        }

        private string ProcessRowString(object value)
        {
            if (value != null)
            {
                return value.ToString();
            }
            return string.Empty;
        }

        private DateTime? ProcessRowDateTime(object value)
        {
            if (value is DateTime time)
            {
                return time;
            }
            return null;
        }

        private decimal ProcessRowDecimal(object value)
        {
            if (value != null)
            {
                if (decimal.TryParse(value.ToString(), out decimal result))
                {
                    return result;
                }
            }
            return 0;
        }

        public void BuildTotalReport(string filePath, string title, DataTable dt)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                throw new ArgumentNullException("The filePath cannot be null.");
            }
            if (dt == null)
            {
                throw new ArgumentNullException("The dt cannot be null.");
            }

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.AddWorksheet("汇总");


                //设置列宽
                worksheet.Column(1).Width = 15;
                worksheet.Column(2).Width = 15;
                worksheet.Column(3).Width = 15;
                worksheet.Column(4).Width = 15;
                worksheet.Column(5).Width = 20;
                worksheet.Column(6).Width = 15;
                worksheet.Column(7).Width = 20;

                var row = 1;
                worksheet.Row(row).Height = 40;
                worksheet.Range(row, 1, row, 7).Merge();
                var titleCell = worksheet.Cell(row, 1);
                titleCell.Value = title;
                titleCell.Style.Font.SetBold(true);
                titleCell.Style.Font.SetFontSize(16);
                SetHVCenter(titleCell);


                row = 2;
                worksheet.Row(row).Height = 25;
                for (int i = 0; i < dt.Columns.Count; i++)
                {

                    var cellName = dt.Columns[i].ColumnName;
                    IXLCell cell = worksheet.Cell(row, i + 1);
                    cell.Value = cellName;
                    SetCellBolder(cell);
                    SetFontBold(cell);
                    SetHVCenter(cell);


                }

                row = 3;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    worksheet.Row(row).Height = 25;
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        IXLCell cell = worksheet.Cell(row, j + 1);
                        var value = XLCellValue.FromObject(dt.Rows[i][j]);
                        cell.Value = value;
                        SetCellBolder(cell);
                        SetHVCenter(cell);
                    }
                    row += 1;
                }

                //合计记录
                long grossWeights = SumValue(dt, 1);
                long tareWeights = SumValue(dt, 2);
                long netWeights = SumValue(dt, 3);
                decimal cubics = SumDecimalValue(dt, 4);
                long nums = SumValue(dt, 5);

                worksheet.Row(row).Height = 25;
                var sum1 = worksheet.Cell(row, 1);
                sum1.Value = "合计";
                SetHVCenter(sum1);
                SetFontBold(sum1);
                SetCellBolder(sum1);

                var sum2 = worksheet.Cell(row, 2);
                sum2.Value = grossWeights;
                SetHVCenter(sum2);
                SetFontBold(sum2);
                SetCellBolder(sum2);

                var sum3 = worksheet.Cell(row, 3);
                sum3.Value = tareWeights;
                SetHVCenter(sum3);
                SetFontBold(sum3);
                SetCellBolder(sum3);

                var sum4 = worksheet.Cell(row, 4);
                sum4.Value = netWeights;
                SetHVCenter(sum4);
                SetFontBold(sum4);
                SetCellBolder(sum4);

                var sum5 = worksheet.Cell(row, 5);
                sum5.Value = cubics;
                SetHVCenter(sum5);
                SetFontBold(sum5);
                SetCellBolder(sum5);

                var sum6 = worksheet.Cell(row, 6);
                sum6.Value = nums;
                SetHVCenter(sum6);
                SetFontBold(sum6);
                SetCellBolder(sum6);

                var sum7 = worksheet.Cell(row, 7);
                SetHVCenter(sum7);
                SetFontBold(sum7);
                SetCellBolder(sum7);

                row += 1;
                // =============
                worksheet.Row(row).Height = 25;
                worksheet.Range(row, 1, row, 4).Merge();
                var footer1 = worksheet.Cell(row, 1);
                footer1.Value = "管理单位：";
                SetFooterCell(footer1);

                var footer2 = worksheet.Cell(row, 5);
                footer2.Value = "记录人签字：";
                SetFooterCell(footer2);

                row += 1;
                worksheet.Row(row).Height = 25;
                var footer3 = worksheet.Cell(row, 5);
                footer3.Value = "审核人签字：";
                SetFooterCell(footer3);

                row += 1;
                worksheet.Row(row).Height = 25;
                var footer4 = worksheet.Cell(row, 5);
                footer4.Value = "日      期：";
                SetFooterCell(footer4);

                workbook.SaveAs(filePath);
            }

        }

        private void SetFooterCell(IXLCell cell)
        {
            SetVCenter(cell);
            SetFont(cell);
        }

        private void SetFont(IXLCell cell)
        {
            cell.Style.Font.SetFontName("宋体"); //宋体
        }

        private long SumValue(DataTable dt, int column)
        {
            long sum = 0;
            foreach (DataRow row in dt.Rows)
            {
                var obj = row[column];
                if (obj == null) continue;
                if (int.TryParse(obj.ToString(), out var val))
                {
                    sum += val;
                }
            }
            return sum;
        }
        private decimal SumDecimalValue(DataTable dt, int column)
        {
            decimal sum = 0;
            foreach (DataRow row in dt.Rows)
            {
                var obj = row[column];
                if (obj == null) continue;
                if (decimal.TryParse(obj.ToString(), out var val))
                {
                    sum += val;
                }
            }
            return sum;
        }

        private void SetFontBold(IXLCell cell)
        {
            cell.Style.Font.SetBold(true);
        }

        private void SetCellBolder(IXLCell cell)
        {
            cell.Style.Border.SetTopBorder(XLBorderStyleValues.Thin)
                    .Border.SetRightBorder(XLBorderStyleValues.Thin)
                    .Border.SetBottomBorder(XLBorderStyleValues.Thin)
                    .Border.SetLeftBorder(XLBorderStyleValues.Thin);
        }

        private void SetHVCenter(IXLCell cell)
        {
            SetVCenter(cell);
            SetHCenter(cell);
        }

        private void SetVCenter(IXLCell cell)
        {
            cell.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
        }

        private void SetHCenter(IXLCell cell)
        {
            cell.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
        }

        //=====================================================================
        // 内部对象
        //=====================================================================
        class Detail
        {
            public List<Item> Items { get; set; }
            public Summary Summary { get; set; }
        }

        class Summary
        {
            public string CarNumber { get; set; }
            public int GrossWeight { get; set; }
            public int TareWeight { get; set; }
            public int NetWeight { get; set; }
            public decimal Cubic { get; set; }
            public int Num { get; set; }
            public string remark { get; set; }
        }

        class Item
        {
            public int Seq { get; set; }
            public string CarNumber { get; set; }
            public int GrossWeight { get; set; }
            public int TareWeight { get; set; }
            public int NetWeight { get; set; }

            public DateTime? GrossWeightTime { get; set; }
            public DateTime? TareWeightTime { get; set; }
            public decimal Cubic { get; set; }
            public string Unit { get; set; }
        }

        //class SubSummary
        //{
        //    public string CarNumber { get; set; }
        //    public int GrossWeight { get; set; }
        //    public int TareWeight { get; set; }
        //    public int NetWeight { get; set; }
        //    public int Num { get; set; }
        //}
    }
}
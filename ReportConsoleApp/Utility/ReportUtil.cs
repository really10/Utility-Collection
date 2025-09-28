using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
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
                worksheet.Column(5).Width = 15;
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
                long cubics = SumValue(dt, 4);
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
                SetVCenter(footer1);

                var footer2 = worksheet.Cell(row, 5);
                footer2.Value = "记录人签字：";
                SetVCenter(footer2);

                row += 1;
                worksheet.Row(row).Height = 25;
                var footer3 = worksheet.Cell(row, 5);
                footer3.Value = "审核人签字：";
                SetVCenter(footer3);

                row += 1;
                worksheet.Row(row).Height = 25;
                var footer4 = worksheet.Cell(row, 5);
                footer4.Value = "日      期：";
                SetVCenter(footer4);

                workbook.SaveAs(filePath);
            }

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
            cell.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            cell.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
        }

        private void SetVCenter(IXLCell cell)
        {
            cell.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
        }

        private void SetHCenter(IXLCell cell)
        {
            cell.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
        }
    }
}

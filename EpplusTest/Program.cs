using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EpplusTest
{
    public class Program
    {
        static void Main(string[] args)
        {
            FileInfo newFile = new FileInfo(@"d:\test.xlsx");
            if (newFile.Exists)
            {
                newFile.Delete();
                newFile = new FileInfo(@"d:\test.xlsx");
            }

            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                #region  创建多个Sheet页
                for (int i = 0; i < 5; i++)
                {
                    package.Workbook.Worksheets.Add("Demo" + i);
                }
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Demo0"];
                ExcelWorksheet worksheet1 = package.Workbook.Worksheets["Demo1"];

                #endregion

                #region 1 模拟填充数据
                worksheet1.Cells[1, 1].Value = "名称";
                worksheet1.Cells[1, 2].Value = "价格";
                worksheet1.Cells[1, 3].Value = "销量";

                worksheet1.Cells[2, 1].Value = "苹果";
                worksheet1.Cells[2, 2].Value = 56;
                worksheet1.Cells[2, 3].Value = 100;

                worksheet1.Cells[3, 1].Value = "华为";
                worksheet1.Cells[3, 2].Value = 45;
                worksheet1.Cells[3, 3].Value = 150;

                worksheet1.Cells[4, 1].Value = "小米";
                worksheet1.Cells[4, 2].Value = 38;
                worksheet1.Cells[4, 3].Value = 130;

                worksheet1.Cells[5, 1].Value = "OPPO";
                worksheet1.Cells[5, 2].Value = 22;
                worksheet1.Cells[5, 3].Value = 200;
                #endregion

                #region 2 构造图表
                worksheet.Cells.Style.WrapText = true;
                worksheet.View.ShowGridLines = false;//去掉sheet的网格线
                using (ExcelRange range = worksheet.Cells[1, 1, 5, 3])
                {
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                }

                using (ExcelRange range = worksheet.Cells[1, 1, 1, 3])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Font.Color.SetColor(Color.White);
                    range.Style.Font.Name = "微软雅黑";
                    range.Style.Font.Size = 12;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(128, 128, 128));
                }

                worksheet1.Cells[1, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));
                worksheet1.Cells[1, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));
                worksheet1.Cells[1, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));

                worksheet1.Cells[2, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));
                worksheet1.Cells[2, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));
                worksheet1.Cells[2, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));

                worksheet1.Cells[3, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));
                worksheet1.Cells[3, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));
                worksheet1.Cells[3, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));

                worksheet1.Cells[4, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));
                worksheet1.Cells[4, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));
                worksheet1.Cells[4, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));

                worksheet1.Cells[5, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));
                worksheet1.Cells[5, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));
                worksheet1.Cells[5, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));

                ExcelChart chart = worksheet.Drawings.AddChart("chart", eChartType.ColumnClustered);
                ExcelChartSerie serie = chart.Series.Add(worksheet1.Cells[2, 3, 5, 3], worksheet1.Cells[2, 1, 5, 1]);//引用worksheet1的数据填充图表的X轴和Y轴
                serie.HeaderAddress = worksheet1.Cells[1, 3];
                #endregion

                #region 3 设置图表的样式
                chart.SetPosition(40, 10);
                chart.SetSize(500, 300);
                chart.Title.Text = "销量走势";
                chart.Title.Font.Color = Color.FromArgb(89, 89, 89);
                chart.Title.Font.Size = 15;
                chart.Title.Font.Bold = true;
                chart.Style = eChartStyle.Style15;
                chart.Legend.Border.LineStyle = eLineStyle.SystemDash;
                chart.Legend.Border.Fill.Color = Color.FromArgb(217, 217, 217);
                #endregion
                package.Save();
            }
        }
    }
}
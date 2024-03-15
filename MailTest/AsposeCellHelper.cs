using Aspose.Cells;
using Org.BouncyCastle.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailTest {
    public class AsposeHelper {

        public static void SetLicense() {

            var license_2021_8_27_NetStandard20 = "PExpY2Vuc2U+CiAgPERhdGE+CiAgICA8TGljZW5zZWRUbz5TdXpob3UgQXVuYm94IFNvZnR3YXJlIENvLiwgTHRkLjwvTGljZW5zZWRUbz4KICAgIDxFbWFpbFRvPnNhbGVzQGF1bnRlYy5jb208L0VtYWlsVG8+CiAgICA8TGljZW5zZVR5cGU+RGV2ZWxvcGVyIE9FTTwvTGljZW5zZVR5cGU+CiAgICA8TGljZW5zZU5vdGU+TGltaXRlZCB0byAxIGRldmVsb3BlciwgdW5saW1pdGVkIHBoeXNpY2FsIGxvY2F0aW9uczwvTGljZW5zZU5vdGU+CiAgICA8T3JkZXJJRD4yMDA2MDIwMTI2MzM8L09yZGVySUQ+CiAgICA8VXNlcklEPjEzNDk3NjAwNjwvVXNlcklEPgogICAgPE9FTT5UaGlzIGlzIGEgcmVkaXN0cmlidXRhYmxlIGxpY2Vuc2U8L09FTT4KICAgIDxQcm9kdWN0cz4KICAgICAgPFByb2R1Y3Q+QXNwb3NlLlRvdGFsIGZvciAuTkVUPC9Qcm9kdWN0PgogICAgPC9Qcm9kdWN0cz4KICAgIDxFZGl0aW9uVHlwZT5FbnRlcnByaXNlPC9FZGl0aW9uVHlwZT4KICAgIDxTZXJpYWxOdW1iZXI+OTM2ZTVmZDEtODY2Mi00YWJmLTk1YmQtYzhkYzBmNTNhZmE2PC9TZXJpYWxOdW1iZXI+CiAgICA8U3Vic2NyaXB0aW9uRXhwaXJ5PjIwMjEwODI3PC9TdWJzY3JpcHRpb25FeHBpcnk+CiAgICA8TGljZW5zZVZlcnNpb24+My4wPC9MaWNlbnNlVmVyc2lvbj4KICAgIDxMaWNlbnNlSW5zdHJ1Y3Rpb25zPmh0dHBzOi8vcHVyY2hhc2UuYXNwb3NlLmNvbS9wb2xpY2llcy91c2UtbGljZW5zZTwvTGljZW5zZUluc3RydWN0aW9ucz4KICA8L0RhdGE+CiAgPFNpZ25hdHVyZT5wSkpjQndRdnYxV1NxZ1kyOHFJYUFKSysvTFFVWWRrQ2x5THE2RUNLU0xDQ3dMNkEwMkJFTnh5L3JzQ1V3UExXbjV2bTl0TDRQRXE1aFAzY2s0WnhEejFiK1JIWTBuQkh1SEhBY01TL1BSeEJES0NGbWg1QVFZRTlrT0FxSzM5NVBSWmJRSGowOUNGTElVUzBMdnRmVkp5cUhjblJvU3dPQnVqT1oyeDc4WFE9PC9TaWduYXR1cmU+CjwvTGljZW5zZT4=";            
            var streamCell = new MemoryStream(Convert.FromBase64String(license_2021_8_27_NetStandard20));
            new Aspose.Cells.License().SetLicense(streamCell);
        }

        public static byte[] CreateWorkBook(string sheetName) {
            Workbook wb = new Workbook();
            wb.Worksheets.Clear();
            var sheet = wb.Worksheets.Add(sheetName);

            var titleList = Enumerable.Range(1, 7).ToList().Select(t =>$"标题{t}").ToList();
            var dataList = Enumerable.Range(1,10).ToList().Select(t=>$"测试{t}").ToList();

            //设置一级表头
            var headerStyle = wb.CreateStyle();
            headerStyle.Font.Size = 18;
            headerStyle.Font.Name = "微软雅黑";
            headerStyle.Font.IsBold = true;
            headerStyle.HorizontalAlignment = TextAlignmentType.Center;
            headerStyle.VerticalAlignment = TextAlignmentType.Center;

            sheet.Cells.SetRowHeight(0, 45);
            sheet.Cells[0, 0].SetStyle(headerStyle);
            sheet.Cells[0, 0].PutValue(sheetName);
            sheet.Cells.Merge(firstRow: 0, firstColumn: 0, totalRows: 1, totalColumns: titleList.Count);


            //设置二级表头
            sheet.Cells.SetRowHeight(1, 40);
            sheet.Cells.SetColumnWidth(0, 60);
            sheet.Cells.SetColumnWidth(1, 40);
            sheet.Cells.SetColumnWidth(2, 12);
            sheet.Cells.SetColumnWidth(3, 12);
            sheet.Cells.SetColumnWidth(4, 12);
            sheet.Cells.SetColumnWidth(5, 40);
            sheet.Cells.SetColumnWidth(6, 50);
            sheet.Cells.SetColumnWidth(7, 100);


            var defaultStyle = wb.DefaultStyle;
            defaultStyle.HorizontalAlignment = TextAlignmentType.Center;
            defaultStyle.VerticalAlignment = TextAlignmentType.Center;
            defaultStyle.Font.IsBold = true;

            for (int columnIndex = 0; columnIndex < titleList.Count; columnIndex++) {
                var columnName = titleList[columnIndex];
                sheet.Cells[1, columnIndex].SetStyle(defaultStyle);
                sheet.Cells[1, columnIndex].PutValue(columnName);
            }

            //设置数据
            for (int rowIndex = 0; rowIndex < dataList.Count; rowIndex++) {
                
                for (int columnIndex = 0; columnIndex < titleList.Count; columnIndex++) {

                    var value = dataList[rowIndex]?.ToString() ?? string.Empty;
                    if (columnIndex == titleList.Count - 1) {
                        value = "http://localhost:8000/fulltext.aspx?lib=chl&gid=31023";
                        sheet.Hyperlinks.Add(firstRow: rowIndex + 2, firstColumn: columnIndex, totalRows: 1, totalColumns: 1, address: value);
                    }

                    sheet.Cells[rowIndex + 2, columnIndex].PutValue(value);
                }
            }



            using (MemoryStream ms = new MemoryStream()) {
                wb.Save(ms, SaveFormat.Xlsx);
                return ms.ToArray();
            }
        }
    }
}

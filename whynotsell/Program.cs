using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;

namespace whynotsell
{
    class Program
    {
        static void Main(string[] args)
        {
            DirectoryInfo di = new DirectoryInfo("Result");
            if (!di.Exists) di.Create();

            var linksUrl = "http://www.whynotsellreport.com/stock_rank/?page={0}";

            for (int i = 1; i < 101; i++)
            {
                var linkList = GetData(string.Format(linksUrl, i));
                var page = new BeautifulWeb.BeautifulPage(linkList);
                var tableData = page.SelectNodes("//table[@class='bordered']/tbody//tr/td/a");

                foreach (var item in tableData)
                {
                    Thread.Sleep(new Random().Next(500, 1000));

                    var pageResult = GetData($"http://www.whynotsellreport.com{item.Href}");

                    if (!string.IsNullOrEmpty(pageResult))
                    {
                        if (SaveExcel(item.Text, pageResult))
                        {
                            Console.WriteLine($"[{DateTime.Now}] {item.Text} 수집 완료!!!");
                        }
                        else
                        {
                            Console.WriteLine($"[{DateTime.Now}] {item.Text} 수집 실패!!!");
                        }
                    }
                }
            }
        }

        static string GetData(string url)
        {
            WebClient client = new WebClient();
            client.Encoding = Encoding.GetEncoding("utf-8");
            return client.DownloadString(url);
        }

        static bool SaveExcel(string name, string pageResult)
        {
            var page = new BeautifulWeb.BeautifulPage(pageResult);
            var tableData = page.SelectNodes("//table[@class='bordered']/tbody/tr");

            FileInfo excelFile = new FileInfo($"Result/{name}.xlsx");
            if (excelFile.Exists) excelFile.Delete();

            try
            {
                string[] sheets = new string[] { "데이터" };
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                using (ExcelPackage excel = new ExcelPackage())
                {
                    excel.Workbook.Worksheets.Add(sheets[0]);
                    ExcelWorksheet sheet = excel.Workbook.Worksheets[sheets[0]];

                    int itemCnt = 6; // ? 열 갯수

                    List<Object[]> dataRows = new List<Object[]>();
                    string[] dataRow = new string[itemCnt];
                    dataRow[0] = "날짜";
                    dataRow[1] = "애널리스트";
                    dataRow[2] = "제목";
                    dataRow[3] = "목표가";
                    dataRow[4] = "6개월 후";
                    dataRow[5] = "1년 후";
                    dataRows.Add(dataRow);

                    foreach (var item in tableData)
                    {
                        var td = item.HtmlNode.Elements("td");

                        int index = 0;
                        object[] _dataRow = new object[itemCnt];
                        foreach (var subItem in td)
                        {
                            _dataRow[index++] = subItem.InnerText;
                        }

                        dataRows.Add(_dataRow);
                    }

                    // 열 번호는 0부터 시작이므로 0으로 고정. 데이터 갯수에 따라 행 번호만 설정한다.
                    string headerRange = String.Format("A1:{0}1", Char.ConvertFromUtf32(itemCnt + 64));
                    sheet.Cells[headerRange].LoadFromArrays(dataRows);

                    // 각 열의 width 를 지정
                    sheet.Column(1).Width = 20;
                    sheet.Column(2).Width = 20;
                    sheet.Column(3).Width = 80;
                    sheet.Column(4).Width = 20;
                    sheet.Column(5).Width = 20;
                    sheet.Column(6).Width = 20;

                    excel.SaveAs(excelFile);
                }
            }
            catch (Exception)
            {
                return false;
            }
            
            return true;
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Policy;
using System.Text;
using System.Text.Json.Serialization;
using Jil;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace AwesomeWebsitesExcelTool
{
    class Program
    {
        static void Main(string[] args)
        {
            var excelPath = Console.ReadLine();
            if (excelPath!.StartsWith("\""))
            {
                excelPath = excelPath.Substring(1, excelPath.Length - 2);
            }

            if (!string.IsNullOrWhiteSpace(excelPath) && excelPath.EndsWith(".xlsx") && File.Exists(excelPath))
            {
                var dataList = GetData(excelPath);
                var outputJsonPath = Path.GetDirectoryName(excelPath) + @"\list.json";
                var outputMdPath = Path.GetDirectoryName(excelPath) + @"\list.md";
                OutputJson(dataList, outputJsonPath);
                OutputToMd(dataList, outputMdPath);
            }

            Console.WriteLine("Finished!");
        }

        static DataList GetData(string excelPath)
        {
            IWorkbook excel = new XSSFWorkbook(excelPath);
            ISheet sheet = excel.GetSheetAt(0);
            int rowsCount = sheet.LastRowNum;
            var list = new DataList();
            for (int i = 1; i < rowsCount; i++)
            {
                var row = sheet.GetRow(i);
                if (row != null)
                {
                    var url = row.GetCell(0).ToString()?.Trim();
                    if (string.IsNullOrWhiteSpace(url))
                    {
                        break;
                    }
                    // else
                    var tags = row.GetCell(2).ToString()
                        !.Trim()
                        .Replace(", ", ",")
                        .Split(",")
                        .ToList();
                    var urlInfo = new UrlInfo()
                    {
                        Url = url,
                        SiteName = row.GetCell(1).ToString()!.Trim(),
                        Tags = tags,
                        Introduction = row.GetCell(3).ToString()!.Trim()
                    };
                    list.AddUrlInfo(urlInfo);
                }



            }
            return list;
        }

        static void OutputJson(DataList list, string outputPath)
        {
            var json = JSON.Serialize(list);

            File.WriteAllText(outputPath, json);
        }

        static void OutputToMd(DataList list, string outputPath)
        {
            using var fs = new FileStream(outputPath, FileMode.OpenOrCreate, FileAccess.Write,FileShare.Read);
            using var sw = new StreamWriter(fs, Encoding.UTF8);

            sw.WriteLine("**Update:** " + list.Update);
            sw.WriteLine("**Websites:** ");

            sw.WriteLine(@"| url| site name | tags| brief introduction|
| --------------- | --------------- | --------------- | --------------- |");

            foreach (var urlInfo in list.Data)
            {
                sw.WriteLine($"| {urlInfo.Url} | {urlInfo.SiteName} |  {string.Join(",", urlInfo.Tags)} | {urlInfo.Introduction} |");
            }

            sw.WriteLine();
        }
    }

    class UrlInfo
    {
        [Jil.JilDirective("url")]
        public string Url { get; set; }

        [JilDirective("site name")]
        public string SiteName { get; set; }


        [JilDirective("tags")]
        public List<string> Tags { get; set; } = new List<string>();

        [JilDirective("brief introduction")]
        public string Introduction { get; set; }
    }

    class DataList
    {
        [JilDirective("update")]
        public string Update { get; set; } = DateTime.Now.ToString("yyyyMMdd");

        [JilDirective("data")] 
        public List<UrlInfo> Data { get; set; } = new List<UrlInfo>();

        public void AddUrlInfo(UrlInfo urlInfo)
        {
            this.Data.Add(urlInfo);
        }
    }
}

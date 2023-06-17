
using MiniExcelLibs;
using MiniExcelLibs.OpenXml;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.IO;
using Dapper;

namespace MiniExcelDemo
{
    class Program
    {
        static void Main(string[] args)
        {

            ExportByMySQL();


        }
        /// <summary>
        /// 通过MySQL数据库进行读取
        /// </summary>
        private static void ExportByMySQL()
        {
            string mysqlConnectionString = "Database=test;Data Source=127.0.0.1;User Id=root;Password=12345678;CharSet=utf8;port=3306";
            using (MySqlConnection conn = new MySqlConnection(mysqlConnectionString))
            {
                var reader = conn.ExecuteReader(@"select name 姓名,age 年龄,address 地址，hobby 兴趣爱好 from t_users ");
                var sheets = new Dictionary<string, object>();
                string path = "Demo" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
                sheets.Add("sheet1", reader);
                MiniExcel.SaveAs(path, sheets);
                MiniExcel.MergeSameCells("Demo1" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx", path);

            }
        }

        /// <summary>
        /// 基于模板导出
        /// </summary>
        private static void ExportTemplete()
        {
            // 1. 通过实体方式填充
            var valueByModel = new
            {
                Name = "小明",
                Age = 35,
                Address = "北京"
            };
            MiniExcel.SaveAsByTemplate("D:\\test0614实体方式.xlsx", "templete.xlsx", valueByModel);

            // 2. 通过字典方式填充
            var value = new Dictionary<string, object>()
            {
                ["Name"] = "小明",
                ["Age"] = 30,
                ["Address"] = "苏州"
            };
            MiniExcel.SaveAsByTemplate("D:\\test0614字典方式.xlsx", "templete.xlsx", value);

            // 3 通过实体list方式
            var value3 = new
            {
                users = new[] {
                    new {Name="小明",Age=30,Address="苏州"},
                    new {Name="小李",Age=25,Address="上海"},
                    new {Name="小张",Age=32,Address="南京"},
                    new {Name="小孙",Age=36,Address="常州"},
                    new {Name="小王",Age=27,Address="徐州"},
                }
            };

            MiniExcel.SaveAsByTemplate("D:\\test0614_实体list方式.xlsx", "templete2.xlsx", value3);

            // 4 通过字典list方式
            var value4 = new Dictionary<string, object>()
            {
                ["users"] = new[] {
                    new {Name="小明",Age=30,Address="苏州2"},
                    new {Name="小李",Age=25,Address="上海2"},
                    new {Name="小张",Age=32,Address="南京2"},
                    new {Name="小孙",Age=36,Address="常州2"},
                    new {Name="小王",Age=27,Address="徐州2"}
                }
            };

            MiniExcel.SaveAsByTemplate("D:\\test0614_字典list方式.xlsx", "templete2.xlsx", value4);
        }

        /// <summary>
        /// 导出包含图片
        /// </summary>
        private static void ExportAndImage()
        {
            var ss = File.ReadAllBytes("D:\\111.jpg");
            // 通过实体list方式
            var value = new[]
            {
              new {Name="小明",Age=30,Address="苏州",Image=ss},
              new {Name="小李",Age=25,Address="上海",Image=ss},
              new {Name="小张",Age=32,Address="南京",Image=ss},
              new {Name="小孙",Age=36,Address="常州",Image=ss},
              new {Name="小王",Age=27,Address="徐州",Image=ss}
            };

            // 不启用图片转换
            // var config = new OpenXmlConfiguration { EnableConvertByteArray = true };          

            MiniExcel.SaveAs("D:\\test0617_包含图片方式.xlsx", value, true, "Sheet1", ExcelType.XLSX);
        }
        /// <summary>
        /// 读取Excel
        /// </summary>
        private static void ReadExcel()
        {
            string path = "D:\\test0614_字典list方式.xlsx";
            Console.WriteLine("输出列头\n");
            var cols0 = MiniExcel.GetColumns(path);
            // 输出 A、B、C
            // useHeaderRow: true 增加参数
            var cols = MiniExcel.GetColumns(path, useHeaderRow: true);
            foreach (var item in cols)
            {
                Console.WriteLine(item);
            }
            Console.WriteLine("\n输出每一行的数据 用逗号拼接\n");
            // 读取每一行的数据 用逗号拼接
            var sheetNames = MiniExcel.GetSheetNames(path);
            foreach (IDictionary<string, object> row in MiniExcel.Query(path, useHeaderRow: true))
            {
                var str = string.Join(",", row.Values);
                Console.WriteLine(str);
            }

            Console.ReadKey();
        }

    }
}

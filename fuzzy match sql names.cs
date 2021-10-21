using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using FuzzySharp;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace FuzzyMatch
{
    class Program
    {
        static void Main(string[] args)
        {

            Stopwatch watch = new Stopwatch();
            watch.Start();

            Console.WriteLine("Getting Connection ...");

            var datasource = @"hsc-sql-2016\BITEAM";//your server
            var database = "TrakCareBI"; //your database name


            //connection string 
            string connString = @"Data Source=" + datasource + ";Initial Catalog="
                        + database + ";Persist Security Info=True;Trusted_Connection=True";

            //create instanace of database connection
            SqlConnection conn = new SqlConnection(connString);
            conn.Open();

            var namesSQL = @"
SELECT *
FROM OPENQUERY(HSSDPRD, 

'SELECT TOP 1000
         PAPMI_No as URN
       , PAPMI_Name2 as FirstName
       , PAPMI_Name as LastName
       , PAPMI_RowId->PAPER_Dob as DOB

FROM PA_PatMas

WHERE PAPMI_Name2 NOT LIKE ''zz%''
AND PAPMI_Name NOT LIKE ''zz%''


')";

            DataTable table = new DataTable();
            table.Columns.Add("URN", typeof(int));
            table.Columns.Add("FirstName", typeof(string));
            table.Columns.Add("LastName", typeof(string));
            table.Columns.Add("DOB", typeof(string));

            SqlCommand cmd = new SqlCommand(namesSQL, conn);

            using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
            {
                adapter.Fill(table);
            }
            conn.Close();

            //Console info
            watch.Stop();
            TimeSpan SqlTime = watch.Elapsed;
            Console.WriteLine($"SQL took {SqlTime.Minutes} minuites and {SqlTime.Seconds} seconds to return query");
            watch.Restart();
            Console.WriteLine("Working...");

            DataView dv = table.DefaultView;
            dv.Sort = "FirstName";
            DataTable sortedDT = dv.ToTable();

            string[] letters =
            {
                "A",
                "B",
                "C",
                "D",
                "E",
                "F",
                "G",
                "H",
                "I",
                "J",
                "K",
                "L",
                "M",
                "N",
                "O",
                "P",
                "Q",
                "R",
                "S",
                "T",
                "U",
                "V",
                "W",
                "X",
                "Y",
                "Z"
            };

            var alphaDict = new List<string>(letters);

            List<string> rowsList = new List<string>();
            string value = string.Empty;


            foreach (DataRow row in sortedDT.Rows)
            {
                value = value += string.Join(" ", row["URN"].ToString(), row["FirstName"].ToString(), row["LastName"].ToString(), row["DOB"].ToString().Replace("00:00:00", "")); 
                rowsList.Add(value);
                value = string.Empty;
            }

            DataTable final = new DataTable();
            final.Columns.Add("X", typeof(string));
            final.Columns.Add("Y", typeof(string));
            final.Columns.Add("Score", typeof(int));

            foreach (var letter in letters)
            {
                for (int i = 0; i < rowsList.Count - 1; i++)
                {
                    for (int j = i + 1; j < rowsList.Count; j++)
                    {
                        var matchResult1 = Regex.Match(rowsList[i], @"^([\w\-]+)");
                        var firstWord1 = matchResult1.Value;
                        var name1 = rowsList[i].Substring(firstWord1.Length +1);

                        var matchResult2 = Regex.Match(rowsList[i], @"^([\w\-]+)");
                        var firstWord2 = matchResult2.Value;
                        var name2 = rowsList[j].Substring(firstWord2.Length +1);

                        if (name1.Split()[1].StartsWith(letter.ToString()) && name2.Split()[1].StartsWith(letter))
                        {
                            var ratio = Fuzz.Ratio(name1, name2);

                            if (ratio < 100 && ratio > 95)
                            {
                                final.Rows.Add(rowsList[i], rowsList[j], ratio);
                                Console.WriteLine($"{name1} \t-->\t{name2} \t=\t{ratio} similarity");
                            }
                        }
                    }
                }
            }

            var xlsxFile = $@"M:\My Documents\Tests\FuzzyMatch\FuzzyResults.xlsx";

            if (File.Exists(xlsxFile))
            {
                File.Delete(xlsxFile);
            }

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            FileInfo fileInfo = new FileInfo(xlsxFile);
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {

                ExcelWorksheet ws = package.Workbook.Worksheets.Add("Fuzzies");

                ws.Cells["A1"].LoadFromDataTable(final, true);
                ws.Cells.AutoFitColumns();
                ws.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                ws.View.FreezePanes(2, 1);
                package.Save();
                package.Dispose();
                watch.Stop();
                TimeSpan C_SharpTime = watch.Elapsed;
                Console.WriteLine($"SQL took {C_SharpTime.Minutes} minuites and {C_SharpTime.Seconds} seconds to process and write the data.");
                Console.WriteLine("Finished!");

            }
        }
    }
}

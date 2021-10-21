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

            var datasource = @"hsc-sql-2016\BITEAM";
            var database = "TrakCareBI"; 

            //Connection string 
            string connString = @"Data Source=" + datasource + ";Initial Catalog="
                        + database + ";Persist Security Info=True;Trusted_Connection=True";

            //Create instanace of database connection
            SqlConnection conn = new SqlConnection(connString);
            conn.Open();

            var namesSQL = @"
SELECT *
FROM OPENQUERY(HSSDPRD, 

'SELECT TOP 5000
         PAPMI_No as URN
       , PAPMI_Name2 as FirstName
       , PAPMI_Name as LastName
       , PAPMI_RowId->PAPER_Dob as DOB
       , PAPMI_RowId->PAPER_Sex_DR->CTSEX_Desc as Gender


FROM PA_PatMas

WHERE PAPMI_Name2 NOT LIKE ''zz%''
AND PAPMI_Name NOT LIKE ''zz%''

')";

            //Create DataTable to hold SQL query data and fill
            DataTable table = new DataTable();
            table.Columns.Add("URN", typeof(int));
            table.Columns.Add("FirstName", typeof(string));
            table.Columns.Add("LastName", typeof(string));
            table.Columns.Add("DOB", typeof(string));
            table.Columns.Add("Gender", typeof(string));

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
            Console.WriteLine("Working...\n");

            //Filter by gender and split data into 2 DataTables before adding to DataSet
            DataView dvF = table.DefaultView;
            dvF.RowFilter = "Gender = 'Female'";
            DataTable femaleDT = dvF.ToTable();

            DataView dvM = table.DefaultView;
            dvF.RowFilter = "Gender = 'Male'";
            DataTable maleDT = dvF.ToTable();

            DataSet GendersDS = new DataSet();
            GendersDS.Tables.Add(femaleDT);
            GendersDS.Tables.Add(maleDT);

            GendersDS.Tables[0].TableName = "Female";
            GendersDS.Tables[1].TableName = "Male";

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

            string value = string.Empty;

            Dictionary<int, string> rowsListDict = new Dictionary<int, string>();

            foreach(DataTable genderGroup in GendersDS.Tables)

            {
                var n = 0;
                foreach (DataRow row in genderGroup.Rows)
                {
                    value = value += string.Join(" ", row["URN"].ToString(), row["FirstName"].ToString(), row["LastName"].ToString(), row["DOB"].ToString().Replace("00:00:00", ""));
                    rowsListDict.Add(n, value);
                    value = string.Empty;
                    n = n + 1;
                }

                DataTable final = new DataTable();
                final.Columns.Add("X", typeof(string));
                final.Columns.Add("Y", typeof(string));
                final.Columns.Add("Score", typeof(int));

                foreach (var letter in letters)
                {
                    for (int i = 0; i < rowsListDict.Count - 1; i++)
                    {
                        for (int j = i + 1; j < rowsListDict.Count; j++)
                        {
                            var matchResult1 = Regex.Match(rowsListDict[i], @"^([\w\-]+)");
                            var firstWord1 = matchResult1.Value;
                            var name1 = rowsListDict[i].Substring(firstWord1.Length + 1);

                            var matchResult2 = Regex.Match(rowsListDict[i], @"^([\w\-]+)");
                            var firstWord2 = matchResult2.Value;
                            var name2 = rowsListDict[j].Substring(firstWord2.Length + 1);

                            if (name1.StartsWith(letter.ToString()) && name2.StartsWith(letter))    //name1.StartsWith(letter.ToString()) && name2.StartsWith(letter)
                            {
                                var ratio = Fuzz.Ratio(name1, name2);

                                if (ratio < 100 && ratio > 90)
                                {
                                    final.Rows.Add(rowsListDict[i], rowsListDict[j], ratio);
                                    Console.WriteLine($"{name1} \t-->\t{name2} \t=\t{ratio} similarity");
                                }
                            }
                        }
                    }
                }

                var xlsxFile = $@"M:\My Documents\Tests\FuzzyMatch\{genderGroup}_FuzzyResults.xlsx";

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
                }
                rowsListDict.Clear();
            }
            watch.Stop();
            TimeSpan C_SharpTime = watch.Elapsed;
            Console.WriteLine($"SQL took {C_SharpTime.Minutes} minuites and {C_SharpTime.Seconds} seconds to process and write the data.");
            Console.WriteLine("Finished!");
        }  
    }
}


using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CsvHelper;
using FuzzySharp;

namespace FuzzyMatch
{
    class Program
    {
        static void Main(string[] args)
        {
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
FROM OPENQUERY(HSSDPRD, 'SELECT TOP 500
--PAPMI_No as UnitNumber
PAPMI_Name2 as FirstName
--, PAPMI_Name3 as MiddleName
, PAPMI_Name as Surname
, CASE WHEN PAPMI_Name LIKE ''zz % '' THEN ''Remove''
WHEN PAPMI_Name2 LIKE ''zz % '' THEN ''Remove''
ELSE NULL
END as Remove
--, PAPMI_DOB as DOB
--, PAPMI_Sex_DR->CTSEX_Desc as Gender
--, MRG_PAPMI_To_DR->PAPMI_No as MergedTo

FROM PA_PatMas
LEFT OUTER JOIN PA_MergePatient
ON MRG_PAPMI_From_DR = PAPMI_RowID

WHERE CASE
WHEN PAPMI_Name LIKE ''zz % '' THEN ''Remove''
WHEN PAPMI_Name2 LIKE ''zz % '' THEN ''Remove''
ELSE NULL
END IS NULL

ORDER BY PAPMI_Name, PAPMI_Name2

')";

            SqlCommand cmd = new SqlCommand(namesSQL, conn);

            DataTable table = new DataTable();

            using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
            {
                adapter.Fill(table);
            }
            conn.Close();




            List<string> rowsList = new List<string>();
            string value = string.Empty;



            foreach (DataRow row in table.Rows)
            {
                value = value += string.Join(" ", row.ItemArray);
                rowsList.Add(value);
                value = string.Empty;
            }



            var tupleList = new List<Tuple<string, string, int>>();
            for (int i = 0; i < rowsList.Count - 1; i++)
            {
                for (int j = i + 1; j < rowsList.Count; j++)
                {
                    var ratio = Fuzz.Ratio(rowsList[i], rowsList[j]);

                    if(ratio < 100)
                    {
                        var author = new Tuple<string, string, int>(rowsList[i], rowsList[j], ratio);
                        tupleList.Add(author);

                    }

                }
            }



            var sortedTups = tupleList.OrderByDescending(t => t.Item3).ToList();

            foreach (var t in sortedTups)
            {
                Console.WriteLine(t);
            }

            using (var writer = new StreamWriter(@"M:\My Documents\Tests\FuzzyMatch\FuzzyResults.csv"))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {

                foreach (var t in sortedTups)
                {
                    csv.WriteRecord(t);
                    csv.NextRecord();
                }
            }


        }

    }
}

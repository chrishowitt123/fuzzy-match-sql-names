using FuzzySharp;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace importDictFromFile
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
FROM OPENQUERY(HSSDPRD, '
SELECT 
--PAPMI_No as UnitNumber
PAPMI_Name2 as FirstName
--, PAPMI_Name3 as MiddleName
,PAPMI_Name as Surname
--, PAPMI_DOB as DOB
--, PAPMI_Sex_DR->CTSEX_Desc as Gender
--, MRG_PAPMI_To_DR->PAPMI_No as MergedTo

FROM PA_PatMas
LEFT OUTER JOIN PA_MergePatient
ON MRG_PAPMI_From_DR = PAPMI_RowID

WHERE PAPMI_Name NOT LIKE ''zz % ''


ORDER BY PAPMI_Name, PAPMI_Name2

')";

        SqlCommand cmd = new SqlCommand(namesSQL, conn);

        DataTable table = new DataTable();

            using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
            {
                adapter.Fill(table);
            }
            conn.Close();





            List<string> names1 = new List<string>();


            foreach (DataRow row in table.Rows)
            {
                foreach (DataColumn column in table.Columns)
                {
                    names1.Add(row[column].ToString());

                }

            }

            List<string> firstNames = new List<string>();
            for (int i = 0; i < names1.Count; i += 1)
            {
                firstNames.Add(names1[i]);
            }


            foreach (var fn in firstNames)
            {
                Console.WriteLine(fn);
                Console.WriteLine("\n");
            }


                foreach (var name in names1)
            {
                Console.WriteLine(name);

            }
            Console.WriteLine("Finidhed!");
            //names1.Add("Dave Adams");
            //names1.Add("Steve Bee");
            //names1.Add("Daves Adams");
            //names1.Add("William Keys");

            var tupleList = new List<Tuple<string, string, int>>();
            for (int i = 0; i < names1.Count - 1; i++)
            {
                for (int j = i + 1; j < names1.Count; j++)
                {
                    var ratio = Fuzz.Ratio(names1[i], names1[j]);


                    var author = new Tuple<string, string, int>(names1[i], names1[j], ratio);
                    tupleList.Add(author);
                }
            }

            var sortedTups = tupleList.OrderByDescending(t => t.Item3).ToList();

            foreach (var t in sortedTups)
            {
                Console.WriteLine(t);
            }
        }
    }
}

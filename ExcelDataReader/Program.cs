using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using ExcelDataReader.Controllers;
using System.Data.OleDb;
using System.Data;

namespace ExcelDataReader
{
    class Program
    {
        static void Main(string[] args)
        {

            var infile = FileController.GetTestDataFileName();



            //string name = "Sheet2";
            //string constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + infile + "; Extended Properties = 'Excel 12.0;HDR=YES;IMEX=1;';";
            //OleDbConnection con = new OleDbConnection(constr);
            //OleDbCommand oconn = new OleDbCommand("Select * From [" + name + "$]", con);
            //con.Open();
            //OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
            //DataTable data = new DataTable();
            //sda.Fill(data);

        }
    }
}

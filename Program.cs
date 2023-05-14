using System.Data;
using System.Data.OleDb;
using Spectre.Console;

namespace ExcellDataReader;

public class Program
{
   static void Main(string[] args)
   {
      string path = @"C:\Users\99450\Desktop\RNET102\RNET102-Tasks\RNET102\ExcelExample\ExcelExample\example_data.xlsx";
      OleDbConnection connection = new(connectionString: $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={path};Extended Properties=\"Excel 8.0; HDR = YES\";");

      connection.Open();

      string query = "SELECT * FROM [SHEET1$]";
      OleDbDataAdapter da = new OleDbDataAdapter(query, connection);


      DataTable dt = new DataTable();

      da.Fill(dt);
      connection.Dispose();

      var table = new Table();

      foreach (var item in dt.Columns)
      {
         table.AddColumn(item.ToString().Trim());
      }

      List<string> rows = new List<string>();
      foreach (DataRow item in dt.Rows)
      {
         foreach (var cell in item.ItemArray)
         {
            rows.Add(cell.ToString());
         }
         table.AddRow(rows.ToArray());
         rows.Clear();
      }

      AnsiConsole.Write(table); 

      Console.ReadLine();
   }
}


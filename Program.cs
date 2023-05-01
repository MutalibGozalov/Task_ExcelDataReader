using ExcelDataReader;
using System.Data;
using System.Text;



internal partial class Program
{
    private static void Main(string[] args)
    {
        string filePath = @"C:\Users\99450\Desktop\RNET102\Tasks\ExcellDataReader\Laptop comparsion.xlsx";
        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
        {
            Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var reader = ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration(){ FallbackEncoding = Encoding.GetEncoding("UTF-8")}))
            {
                var result = reader.AsDataSet();
                var table = result.Tables[0]; // Access the first worksheet in the file
                // Use the DataTable API to access the rows and columns
                foreach (DataRow row in table.Rows)
                {
                    // var id = row["Name"].ToString(); // Access a column by name
                    var name = row[0].ToString(); // Access a column by index
                    // var price = Convert.ToDecimal(row["Cost"]); // Convert the value to a decimal
                    Console.WriteLine(name.PadRight(10, ' ') + " | ");
                }

            }
        }


        //Console.WriteLine("Hello, World!");
    }
}
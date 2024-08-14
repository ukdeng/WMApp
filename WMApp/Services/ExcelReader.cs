using WMApp.Interfaces;
using WMApp.Models;
using ExcelDataReader;
using System.Text;


namespace WMApp.Services
{
    public class ExcelReader : IExcelReader
    {
        public List<Broker> GetBrokers()
        {
            List<Broker> brokers = new();

            var fileDirectory = $"{Directory.GetCurrentDirectory()}/wwwroot/sheet";

            if (!Directory.Exists(fileDirectory))
            {
                Directory.CreateDirectory(fileDirectory);
            }

            var filePath = Path.Combine(fileDirectory, "Advisors.xlsx");

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                var excelData = new List<List<object>>();

                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                using var reader = ExcelReaderFactory.CreateReader(stream);
                do
                {
                    while (reader.Read())
                    {
                        var rowData = new List<object>();

                        for (int column = 0; column < reader.FieldCount; column++)
                        {
                            rowData.Add(reader.GetValue(column));
                        }

                        brokers.Add(new Broker(reader.Name, rowData[0].ToString(), rowData[1].ToString()));
                    }
                } while (reader.NextResult());
            }

            return brokers;
        }
    }
}

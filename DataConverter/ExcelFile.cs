using ExcelDataReader;
using System.Data;

namespace DataConverter
{
    public class ExcelFile(string path)
    {
        private string Path { get; } = path;

        public DataSet Import()
        {
            using var s = File.OpenRead(Path);
            using var reader = ExcelReaderFactory.CreateReader(s);

            return reader.AsDataSet(new()
            {
                ConfigureDataTable = tableReader => new()
                {
                    UseHeaderRow = true,
                },
                UseColumnDataType = true,
            });
        }
    }
}

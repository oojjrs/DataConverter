using ExcelDataReader;
using System.Data;
using System.IO;

namespace DataConverter
{
    public class ExcelFile
    {
        private string _path;

        public ExcelFile(string path)
        {
            _path = path;
        }

        public DataSet Import()
        {
            using (var s = File.OpenRead(_path))
            {
                using (var reader = ExcelReaderFactory.CreateReader(s))
                {
                    return reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = tableReader => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true,
                        },
                        UseColumnDataType = true,
                    });
                }
            }
        }
    }
}

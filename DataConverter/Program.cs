using DataConverter;
using System.Data;
using System.Diagnostics;

// 최신 버전의 .Net에서 ExcelDataReader를 사용하기 위해 필요한 조치
System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

var dataRoot = args.Length > 0 ? args[0] : @"D:\ProjectA\Data";
var outputRoot = args.Length > 1 ? args[1] : @".";

var tds = new DataSet();
{
    Console.WriteLine("Get Files From : {0}", Path.GetFullPath(dataRoot));
    var files = Directory.EnumerateFiles(dataRoot, "*.xlsx").Where(file => Path.GetFileName(file).StartsWith("~") == false).Select(file => new ExcelFile(Path.GetFullPath(file))).ToArray();

    Console.WriteLine("Import {0} Files...", files.Length);
    foreach (var ds in files.Select(file => file.Import()))
        tds.Merge(ds);
}

Directory.CreateDirectory(outputRoot);

{
    Console.WriteLine("Convert to xml...");
    var sw = Stopwatch.StartNew();
    foreach (DataTable table in tds.Tables)
    {
        // 루트 이름이 이걸로 박혀서
        tds.DataSetName = "ArrayOf" + table.TableName;

        Console.WriteLine($"{table.TableName}");
        table.WriteXml(Path.Combine(outputRoot, table.TableName + ".xml"));
    }

    Console.WriteLine($"Complete. ({sw.ElapsedMilliseconds} ms)");
}

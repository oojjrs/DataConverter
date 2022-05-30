using Microsoft.CSharp;
using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml.Serialization;

namespace DataConverter
{
    public class DataConverterMain
    {
        private static void ConvertData(IEnumerable<Type> types, string root, string output)
        {
            Directory.CreateDirectory(output);

            Console.WriteLine("Get Files From : {0}", Path.GetFullPath(root));
            var files = Directory.EnumerateFiles(root, "*.xlsx").Where(file => Path.GetFileName(file).StartsWith("~") == false).Select(file => new ExcelFile(Path.GetFullPath(file))).ToArray();

            Console.WriteLine("Import {0} Files...", files.Length);
            var ds = MergeAll(files.Select(file => file.Import()));

            Console.WriteLine("Convert to C# Struct...");

            var sw = Stopwatch.StartNew();
            foreach (var type in types)
            {
                var table = ds.Tables[type.Name];
                if (table == null)
                {
                    Console.WriteLine("{0} TABLE DOES NOT EXIST.", type.Name);
                    continue;
                }

                var instances = CreateInstanceValues(ds, type, table.Select());
                using (var fs = new FileStream(Path.Combine(output, type.Name + ".xml"), FileMode.Create))
                {
                    var s = new XmlSerializer(instances.GetType());
                    s.Serialize(fs, instances);
                }
            }
            Console.WriteLine($"Complete. ({sw.ElapsedMilliseconds} ms)");
        }

        private static object CreateInstanceValues(DataSet ds, Type type, DataRow[] rows)
        {
            var props = type.GetProperties(BindingFlags.DeclaredOnly | BindingFlags.Instance | BindingFlags.Public);
            var primitives = props.Where(prop => prop.PropertyType.IsPrimitive || (prop.PropertyType == typeof(String))).ToArray();
            var enums = props.Where(prop => prop.PropertyType.IsEnum).ToArray();
            var nonPrimitives = props.Except(primitives).Except(enums).ToArray();

            var instances = new List<object>();
            foreach (DataRow row in rows)
            {
                var instance = Activator.CreateInstance(type);
                foreach (var prop in primitives)
                {
                    var value = row[prop.Name];
                    if (value == DBNull.Value)
                    {
                        if (prop.PropertyType.IsValueType)
                            value = Activator.CreateInstance(prop.PropertyType);
                    }

                    prop.SetValue(instance, Convert.ChangeType(value, prop.PropertyType));
                }

                foreach (var prop in enums)
                {
                    var value = row[prop.Name];
                    prop.SetValue(instance, Enum.Parse(prop.PropertyType, value.ToString(), true));
                }

                instances.Add(instance);
            }

            var keyProp = props.FirstOrDefault(prop => prop.Name == "Key");
            if (keyProp != null)
            {
                foreach (var instance in instances)
                {
                    foreach (var prop in nonPrimitives)
                    {
                        if (prop.PropertyType.IsArray == false)
                            throw new NotSupportedException();

                        var refName = type.Name + prop.Name;
                        var refTable = ds.Tables[refName];
                        if (refTable == null)
                        {
                            Console.WriteLine("{0} TABLE DOES NOT EXIST.", refName);
                            continue;
                        }

                        var keyValue = keyProp.GetValue(instance).ToString();
                        var thisRows = refTable.Select(string.Format("{0} = '{1}'", keyProp.Name, keyValue));
                        var values = CreateSmartValues(ds, prop.PropertyType.GetElementType(), thisRows);
                        prop.SetValue(instance, values);
                    }
                }
            }

            var container = Array.CreateInstance(type, instances.Count);
            for (int i = 0; i < instances.Count; ++i)
                container.SetValue(instances[i], i);

            return container;
        }

        private static object CreatePrimitiveValues(Type type, DataRow[] rows)
        {
            var instances = new List<object>();
            foreach (DataRow row in rows)
            {
                var value = row[type.Name];
                if (value == DBNull.Value)
                {
                    if (type.IsValueType)
                        value = Activator.CreateInstance(type);
                }

                if (type.IsEnum)
                    instances.Add(Enum.Parse(type, value.ToString(), true));
                else
                    instances.Add(Convert.ChangeType(value, type));
            }

            var container = Array.CreateInstance(type, instances.Count);
            for (int i = 0; i < instances.Count; ++i)
                container.SetValue(instances[i], i);

            return container;
        }

        private static object CreateSmartValues(DataSet ds, Type type, DataRow[] rows)
        {
            if (type.IsPrimitive || type.IsEnum || (type == typeof(String)))
                return CreatePrimitiveValues(type, rows);
            else
                return CreateInstanceValues(ds, type, rows);
        }

        private static Type[] GetAllDataTypes(string root)
        {
            var sw = Stopwatch.StartNew();
            var provider = new CSharpCodeProvider();
            var files = Directory.EnumerateFiles(root, "*.cs", SearchOption.TopDirectoryOnly).Where(t => t.EndsWith("Data.cs") || t.EndsWith("Type.cs")).ToArray();
            var ret = provider.CompileAssemblyFromFile(new CompilerParameters()
            {
                GenerateInMemory = true,
            }, files);

            if (ret.Errors.HasErrors)
            {
                foreach (CompilerError error in ret.Errors)
                    Console.WriteLine(error);

                return Array.Empty<Type>();
            }

            var types = ret.CompiledAssembly.GetTypes().Where(t => t.Name.EndsWith("Data")).ToArray();

            Console.WriteLine($"Types : {types.Length} ({sw.ElapsedMilliseconds} ms)");
            return types;
        }

        private static void Main(string[] args)
        {
            var typeRoot = args.Length > 0 ? args[0] : @"D:\Meister2\Meister2\Assets\Sources\Scripts\Entity";
            var dataRoot = args.Length > 1 ? args[1] : @"D:\Meister2\Meister2\Data";
            var outputRoot = args.Length > 2 ? args[2] : @"D:\Meister2\Meister2\Assets\Sources\Data";

            ConvertData(GetAllDataTypes(typeRoot), dataRoot, outputRoot);

            Console.Write("Press any key to continue...");
            Console.ReadKey();
        }

        private static DataSet MergeAll(IEnumerable<DataSet> dss)
        {
            var all = new DataSet();
            foreach (var ds in dss)
                all.Merge(ds);
            return all;
        }
    }
}

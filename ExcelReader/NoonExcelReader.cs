using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace ExcelReader
{
    public static class NoonExcelReader
    {
        private static Stream GetStream(string path)
        {
            if (!File.Exists(path)) throw new FileNotFoundException("File not found.", path, null);
            if (Path.GetExtension(path) != ".xlsx") throw new InvalidCastException(); 

            var _tempPath = string.Format("{0}/{1}.{2}", Path.GetTempPath(), new Guid(), ".tmp");

            File.Copy(path, _tempPath, true);
            return File.Open(_tempPath, FileMode.Open);
        }

        public static List<List<string>> Read(string path)
        {
            return Read(GetStream(path));
        }

        private static List<List<string>> Read(Stream fileStream)
        {
            IExcelDataReader reader = ExcelReaderFactory.CreateReader(fileStream);
            var dset = reader.AsDataSet();

            var list = new List<List<string>>();
            list.Add(new List<string>());
            var table = dset.Tables[0];

            var _row = 0;
            foreach (DataRow row in table.Rows)
            {
                var count = 0;
                foreach (object col in row.ItemArray)
                {
                    if (col == DBNull.Value) continue;

                    if (!string.IsNullOrEmpty(col.ToString()))
                    {
                        list[_row].Add(col.ToString());
                        count++;
                    }
                }

                if (count > 0)
                {
                    list.Add(new List<string>());
                    _row++;
                }
            }

            list.Remove(list.Last());

            return list;
        }

        private static List<T> Convert<T>(List<List<string>> matrix)
        {
            var titles = matrix.First();
            matrix.RemoveAt(0);
            List<T> list = new List<T>();

            foreach (var row in matrix)
            {
                var _ob = (T)typeof(T).GetConstructor(Type.GetTypeArray(new object[0])).Invoke(null);

                foreach (var prop in typeof(T).GetProperties())
                {
                    if (!titles.Contains(prop.Name)) continue;
                    var index = titles.IndexOf(prop.Name);
                    prop.SetValue(_ob, row[index]);
                }

                list.Add(_ob);
            }

            return list.ConvertAll<T>(o => (T)o);
        }

        public static List<T> ReadConvert<T>(string path)
        {
            return Convert<T>(Read(GetStream(path)));
        }

        private static List<T> ReadConvert<T>(Stream fileStream)
        {
            return Convert<T>(Read(fileStream));
        }

        private static bool CheckValid(Stream fileStream, Type modelType)
        {
            var props = modelType.GetProperties().ToList().ConvertAll<string>(t => t.Name);
            var matrix = Read(fileStream);

            if (matrix.Count > 0)
            {
                bool isValid = true;
                matrix[0].ForEach(col =>
                {
                    isValid &= props.Contains(col);
                });

                return isValid;
            }

            return false;
        }
    }
}

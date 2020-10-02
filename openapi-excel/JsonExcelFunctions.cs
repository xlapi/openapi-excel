using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Linq;

namespace openapi_excel
{
    public class JsonExcelFunctions
    {
        [ExcelFunction(Description = "JSON Path")]
        public static string JSONPATH(string contents, string path)
        {
            return JObject.Parse(contents).SelectToken(path).ToString();
        }

        [ExcelFunction(Description = "JSON Path")]
        public static string JsonListToRows(string contents, string path)
        {
            var items = JObject.Parse(contents).SelectTokens(path).Select(x => x.ToString());

            ExcelReference origin = (ExcelReference)XlCall.Excel(XlCall.xlfCaller);

            int i = origin.RowLast;
            foreach (var item in items)
            {
                ExcelReference r = new ExcelReference(i, origin.ColumnLast);
                r.SetValue(item);
                i++;
            }
            return "Ok";
        }

        [ExcelFunction(Description = "JSON Path")]
        public static string JsonListToRowsEasy(string contents)
        {
            var items = JArray.Parse(contents).Children();

            ExcelReference origin = (ExcelReference)XlCall.Excel(XlCall.xlfCaller);

            int i = origin.RowLast + 1;

            string firstCellResult = null;
            if (items.Any())
            {
                firstCellResult = items.First().ToString();
            }
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                foreach (var item in items.Skip(1))
                {
                    ExcelReference r = new ExcelReference(i, origin.ColumnLast);
                    r.SetValue(item.ToString());
                    i++;
                }
            });
            
            return firstCellResult;
        }

        [ExcelFunction(Description = "JSON Table")]
        public static string JsonListOfObjectsToTable(string contents)
        {
            var items = JArray.Parse(contents).Children();

            ExcelReference origin = (ExcelReference)XlCall.Excel(XlCall.xlfCaller);

            string firstCellResult = null;

            List<string> titles = new List<string>();

            if (items.Any())
            {
                var obj = items.First() as JObject;
                if (obj != null)
                {
                    titles = obj.Properties().Select(p => p.Name).ToList();
                    firstCellResult = titles.FirstOrDefault();
                }
            }
            int rowCursor = origin.RowLast;
            var originColumn = origin.ColumnLast;

            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                // Write titles
                var titleCursor = originColumn + 1;
                foreach (var title in titles.Skip(1))
                {
                    ExcelReference r = new ExcelReference(rowCursor, titleCursor);
                    r.SetValue(title);
                    titleCursor++;
                }

                rowCursor++;
                foreach (var item in items)
                {
                    var columnCursor = originColumn;
                    foreach (var property in titles)
                    {
                        var obj = item as JObject;
                        if (obj != null)
                        {
                            var propertyOfItem = obj.Properties().SingleOrDefault(p => p.Name == property);
                            if (propertyOfItem != null)
                            {
                                ExcelReference r = new ExcelReference(rowCursor, columnCursor);
                                r.SetValue(propertyOfItem.Value.ToString());
                                columnCursor++;
                            }
                            
                        }
                        
                    }
                    rowCursor++;
                }
            });

            return firstCellResult;
        }
    }
}

using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using Microsoft.OpenApi.Models;
using System.Linq;
using System.Windows.Forms;
using ComApp = Microsoft.Office.Interop.Excel.Application;

namespace openapi_excel
{
    class ExcelSheetWriter
    {
        public static void WriteDocumentationSheet()
        {
            var XlApp = ExcelDnaUtil.Application as ComApp;
            if (XlApp.ActiveWorkbook == null)
            {
                XlApp.Workbooks.Add();
            }

            var sheet = (Worksheet)XlApp.Sheets.Add();

            sheet.Name = "API Documentation";

            sheet.Range["A1"].Value = "Api Documentation";

            var currentRow = 2;

            foreach (var path in SwaggerRegistry.Api.Paths)
            {
                currentRow++;
                sheet.Range[$"B{currentRow}"].Value = "Url";
                sheet.Range[$"C{currentRow}"].Value = "Description";
                currentRow++;
                sheet.Range[$"B{currentRow}"].Value = path.Key;
                sheet.Range[$"C{currentRow}"].Value = path.Value.Description;

                currentRow++;
                currentRow++;

                foreach (var operation in path.Value.Operations)
                {
                    sheet.Range[$"C{currentRow}"].Value = "HTTP Method";
                    sheet.Range[$"D{currentRow}"].Value = "Operation Id";
                    sheet.Range[$"E{currentRow}"].Value = "Description";
                    sheet.Range[$"F{currentRow}"].Value = "Example";
                    currentRow++;

                    sheet.Range[$"C{currentRow}"].Value = operation.Key.ToString();
                    sheet.Range[$"D{currentRow}"].Value = operation.Value.OperationId;
                    sheet.Range[$"E{currentRow}"].Value = operation.Value.Description;

                    var paramsString = string.Join(",", operation.Value.Parameters.Select(p => $"{p.Name}Value"));
                    sheet.Range[$"F{currentRow}"].NumberFormat = "@";
                    sheet.Range[$"F{currentRow}"].Value = $"={operation.Value.OperationId}({paramsString})";

                    currentRow++;
                }
                currentRow++;
            }
        }

        internal static void AddOperationToSheet(OpenApiOperation operation)
        {
            var XlApp = ExcelDnaUtil.Application as ComApp;
            if (XlApp.ActiveWorkbook == null)
            {
                MessageBox.Show($"Please open a worksheet first");
                return;
            }

            if (XlApp.ActiveCell == null)
            {
                MessageBox.Show($"Please select a cell first");
                return;
            }

            var cell = XlApp.ActiveCell;
            var paramsString = string.Join(",", operation.Parameters.Select(p => $"{p.Name}Value"));
            cell.Value = $"={operation.OperationId}({paramsString})";
        }
    }
}

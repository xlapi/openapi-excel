using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using Microsoft.OpenApi.Models;
using Newtonsoft.Json.Linq;
using openapi_excel.DI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace openapi_excel
{
    class FunctionRegistrar
    {
        public static RegistrationResults RegisterApi(OpenApiDocument apiDefinition, bool reregister = false)
        {
            List<Delegate> delegates = new List<Delegate>();
            List<object> funcAttribs = new List<object>();
            List<List<object>> argAttribsList = new List<List<object>>();

            var functionsAdded = new List<string>();

            foreach (var path in apiDefinition.Paths)
            {
                foreach (var operation in path.Value.Operations)
                {
                    delegates.Add(CreateDelegateForOperation(path.Key, path.Value, operation.Key, operation.Value));

                    ExcelFunctionAttribute att = new ExcelFunctionAttribute();

                    att.Name = operation.Value.OperationId;

                    att.Description = operation.Value.Description;
                    att.HelpTopic = apiDefinition.ExternalDocs?.Url?.ToString();
                    att.SuppressOverwriteError = reregister;

                    funcAttribs.Add(att);
                    List<object> argAttribs = new List<object>();

                    foreach (var parameter in operation.Value.Parameters)
                    {
                        ExcelArgumentAttribute atta1 = new ExcelArgumentAttribute();
                        atta1.Name = parameter.Name;
                        atta1.Description = parameter.Description;

                        argAttribs.Add(atta1);
                    }
                    
                    argAttribsList.Add(argAttribs);

                    functionsAdded.Add(att.Name);
                }
            }

            ExcelIntegration.RegisterDelegates(delegates, funcAttribs, argAttribsList);

            var registrationResults = new RegistrationResults
            {
                FunctionsAdded = functionsAdded
            };

            return registrationResults;
        }

        public static void DeregisterFunctions(RegistrationResults registeredFunctions)
        {
            foreach (var function in registeredFunctions.FunctionsAdded)
            {
                // get the path to the XLL 
                var xllName = XlCall.Excel(XlCall.xlGetName);

                // get the registered ID for this function and unregister 
                var regId = XlCall.Excel(XlCall.xlfEvaluate, function);
                XlCall.Excel(XlCall.xlfSetName, function);
                XlCall.Excel(XlCall.xlfUnregister, regId);

                //var reregId = XlCall.Excel(XlCall.xlfRegister, xllName, "xlAutoRemove", "I", function, ExcelMissing.Value, 2);
                //XlCall.Excel(XlCall.xlfSetName, function);
                //XlCall.Excel(XlCall.xlfUnregister, reregId);
            }
        }

        private static Delegate CreateDelegateForOperation(string apiPath, OpenApiPathItem pathItem, OperationType operationType, OpenApiOperation operation)
        {
            var dict = Expression.Variable(typeof(Dictionary<string, object>));
            ParameterExpression result = Expression.Parameter(typeof(object), "result");

            var dictAdd = typeof(Dictionary<string, object>).GetMethod("Add", BindingFlags.Public | BindingFlags.Instance, null, new[] { typeof(string), typeof(object) }, null);

            var dictNew = Expression.Assign(dict, Expression.New(typeof(Dictionary<string, object>)));

            var expressions = new List<Expression>();
            
            var inputs = new List<ParameterExpression>();
            expressions.Add(dictNew);

            var operationPathInput = Expression.Constant(apiPath);
            var operationInput = Expression.Constant(operation);

            foreach (var key in operation.Parameters)
            {
                ParameterExpression apiParam = Expression.Parameter(typeof(object), "object" + key.Name);
                inputs.Add(apiParam);
                expressions.Add(apiParam);

                var keyE = Expression.Constant(key.Name);
                expressions.Add(Expression.Call(dict, dictAdd, keyE, apiParam));
            }

            // Add a params thing at the end
            ParameterExpression finalParam = Expression.Parameter(typeof(string), "stringParams");
            inputs.Add(finalParam);
            expressions.Add(finalParam);

            expressions.Add(Expression.Assign(
                    result, Expression.Call(
                    null,
                    typeof(ApiCaller).GetMethod("CallApi", new Type[] { typeof(string), typeof(OpenApiOperation), typeof(Dictionary<string, object>), typeof(string) }),
                    operationPathInput,
                    operationInput,
                    dict,
                    finalParam
                   ))
                );

            var block = Expression.Block(new[] { dict, result }, expressions);
            var del = Expression.Lambda(block, inputs.ToArray());
            return del.Compile();
        }
    }

    public static class ApiCaller
    {
        public static object CallApi(string path, OpenApiOperation operation, Dictionary<string, object> paramsArgs, string stringParams)
        {
            ExcelReference origin = (ExcelReference)XlCall.Excel(XlCall.xlfCaller);

            var options = new List<string>();
            if (!string.IsNullOrWhiteSpace(stringParams))
            {
                options = stringParams.Split(',').ToList();
            }

            var paramsDict = string.Join("", paramsArgs.Select(x => $"{x.Key}, {x.Value}"));

            return ExcelAsyncUtil.Run("CallApi", new[] { path, operation.OperationId, paramsDict }, () =>
            {
                var client = Resolver.Instance.Create<SwaggerClient>();
                var result = client.Call(path, operation.OperationId, paramsArgs).ConfigureAwait(false).GetAwaiter().GetResult();

                if (options.Contains("raw"))
                {
                    return result;
                }

                if (options.Any(x => x.StartsWith("path")))
                {
                    var pathString = options.First(x => x.StartsWith("path"));
                    if (pathString.Contains("="))
                    {
                        var jsonPath = pathString.Split('=')[1];
                        var match = JToken.Parse(result).SelectToken(jsonPath);
                        if (match != null)
                        {
                            return match.ToString();
                        }
                        else
                        {
                            return $"No match for {jsonPath}";
                        }
                    }
                }

                // If result is list, try to expand out fields

                var jsonToken = JToken.Parse(result);

                if (jsonToken is JArray)
                {
                    var thing = (JArray)jsonToken;
                    var items = thing.Children();

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

                    if (options.Any(o => o.StartsWith("fields")))
                    {
                        var fieldsString = options.First(x => x.StartsWith("fields"));
                        var fieldsToInclude = fieldsString.Split('=')[1].Split(';');

                        titles = titles.Where(t => fieldsToInclude.Contains(t)).ToList();
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
                        Application xlApp = (Application)ExcelDnaUtil.Application;

                        Workbook wb = xlApp.ActiveWorkbook;
                        Worksheet ws = wb.ActiveSheet;
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
                                    WriteValue(rowCursor, columnCursor, propertyOfItem);
                                    columnCursor++;
                                }

                            }
                            rowCursor++;
                        }
                    });

                    return firstCellResult;
                }
                if (jsonToken is JObject)
                {
                    var item = (JObject)jsonToken;
                    var properties = item.DeserializeAndFlatten();

                    string firstCellResult = null;
                    if (properties.Any())
                    {
                        firstCellResult = properties.First().Key;
                    }

                    var titleColumn = origin.ColumnLast;
                    var valuesColumn = titleColumn + 1;
                    var i = origin.RowLast;

                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        foreach (var pair in properties)
                        {
                            // Skip first
                            if (i != origin.RowLast)
                            {
                                ExcelReference r = new ExcelReference(i, titleColumn);
                                r.SetValue(pair.Key.ToString());
                            }

                            WriteValue(i, valuesColumn, pair.Value);
                            i++;
                        }
                    });

                    return firstCellResult;
                }
                return result;
            });
        }

        private static void WriteValue(int rowCursor, int columnCursor, JProperty propertyOfItem)
        {
            if (propertyOfItem != null)
            {
                if (propertyOfItem.HasValues)
                {
                    ExcelReference r = new ExcelReference(rowCursor, columnCursor);

                    var theValue = propertyOfItem.Value.ToString();

                    if (long.TryParse(theValue, out var valueAsNumber))
                    {
                        r.SetValue(valueAsNumber);
                    }
                    else
                    {
                        r.SetValue(propertyOfItem.Value.ToString());
                    }
                }
            }
        }

        private static void WriteValue(int rowCursor, int columnCursor, object propertyOfItem)
        {
            if (propertyOfItem != null)
            {
                ExcelReference r = new ExcelReference(rowCursor, columnCursor);

                var theValue = propertyOfItem.ToString();

                if (long.TryParse(theValue, out var valueAsNumber))
                {
                    r.SetValue(valueAsNumber);
                }
                else
                {
                    r.SetValue(propertyOfItem.ToString());
                }
            }
        }
    }
}

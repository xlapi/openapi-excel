using Microsoft.OpenApi.Models;
using Microsoft.OpenApi.Readers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace openapi_excel
{
    public class SwaggerClient
    {
        public async Task<OpenApiDocument> GetApiDefinition(string url)
        {
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            var httpClient = new HttpClient();

            var stream = await httpClient.GetStreamAsync(url);

            // Read V3 as YAML
            var openApiDocument = new OpenApiStreamReader().Read(stream, out var diagnostic);
            
            if (diagnostic.Errors.Any())
            {
                throw new ArgumentException($"The provided API has problems and cannot be used. Errors: {string.Join(", ", diagnostic.Errors.Select(e => e.Message))}");
            }

            return openApiDocument;
        }

        internal async Task<string> Call(string path, string operationName, Dictionary<string, object> paramsArgs)
        {
            var api = SwaggerRegistry.Api;

            string baseUrl;
            if (!api.Servers.Any())
            {
                baseUrl = new Uri(SwaggerRegistry.Url).GetLeftPart(UriPartial.Authority);
            }
            else
            {
                baseUrl = api.Servers.First().Url;
            }

            var methodAndOperation = SwaggerRegistry.Api.Paths.Single(x => x.Key == path).Value.Operations.Single(y => y.Value.OperationId == operationName);

            var operation = methodAndOperation.Value;

            var httpClient = new HttpClient();

            Uri baseUri = new Uri(baseUrl + path);

            var httpRequest = new HttpRequestMessage();

            var securityToAdd = new List<SecurityKeyValue>();

            if (methodAndOperation.Value.Security.Any())
            {
                // If basic
                if (methodAndOperation.Value.Security.ToList().SelectMany(x => x.Keys).Any(y => y.Type == SecuritySchemeType.Http && y.Scheme == "basic"))
                {
                    if (SwaggerRegistry.BasicAuthCreds == null)
                    {
                        SecurityManager.Login();
                    }

                    httpRequest.Headers.Authorization =
                        new AuthenticationHeaderValue(
                            "Basic", Convert.ToBase64String(
                                System.Text.ASCIIEncoding.ASCII.GetBytes(
                                    $"{SwaggerRegistry.BasicAuthCreds.Username}:{SwaggerRegistry.BasicAuthCreds.Password}")));
                }

                if (methodAndOperation.Value.Security.SelectMany(x => x.Select(y => y.Key.Type == SecuritySchemeType.ApiKey)).Any())
                {
                    foreach (var neededKey in methodAndOperation.Value.Security.SelectMany(x => x.Where(y => y.Key.Type == SecuritySchemeType.ApiKey)))
                    {
                        // try and get it
                        if (SwaggerRegistry.ApiKeyCredentials.TryGetValue(neededKey.Key.Name, out var apiKey)) {
                            httpRequest.Headers.Add(apiKey.Key, apiKey.Value);
                        }
                        else
                        {
                            throw new ArgumentException("Need to setup keys");
                        }
                        
                    }
                }
            }

            httpRequest.Method = GetHttpMethod(methodAndOperation.Key);

            foreach (var queryParameter in operation.Parameters.Where(x => x.In == ParameterLocation.Query))
            {
                baseUri = baseUri.AddQuery(queryParameter.Name, paramsArgs[queryParameter.Name].ToString());
            }

            var stringUrl = baseUri.ToString();

            foreach (var pathParameter in operation.Parameters.Where(x => x.In == ParameterLocation.Path))
            {
                stringUrl = stringUrl.Replace("{" + pathParameter.Name + "}", paramsArgs[pathParameter.Name].ToString());
            }

            foreach (var headerParameter in operation.Parameters.Where(x => x.In == ParameterLocation.Header))
            {
                httpRequest.Headers.Add(headerParameter.Name, paramsArgs[headerParameter.Name].ToString());
            }

            foreach (var securityThing in securityToAdd)
            {
                httpRequest.Headers.Add(securityThing.Key, securityThing.Value);
            }

            httpRequest.RequestUri = new Uri(stringUrl);

            var response = await httpClient.SendAsync(httpRequest).ConfigureAwait(false);
            var result = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            return result;
        }

        private HttpMethod GetHttpMethod(OperationType key)
        {
            switch (key)
            {
                case OperationType.Get:
                    return HttpMethod.Get;
                case OperationType.Post:
                    return HttpMethod.Post;
                case OperationType.Put:
                    return HttpMethod.Put;
                case OperationType.Delete:
                    return HttpMethod.Delete;
                case OperationType.Head:
                    return HttpMethod.Head;
                case OperationType.Options:
                    return HttpMethod.Options;
                case OperationType.Trace:
                    return HttpMethod.Trace;
                default:
                    return HttpMethod.Get;
            }
        }

        public async Task<ValidationResult> Validate(string url)
        {
            var httpClient = new HttpClient();

            try
            {
                Uri baseUri = new Uri(url);
                await httpClient.GetAsync(baseUri);
                return new ValidationResult { IsOk = true };
            }
            catch (Exception e)
            {
                return new ValidationResult { IsOk = false, Error = e.Message };
            }
        }
    }

    public class ValidationResult
    {
        public bool IsOk { get; set; }
        public string Error { get; set; }
    }
}

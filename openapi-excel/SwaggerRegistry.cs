using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using ExcelDna.Integration;
using openapi_excel.Security;
using openapi_excel.UI.Ribbon;
using Microsoft.OpenApi.Models;
using System.Windows.Forms;

namespace openapi_excel
{
    public class SwaggerRegistry
    {
        // Unfortunately, the ribbon is hard to DI, so give it a Registry it can call
        public static SwaggerRegistry Instance;

        private readonly SwaggerClient _client;

        public static OpenApiDocument Api { get; private set; }
        public static string Url { get; private set; }
        public static RegistrationResults RegisteredFunctions { get; private set; }
        public static BasicAuthCredentials BasicAuthCreds { get; private set; } = null;

        public static ApiLoadStatusContainer ApiLoadStatusStatic = new ApiLoadStatusContainer();

        private SynchronizationContext _syncContext;

        public static Dictionary<string, ApiKey> ApiKeyCredentials { get; set; } = new Dictionary<string, ApiKey>();

        public SwaggerRegistry(SwaggerClient client) {
            _client = client;
            Instance = this;
        }

        public void Register(string url, bool reregister = false)
        {
            _syncContext = SynchronizationContext.Current ?? new WindowsFormsSynchronizationContext();

            Task.Factory.StartNew(() => Register(url));
        }

        public async Task Register(string url)
        {
            try
            {
                var sc = new SwaggerClient();
                var apiDefinition = await sc.GetApiDefinition(url);

                _syncContext.Post(delegate (object state)
                {
                    Api = apiDefinition;
                    Url = url;
                    AfterLoaded(url, apiDefinition);
                }, null);

            }
            catch (Exception e)
            {
                if (e.InnerException != null)
                {
                    var httpException = e.InnerException as WebException;
                    if (httpException != null && httpException.Status == WebExceptionStatus.ConnectFailure)
                    {
                        ApiLoadStatusStatic.Status = ApiLoadStatus.ConnectionFailure;
                        _syncContext.Post(delegate (object state)
                        {
                            AfterLoadedError(url);
                        }, null);
                        return;
                    }
                }
                RibbonController.InvalidateRibbon();
            }
        }

        private void AfterLoaded(string url, OpenApiDocument apiDefinition)
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                RegisteredFunctions = FunctionRegistrar.RegisterApi(apiDefinition, false);
                ApiKeyCredentials = Api.Components.SecuritySchemes.Values.Where(ss => ss.Type == SecuritySchemeType.ApiKey).ToDictionary(s => s.Name, s => new ApiKey { Key = s.Name, Value = "", In = s.In });
            });
            ApiLoadStatusStatic.Status = ApiLoadStatus.Loaded;
            RibbonController.InvalidateRibbon();
        }

        private void AfterLoadedError(string url)
        {
            ApiLoadStatusStatic.Status = ApiLoadStatus.ConnectionFailure;
            RibbonController.InvalidateRibbon();
        }

        public void Remove()
        {
            Api = null;

            if (RegisteredFunctions != null)
            {
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    FunctionRegistrar.DeregisterFunctions(RegisteredFunctions);
                });
            }
        }

        internal static void SaveCredentials(BasicAuthCredentials creds)
        {
            BasicAuthCreds = creds;
        }

        internal void Refresh()
        {
            Register(Url, true);
        }
    }
}

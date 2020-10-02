using ExcelDna.Integration;
using ExcelDna.Registration;
using openapi_excel.DI;
using System.Threading;
using System.Windows.Forms;

namespace openapi_excel
{
    public class Startup : IExcelAddIn
    {
        static SynchronizationContext _syncContext;

        public void AutoOpen()
        {
            SetupExcel();
            RegisterFunctions();
        }

        private void SetupExcel()
        {
            ExcelIntegration.RegisterUnhandledExceptionHandler(
                delegate (object ex) {
                    return string.Format("Error: {0}", ex.ToString());
                }
            );
        }

        public void AutoClose()
        {
        }

        public void RegisterFunctions()
        {
            // DI as best we can do it
            var instance = Resolver.Instance;
            instance.Register<SwaggerClient>((c) => new SwaggerClient());
            instance.Register<SwaggerRegistry>((c) => new SwaggerRegistry(c.Create<SwaggerClient>()));

            ExcelRegistration.GetExcelFunctions()
                             .RegisterFunctions();

            _syncContext = SynchronizationContext.Current ?? new WindowsFormsSynchronizationContext();

            SwaggerRegistry.ApiLoadStatusStatic.Status = ApiLoadStatus.Loading;

            var registry = instance.Create<SwaggerRegistry>();
            registry.Register();
        }
    }
}

using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Xml.Linq;

namespace openapi_excel.UI.Ribbon
{
    [ComVisible(true)]
    public class RibbonController : ExcelRibbon
    {
        private static IRibbonUI _thisRibbon;

        public override string GetCustomUI(string RibbonID)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceName = "openapi_excel.UI.Ribbon.Ribbon.xml";

            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            using (StreamReader reader = new StreamReader(stream))
            {
                string result = reader.ReadToEnd();
                var control = XDocument.Parse(result);

                //var statusLabel = control.Descendants().Single(x => x.Attribute("id")?.Value == "apiConnectionStatus");

                

                return control.ToString();
            }
        }

        public string getStatusLabel(IRibbonControl control)
        {
            if (SwaggerRegistry.ApiLoadStatusStatic?.Status == null)
            {
                return "Initialising";
            }
            switch (SwaggerRegistry.ApiLoadStatusStatic.Status)
            {
                case ApiLoadStatus.ConnectionFailure:
                    return "Connection Failure";
                case ApiLoadStatus.Loaded:
                    return "Connected";
                case ApiLoadStatus.Loading:
                    return "Loading";
                case ApiLoadStatus.UnknownFailure:
                    return "Unknown failure";
                case ApiLoadStatus.NotSet:
                    return "NotSet";
            }
            return "Error";
        }

        public void OnRibbonLoad(IRibbonUI ribbon)
        {
            if (ribbon == null)
            {
                throw new ArgumentNullException(nameof(ribbon));
            }

            _thisRibbon = ribbon;
        }

        public static void InvalidateRibbon()
        {
            if (_thisRibbon != null)
            {
                _thisRibbon.Invalidate();
                _thisRibbon.InvalidateControl("openApiTab");
            }
        }

        public void onIV(IRibbonControl control)
        {
            InvalidateRibbon();
        }

        public void onShowHelp(IRibbonControl control)
        {
            SideBarController.ShowCTP();
        }

        public void onShowLogin(IRibbonControl control)
        {
            SecurityManager.Relogin();
        }

        public void onShowDocumentation(IRibbonControl control)
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                ExcelSheetWriter.WriteDocumentationSheet();
            });
        }

        public void onShowSecurity(IRibbonControl control)
        {
            SecurityManager.ShowApiSecurityForm();
        }
    }
}

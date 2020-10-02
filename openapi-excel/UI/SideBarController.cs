using ExcelDna.Integration.CustomUI;

namespace openapi_excel.UI
{
    public static class SideBarController
    {
        static CustomTaskPane ctp;

        public static void ShowCTP()
        {
            if (ctp == null)
            {
                // Make a new one using ExcelDna.Integration.CustomUI.CustomTaskPaneFactory 
                ctp = CustomTaskPaneFactory.CreateCustomTaskPane(typeof(HelpSideBarControl), SwaggerRegistry.Api.Info.Title);
                ctp.Visible = true;
                ctp.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
            }
            else
            {
                // Just show it again
                ctp.Visible = true;
            }
        }

        public static void DeleteCTP()
        {
            if (ctp != null)
            {
                // Could hide instead, by calling ctp.Visible = false;
                ctp.Delete();
                ctp = null;
            }
        }
    }
}

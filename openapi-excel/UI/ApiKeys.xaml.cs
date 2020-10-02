using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace openapi_excel.UI
{
    /// <summary>
    /// Interaction logic for ApiKeys.xaml
    /// </summary>
    public partial class ApiKeys : UserControl
    {
        public ApiKeys()
        {
            InitializeComponent();

            ApiConfigGrid.ItemsSource = apiThings;
        }

        public List<ApiKey> apiThings = new List<ApiKey>();

        internal void SetApiKeys(List<ApiKey> possibleApiKeys)
        {
            apiThings.AddRange(possibleApiKeys);
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var myWindow = Window.GetWindow(this);
            myWindow.Close();
        }
    }
}

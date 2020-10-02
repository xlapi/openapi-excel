using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;

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

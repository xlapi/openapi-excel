using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace openapi_excel.UI
{
    public partial class AddApiDialog : Form
    {
        public bool WasCancelled { get; set; } = false;
        public string Url { get; set; }

        public AddApiDialog()
        {
            InitializeComponent();
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            WasCancelled = true;
        }

        private async void OnAcceptButton_Click(object sender, EventArgs e)
        {
            var sc = new SwaggerClient();

            ((Button)sender).Text = "working...";
            var validationResult = await sc.Validate(urlTextBox.Text).ConfigureAwait(true);

            if (validationResult.IsOk)
            {
                Url = urlTextBox.Text;
                DialogResult = DialogResult.OK;
                return;
            }
            else
            {
                MessageBox.Show(this, $"Api URL could not be loaded: {validationResult.Error}", "Api Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ((Button)sender).Text = "Ok";
                DialogResult = DialogResult.None;
            }
        }
    }
}

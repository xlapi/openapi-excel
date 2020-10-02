using System;
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
            this.Close();
        }

        private void OnAcceptButton_Click(object sender, EventArgs e)
        {
            Url = this.urlTextBox.Text;
            DialogResult = DialogResult.OK;
            return;
        }
    }
}

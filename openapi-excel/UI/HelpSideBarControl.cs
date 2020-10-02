using System;
using System.Runtime.InteropServices;
using Microsoft.OpenApi.Models;
using Label = System.Windows.Forms.Label;
using Button = System.Windows.Forms.Button;
using System.Windows.Forms;
using System.Drawing;

namespace openapi_excel.UI
{
    [ComVisible(true)]
    public class HelpSideBarControl : UserControl
    {
        public HelpSideBarControl()
        {
            this.SizeChanged += onSizeChanged;
            Draw();
        }

        private const int PADDING = 5;
        private const int LABEL_HEIGHT = 30;
        private const int TOTAL_LABEL_HEIGHT = PADDING + LABEL_HEIGHT;

        private void Draw()
        {
            Controls.Clear();
            var width = this.Width;
            BackColor = Color.FromArgb(243, 242, 241);

            var descriptionLabel = new Label();
            descriptionLabel.Text = SwaggerRegistry.Api.Info.Description;
            descriptionLabel.Location = new Point(8, 10);
            descriptionLabel.Size = new Size(width - 16, TOTAL_LABEL_HEIGHT);
            Controls.Add(descriptionLabel);

            var currentY = 60;
            foreach (var path in SwaggerRegistry.Api.Paths)
            {
                foreach (var operation in path.Value.Operations)
                {
                    var operationLabel = new Label();
                    operationLabel.Text = operation.Value.OperationId;
                    operationLabel.Location = new Point(8, currentY);
                    operationLabel.ForeColor = Color.FromArgb(117, 11, 28);
                    operationLabel.BackColor = Color.FromArgb(255, 255, 255);
                    operationLabel.Padding = new Padding(8);
                    operationLabel.Size = new Size(width, TOTAL_LABEL_HEIGHT);
                    Controls.Add(operationLabel);

                    currentY = currentY + TOTAL_LABEL_HEIGHT + 16;

                    var addOperationButton = new Button();
                    addOperationButton.Text = "Add to worksheet";
                    addOperationButton.Location = new Point(8, currentY);
                    addOperationButton.FlatStyle = FlatStyle.Flat;
                    addOperationButton.FlatAppearance.BorderColor = Color.FromArgb(22, 21, 20);
                    addOperationButton.FlatAppearance.BorderSize = 1;
                    //addOperationButton.ForeColor = Color.FromArgb(117, 11, 28);
                    addOperationButton.Size = new Size(width - 16, TOTAL_LABEL_HEIGHT);
                    addOperationButton.Click += (sender, e) => AddOperationButton_Click(operation.Value);
                    Controls.Add(addOperationButton);

                    currentY = currentY + TOTAL_LABEL_HEIGHT + 16;
                }
            }
        }

        private void AddOperationButton_Click(OpenApiOperation operation)
        {
            ExcelSheetWriter.AddOperationToSheet(operation);
        }

        private void onSizeChanged(object sender, EventArgs e)
        {
            Draw();
        }
    }
}

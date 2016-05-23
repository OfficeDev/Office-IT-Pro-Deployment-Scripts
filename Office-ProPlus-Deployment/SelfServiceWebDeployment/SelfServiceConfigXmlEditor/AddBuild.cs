using System;
using System.Windows.Forms;

namespace SelfServiceConfigXmlEditor
{
    public partial class AddBuild : Form
    {
        public AddBuild()
        {
            InitializeComponent();
        }

        public Build NewBuild { get; set; }

        private void BtAdd_Click(object sender, EventArgs e)
        {
            try
            {
                NewBuild = new Build()
                {
                    ID = txtDisplayName.Text.Trim() + txtLocation.Text.Trim(),
                    DisplayName = txtDisplayName.Text.Trim(),
                    Location = txtLocation.Text.Trim()
                };
                this.DialogResult = DialogResult.OK;
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR: " + ex.Message);
            }
        }

        private void BtCancel_Click(object sender, EventArgs e)
        {
            try
            {


            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR: " + ex.Message);
            }
        }

    }
}

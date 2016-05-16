using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SelfServiceConfigXmlEditor
{
    public partial class Form1 : Form
    {
        private SelfServiceConfig _selfServiceConfig = null;
        private string xmlPath = "";

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                dgvBuilds.AutoGenerateColumns = false;
                dgvBuilds.SelectionChanged += DgvBuilds_SelectionChanged;
                dgvBuilds.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

                xmlPath = Directory.GetCurrentDirectory() + @"\SelfServiceConfig.xml";

                _selfServiceConfig = new SelfServiceConfig();
                _selfServiceConfig.Load(xmlPath);

                LoadBuilds();
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR: " + ex.Message);
            }
        }

        private void LoadBuilds()
        {
            dgvBuilds.DataSource = _selfServiceConfig.Builds;
        }

        #region "Events"

        private void btRefresh_Click(object sender, EventArgs e)
        {
            try
            {
                var xmlPath = Directory.GetCurrentDirectory() + @"\SelfServiceConfig.xml";
                _selfServiceConfig.Load(xmlPath);

                LoadBuilds();
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR: " + ex.Message);
            }
        }

        private void DgvBuilds_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                if (dgvBuilds.SelectedRows.Count == 1)
                {
                    var build = (Build) dgvBuilds.SelectedRows[0].DataBoundItem;
                    dgvLanguages.DataSource = build.Languages;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR: " + ex.Message);
            }
        }

        private void btAddBuild_Click(object sender, EventArgs e)
        {
            try
            {
                var addBuild = new AddBuild
                {
                    StartPosition = FormStartPosition.CenterParent
                };
                var dialogResult = addBuild.ShowDialog();

                if (dialogResult == DialogResult.OK)
                {
                    addBuild.NewBuild.Languages = new List<Language>()
                    {
                        new Language()
                        {
                            ID = "English (en-us)"
                        }
                    };
                    addBuild.NewBuild.Filters = new List<string>();
                    _selfServiceConfig.AddBuild(addBuild.NewBuild);
                    _selfServiceConfig.Save(xmlPath);
                    LoadBuilds();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR: " + ex.Message);
            }
        }

        private void btRemoveBuild_Click(object sender, EventArgs e)
        {
            try
            {
                if (dgvBuilds.SelectedRows.Count > 0)
                {
                    var buildNames = "Do you want to remove the following Builds?" + Environment.NewLine + Environment.NewLine;
                    foreach (DataGridViewRow buildRow in dgvBuilds.SelectedRows)
                    {
                        var build = (Build)buildRow.DataBoundItem;
                        var buildName = build.DisplayName + " " + build.Location;

                        if (!string.IsNullOrEmpty(buildNames))
                        {
                            buildNames += Environment.NewLine;
                        }

                        buildNames += buildName;
                    }

                    var result = MessageBox.Show(buildNames, "Remove Builds", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                    {
                        foreach (DataGridViewRow buildRow in dgvBuilds.SelectedRows)
                        {
                            var build = (Build)buildRow.DataBoundItem;
                            if (build == null) continue;
                            _selfServiceConfig.RemoveBuild(build);
                        }

                        _selfServiceConfig.Save(xmlPath);
                        LoadBuilds();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR: " + ex.Message);
            }
        }

        #endregion


    }
}

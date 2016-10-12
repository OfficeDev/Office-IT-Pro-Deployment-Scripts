using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using MetroDemo.Models;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Annotations;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Enums;
using Microsoft.OfficeProPlus.InstallGenerator.Models;
using Org.BouncyCastle.Utilities.Collections;

namespace Microsoft.OfficeProPlus.InstallGen.Presentation.Models
{
    public class SccmConfiguration : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged = delegate { };

        private ObservableCollection<Language> _languages = new ObservableCollection<Language>();
        private ObservableCollection<Product> _products = new ObservableCollection<Product>();


        public SccmConfiguration()
        {
            Bitnesses = new List<Bitness>();
            Channels = new List<SelectedChannel>();
            DeploymentDirectory = string.Empty;
            DeploymentSource = DeploymentSource.CDN;
            Languages = new ObservableCollection<Language>();
            Products = new ObservableCollection<Product>();
            Version = BranchVersion.Latest;
            ExcludedProducts = new List<ExcludeProduct>();
            MoveFiles = true;
            UpdateOnlyChangedBits = false;

            //textbox
            DistributionPointGroupName = string.Empty;
            DistributionPoint = string.Empty;
            DeploymentExpiryDurationInDays = 15;
            CustomPackageShareName = "OfficeDeployment";
            SiteCode = string.Empty;
            CMPSModulePath = string.Empty;
            ScriptName = "CM-OfficeDeploymentScript.ps1";
            ConfigurationXml = @".\DeploymentFiles\DefaultConfiguration.xml ";
            CustomName = string.Empty;
            Collection = string.Empty;

            //dropdowns
            ProgramType = ProgramType.DeployWithScript;
            DeploymentPurpose = DeploymentPurpose.Required;
            DeploymentType = DeploymentType.DeployWithConfigurationFile;
        }

        public List<SelectedChannel> Channels { get; set; }
        
        public List<Bitness>  Bitnesses { get; set; }

        public ObservableCollection<Language> Languages
        {
            get { return _languages; }
            set
            {
                _languages = value;
                OnPropertyChanged();
            }
        }

        public ObservableCollection<Product> Products
        {
            get { return _products; }
            set
            {
                _products = value; 
                OnPropertyChanged();
            }
        }

        public List<ExcludeProduct> ExcludedProducts { get; set; }

        public string DeploymentDirectory { get; set; }

        public SccmScenario Scenario { get; set; }

        public DeploymentSource DeploymentSource { get; set; }

        public BranchVersion Version { get; set; }

        public ProgramType ProgramType { get; set; }

        public DeploymentPurpose DeploymentPurpose { get; set; }

        public DeploymentType DeploymentType { get; set; }

        public string DistributionPoint { get; set; }

        public string DistributionPointGroupName { get; set; }

        public string CustomPackageShareName { get; set; }

        public string SiteCode { get; set; }

        public string CMPSModulePath { get; set; }

        public string ConfigurationXml { get; set; }

        public string ScriptName { get; set; }

        public string CustomName { get; set; }

        public string Collection { get; set; }

        public bool MoveFiles { get; set; }

        public bool UpdateOnlyChangedBits { get; set; }

        public int DeploymentExpiryDurationInDays { get; set; }


        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}

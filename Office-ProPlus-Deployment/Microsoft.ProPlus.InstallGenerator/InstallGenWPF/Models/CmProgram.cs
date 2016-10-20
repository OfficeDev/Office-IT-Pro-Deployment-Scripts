using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.Remoting.Channels;
using System.Text;
using System.Threading.Tasks;
using MetroDemo.Models;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Annotations;
using Microsoft.OfficeProPlus.InstallGen.Presentation.Enums;
using Microsoft.OfficeProPlus.InstallGenerator.Models;
using Microsoft.VisualBasic;

namespace Microsoft.OfficeProPlus.InstallGen.Presentation.Models
{
    public class CmProgram : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged = delegate { };

        private ObservableCollection<Language> _languages = new ObservableCollection<Language>();
        private ObservableCollection<Product> _products = new ObservableCollection<Product>();
        private ObservableCollection<string> _collections = new ObservableCollection<string>();


        public CmProgram()
        {
            CollectionNames = new ObservableCollection<string>();
            ScriptName = "CM-OfficeDeploymentScript.ps1";
            ConfigurationXml = @".\DeploymentFiles\DefaultConfiguration.xml ";
            CustomName = string.Empty;
            DeploymentPurpose = DeploymentPurpose.Required;
            DeploymentType = DeploymentType.DeployWithConfigurationFile;
            Bitnesses = new List<Bitness>();
            Channels = new List<SelectedChannel>();
            Languages = new ObservableCollection<Language>();
            Products= new ObservableCollection<Product>();
        }

        public List<SelectedChannel> Channels { get; set; }

        public List<Bitness> Bitnesses { get; set; }

        public ObservableCollection<string> CollectionNames {

            get { return _collections; }
            set
            {
                _collections = value;
                OnPropertyChanged();
            }
        }

        public string ScriptName { get; set;}

        public string ConfigurationXml { get; set; }

        public string CustomName { get; set; }

        public DeploymentPurpose DeploymentPurpose { get; set;}

        public DeploymentType DeploymentType { get; set; }

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

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

    }
}

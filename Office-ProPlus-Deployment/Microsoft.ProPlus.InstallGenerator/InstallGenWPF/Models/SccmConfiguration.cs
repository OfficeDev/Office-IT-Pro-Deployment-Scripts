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

namespace Microsoft.OfficeProPlus.InstallGen.Presentation.Models
{
    public class SccmConfiguration : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged = delegate { };

        private ObservableCollection<Language> _languages = new ObservableCollection<Language>();
        private ObservableCollection<Product> _products = new ObservableCollection<Product>();


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

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}

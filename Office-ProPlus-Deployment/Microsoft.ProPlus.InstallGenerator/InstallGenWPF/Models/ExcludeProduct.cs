using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using MetroDemo.Annotations;

namespace MetroDemo.Models
{
    public class ExcludeProduct : INotifyPropertyChanged
    {
        private string _id;
        private string _displayName;
        private ExcludedStatus _excludedStatus;
        private bool _included = true;

        public string Id
        {
            get { return _id; }
            set
            {
                if (value == _id) return;
                _id = value;
                OnPropertyChanged();
            }
        }

        public string DisplayName
        {
            get { return _displayName; }
            set
            {
                if (value == _displayName) return;
                _displayName = value;
                OnPropertyChanged();
            }
        }

        public bool Included
        {
            get { return _included; }
            set { _included = value; }
        }

        public ExcludedStatus Status
        {
            get {
                return Included ? ExcludedStatus.Included : ExcludedStatus.Excluded;
            }
            set
            {
                if (value == _excludedStatus) return;
                _excludedStatus = value;
                OnPropertyChanged();
            }
        }
        
        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            var handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }
    }

}
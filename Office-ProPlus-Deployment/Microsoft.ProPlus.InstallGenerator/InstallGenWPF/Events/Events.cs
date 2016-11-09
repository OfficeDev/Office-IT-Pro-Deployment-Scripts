using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace MetroDemo.Events
{
    public delegate void InstallOfficeEventHandler(object sender, InstallOfficeEventArgs e);

    public delegate void RestartEventHandler(object sender, EventArgs e);

    public delegate void BranchChangedEventHandler(object sender, BranchChangedEventArgs e);

    public delegate void TransitionTabEventHandler(object sender, TransitionTabEventArgs e);

    public delegate void XmlImportedEventHandler(object sender, EventArgs e);

    public delegate void MessageEventHandler(object sender, MessageEventArgs e);

    public class MessageEventArgs : EventArgs
    {
        public string Title { get; set; }

        public string Message { get; set; }
    }

    public class InstallOfficeEventArgs : EventArgs
    {
        public string Xml { get; set; }

    }

    public class BranchChangedEventArgs : EventArgs
    {
        public string BranchName { get; set; }

    }

    public class TransitionTabEventArgs : EventArgs
    {
        public int Index { get; set; }

        public TransitionTabDirection Direction { get; set; }

        [DefaultValue(false)]
        public bool UseIndex { get; set; }

    }

    public delegate void PropertyValueChangedEventHandler(object sender, PropertyValueChangedEventArgs e);

    public class PropertyValueChangedEventArgs : EventArgs
    {
        public string Name { get; set; }

        public string Value { get; set; }
    }


    public enum TransitionTabDirection
    {
        Forward = 0,
        Back = 1
    }

}

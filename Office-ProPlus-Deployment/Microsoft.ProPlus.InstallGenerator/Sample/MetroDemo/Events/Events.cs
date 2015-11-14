using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MetroDemo.Events
{
    public delegate void BranchChangedEventHandler(object sender, BranchChangedEventArgs e);

    public delegate void TransitionTabEventHandler(object sender, TransitionTabEventArgs e);

    public class BranchChangedEventArgs : EventArgs
    {
        public string BranchName { get; set; }

    }

    public class TransitionTabEventArgs : EventArgs
    {
        public int Index { get; set; }

        public TransitionTabDirection Direction { get; set; }

    }


    public enum TransitionTabDirection
    {
        Forward = 0,
        Back = 1
    }

}

using AutoCADConversion.ViewModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace AutoCADConversion
{
    public class VariableMapperViewModel : INotifyPropertyChanged 
    {
        public VariableViewModel DestinationVariable { get; set; }
        public VariableViewModel SourceVariable { get; set; }
        public string Value { get; set; }

        public bool MapVariable { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;
    }
}

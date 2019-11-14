using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using PropertyChanged;
namespace AutoCADConversion.ViewModel
{
    
    public class VariableViewModel : INotifyPropertyChanged
    {
        public string Name { get; set; }
        public int Id { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace AutoCADConversion
{

    public class TitleBlockDefViewModel : INotifyPropertyChanged
    {
        string BlockName { get; set; }
        string CTBFilePath { get; set; }
        string PaperSize { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;
    }
}

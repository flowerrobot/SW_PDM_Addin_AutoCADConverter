using EPDM.Interop.epdm;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace AutoCADConversion.ViewModel
{

    class SettingsViewModel : AutoCADTaskSettings, INotifyPropertyChanged
    {
        public SettingsViewModel(AutoCADTaskSettings source, IEdmVault20 vault)
        {
            OutputPath = source.OutputPath;
            MenuName = source.MenuName;
            CreateMenu = source.CreateMenu;
            MenuDescription = source.MenuDescription;
            CreatePDF = source.CreatePDF;
            foreach (var b in source.Blocks)
                Blocks.Add(b.Value);

            foreach (var v in source.Variables)
                Variables.Add(v);

            IEdmVariableMgr7 variableMgr = (IEdmVariableMgr7)vault;
            IEdmPos5 pos = variableMgr.GetFirstVariablePosition();
            while (!pos.IsNull)
            {
                IEdmVariable5 var = variableMgr.GetNextVariable(pos);
                AllVariables.Add(new VariableViewModel() { Name = var.Name, Id = var.ID });
            }

        }
        public new ObservableCollection<TitleBlockDefViewModel> Blocks { get; set; } = new ObservableCollection<TitleBlockDefViewModel>();
        public new ObservableCollection<VariableMapperViewModel> Variables { get; set; } = new ObservableCollection<VariableMapperViewModel>();
        public  ObservableCollection<VariableViewModel> AllVariables { get; set; } = new ObservableCollection<VariableViewModel>();

        public event PropertyChangedEventHandler PropertyChanged;
    }
   
}

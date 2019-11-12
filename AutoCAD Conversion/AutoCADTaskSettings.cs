using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AutoCADConversion
{

    internal class AutoCADTaskSettings
    {
        public const string Acadtask_Settings = "AcadSettings";
        
        

        public string TaskName { get; set; }
        public string OutputPath { get; set; }
        public string MenuName { get; set; }
        public bool CreateMenu { get; set; }
        public string MenuDescription { get; set; }
        public bool CreatePDF { get; set; }

        public Dictionary<string, TitleBlockDef> Blocks { get; set; } = new Dictionary<string, TitleBlockDef>();
    }
    public struct TitleBlockDef
    {
        string BlockName;
        string CTBFilePath;
        string PaperSize;
    }
}

﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AutoCADConversion
{

    public class AutoCADTaskSettings
    {
        public const string Acadtask_Settings = "AcadSettings";
        public string TaskName { get; set; }
        public string OutputPath { get; set; }
        public string MenuName { get; set; }
        public bool CreateMenu { get; set; }
        public string MenuDescription { get; set; }
        public bool CreatePDF { get; set; }

        public Dictionary<string, TitleBlockDefViewModel> Blocks { get; set; } = new Dictionary<string, TitleBlockDefViewModel>();

        public List<VariableMapperViewModel> Variables { get; set; } = new List<VariableMapperViewModel>();
        public string AutoCADCorePath { get; set; } = @"C:\Program Files\Autodesk\AutoCAD 2018\accoreconsole.exe";
    }

}

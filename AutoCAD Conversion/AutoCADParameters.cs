using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AutoCADConversion
{


    /// <summary>
    /// This isn't actually required as no data is required to be stored
    /// </summary>
     struct AutoCADParameters
    {
        public  int FileId { get; set; }
        public int FolderId { get; set; }
        public int Version { get; set; }
        public string FileName { get; set; }
        public string OutputPath { get; set; }
        public string OutputPathDXF { get; set; }
    }
}

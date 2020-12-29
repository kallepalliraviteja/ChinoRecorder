using System;
using System.Collections.Generic;
using System.Text;

namespace ChinoRecorder
{
    public class Config
    {
        public string DB_Name { get; set; }
        public string DB_Location { get; set; }
        public Recorder[] Recorders { get; set; }
    }
}

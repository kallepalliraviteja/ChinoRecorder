using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ChinoRecorder
{
    public class RecorderParameters
    {
        public float MAIN_STEAM_FLOW { get; set; }
        public float LOAD { get; set; }
        public float MAIN_STM_PR { get; set; }
        public float MAIN_STEAM_TEMP { get; set; }
        public float FW_PR_AT_ECOIL { get; set; }
        public float FW_TMP_AT_ECOIL { get; set; }
        public float HRH_PRESSURE { get; set; }
        public float HRH_TEMPERATURE { get; set; }
        public float CRH_PRESSURE { get; set; }
        public float CRH_TEMPERATURE { get; set; }
        public float REHEAT_SPRAY { get; set; }
        public float FEED_WTR_FLOW { get; set; }
        public float EXT_PR_HPH6_IL { get; set; }
        public float EXT_TEMP_HPH6IL { get; set; }
        public float HPH6_DRIP_TEMP { get; set; }
        public float BFP_DISCH_HDRPR { get; set; }
        public float FW_TMP_HPH6_IL { get; set; }
        public float FW_TEMP_HPH6_OL { get; set; }
        public float Per_DM_MAKEUP { get; set; }
        public float SH_SPRAY { get; set; }
        public DateTime Timestamp { get; set; }
    }
}

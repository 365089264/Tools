using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace StringMatchFunction
{
    public class MatchResult
    {
        public string EQ_CinemaCode { get; set; }
        public string EQ_CinemaName { get; set; }
        public string EQ_Adress { get; set; }
        public float matchRate { get; set; }
        public string matchFilter { get; set; }
        public string SystemNum { get; set; }
        public string EQ_CinemaNameForward { get; set; }
        public string EQ_ForAddress { get; set; }
    }
}

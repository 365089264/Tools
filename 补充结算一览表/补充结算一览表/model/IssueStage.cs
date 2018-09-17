using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace 补充结算一览表.model
{
    public class IssueStage
    {
        public int issuestageid { get; set; }
        public string issuestageguid { get; set; }
        public string filmname { get; set; }
        public string filmversionname { get; set; }
        public string filmseqcode { get; set; }
        public int filmid { get; set; }
        public string playstarttime { get; set; }
        public string playendtime { get; set; }
    }
}

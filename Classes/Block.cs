using System;
using System.Collections.Generic;
using System.Text;

namespace ResourceScheduler
{
    public class Block
    {
        public int ResourceId { get; set; }
        public int Start { get; set; }
        public int End { get; set; }
        public int Duration { get; set; }
        public bool Valid { get; set; }
    }

}

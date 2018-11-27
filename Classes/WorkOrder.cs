using System;
using System.Collections.Generic;
using System.Text;

namespace ResourceScheduler
{
    public class WorkOrder
    {
        public int Id { get; set; }
        public int? ParentId { get; set; }
        public List<Operation> Operations { get; set; }
        public DateTime StartDateTime { get; set; }
        public DateTime EndDateTime { get; set; }
        public int? Start { get; set; }
        public int? End { get; set; }
        public int Level { get; set; }
    }


}

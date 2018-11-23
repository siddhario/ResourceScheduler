using System;
using System.Collections.Generic;
using System.Text;

namespace ResourceScheduler
{

    public class Operation
    {
        public int Id { get; set; }
        public int ResourceId { get; set; }
        public int Duration { get; set; }
        public DateTime? StartDateTime { get; set; }
        public DateTime? EndDateTime { get; set; }
        public int? Start { get; set; }
        public int? End { get; set; }
        public int WorkOrderId { get; set; }
        public string ResourceName { get; set; }
        public string Description { get; set; }
        public int Level { get; set; }
        public int TransportTime { get; set; }
    }
}

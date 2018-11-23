using ResourceScheduler.Classes;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ResourceScheduler
{
    class Program
    {
        private static Random rand = new Random();
        static List<WorkOrder> orders = new List<WorkOrder>();
        static List<Block> blocks = new List<Block>();
        static int ordersCount = 500;
        static int operationsPerOrder = 10;
        static DateTime refTime = new DateTime(2018, 12, 31);
        static int resourceCount = 20;
        static List<Pause> dailyPauses = new List<Pause>();
        static List<Pause> pauses = new List<Pause>();

        static void Main(string[] args)
        {
            Init();

            DateTime startT = DateTime.Now;
            foreach (var order in orders)
            {
                int end = order.End.Value;
                foreach (var operation in order.Operations.OrderByDescending(o => o.Id))
                    end = ScheduleOperation(operation, end);
                order.Start = end;
            }
            DateTime endT = DateTime.Now;

            Console.WriteLine("Total time:" + endT.Subtract(startT).TotalSeconds + " s");

            List<Operation> operations = new List<Operation>();
            foreach (var order in orders)
            {
                // Console.WriteLine("Order:" + order.Id+ " Start:" + refTime.AddMinutes(-order.Start.Value).ToString("dd.MM.yyyy HH:mm")+ " End:  " + refTime.AddMinutes(-order.End.Value).ToString("dd.MM.yyyy HH:mm"));
                //Console.WriteLine("Start:" + refTime.AddMinutes(-order.Start.Value).ToString("dd.MM.yyyy HH:mm"));
                //Console.WriteLine("End:  " + refTime.AddMinutes(-order.End.Value).ToString("dd.MM.yyyy HH:mm"));
                foreach (var operation in order.Operations)
                {
                    operations.Add(operation);
                    //Console.WriteLine(" Operation:" + operation.Id+" Resource id:"+operation.ResourceId+ " Start:    " + refTime.AddMinutes(-operation.Start.Value).ToString("dd.MM.yyyy HH:mm")+ "-" + refTime.AddMinutes(-operation.End.Value).ToString("HH:mm")+ " Duration:      " + operation.Duration + " m");
                    //Console.Write(" Start:    " + refTime.AddMinutes(-operation.Start.Value).ToString("dd.MM.yyyy HH:mm"));
                    //Console.WriteLine("-" + refTime.AddMinutes(-operation.End.Value).ToString("HH:mm"));
                    //Console.WriteLine(" Duration:      " + operation.Duration + " m");
                }
            }


            //foreach (var operation in operations.OrderBy(o => o.ResourceId).ThenBy(o => o.Start))
            //{
            //    Console.Write(" Resource id:" + operation.ResourceId + " Order id:" + operation.WorkOrderId);
            //    Console.Write(" Start:    " + refTime.AddMinutes(-operation.Start.Value).ToString("dd.MM.yyyy HH:mm"));
            //    Console.WriteLine("-" + refTime.AddMinutes(-operation.End.Value).ToString("HH:mm"));
            //    //Console.WriteLine(" Duration:      " + operation.Duration + " m");

            //}

            ExcelBuilder.Start(operations);



            Console.WriteLine("Blocks count:" + blocks.Count);
            Console.ReadKey();
        }

        public static int ScheduleOperation(Operation operation, int offset)
        {
            Block block = null;

            //for given resource search valid block with with sufficient duration and offset less then block start, take first from top (latest)
            try
            {
                block = blocks.Where(b =>
                b.Valid == true &&
                b.ResourceId == operation.ResourceId &&
                b.Duration >= operation.Duration &&
                b.Duration - (offset - b.End) >= operation.Duration &&
                b.Start >= offset).OrderBy(b => b.End).First();

            }
            catch (Exception e)
            {
                Console.WriteLine("O:" + offset + ", d:" + operation.Duration);
                foreach (var b in blocks.Where(bl => bl.ResourceId == operation.ResourceId && bl.Valid == true).OrderBy(bs => bs.Start))
                    Console.WriteLine(b.Start.ToString() + "-" + b.End.ToString() + "-d:" + b.Duration);
                Console.ReadKey();
            }

            Pause pause = pauses.Where(x => x.End <= block.End && x.Start > block.End).OrderBy(x => x.End).SingleOrDefault();
            if (pause != null)
                block.End = pause.Start;

            pause = pauses.Where(x => x.End <= offset && x.Start > offset).OrderBy(x => x.End).SingleOrDefault();
            if (pause != null)
                offset = pause.Start;

            int operationEnd = block.End > offset ? block.End : offset;

            int operationStart = operationEnd + operation.Duration;

            pause = pauses.Where(x => x.End <= operationStart && x.Start >= operationStart).OrderBy(x => x.End).SingleOrDefault();
            if (pause != null)
                operationStart = operationStart + pause.Duration;

            //invalidate block
            block.Valid = false;
            //repartition block
            if (operationStart < block.Start)
            {
                Block b = new Block() { ResourceId = block.ResourceId, Valid = true, Start = block.Start, End = operationStart, Duration = block.Start - operationStart };
                updateDuration(b);
                blocks.Add(b);
            }
            if (operationEnd > block.End)
            {
                Block b = new Block() { ResourceId = block.ResourceId, Valid = true, Start = operationEnd, End = block.End, Duration = operationEnd - block.End };
                updateDuration(b);
                blocks.Add(b);
            }

            operation.Start = operationStart;
            operation.StartDateTime = refTime.AddMinutes(-operation.Start.Value);

            operation.End = operationEnd;
            operation.EndDateTime = refTime.AddMinutes(-operation.End.Value);
            return operationStart;
        }

        public static void Init()
        {
            Random r = new Random();
            List<Resource> resources = new List<Resource>();
            for (int i = 0; i < resourceCount; i++)
                resources.Add(new Resource() { ResourceId = i + 1, ResourceName = GetLetter() + GetLetter() + GetLetter() + "-" + (rand.Next(100, 999)).ToString() });


            for (int i = 0; i < ordersCount; i++)
            {
                List<Operation> operations = new List<Operation>();
                for (int j = 0; j < operationsPerOrder; j++)
                {
                    Resource res = resources[rand.Next(0, resourceCount - 1)];
                    operations.Add(new Operation() { Level = 1, WorkOrderId = i + 1, Id = j + 1, ResourceId = res.ResourceId, ResourceName = res.ResourceName, Duration = r.Next(1, 360), Start = null, End = null });
                }
                orders.Add(new WorkOrder() { Id = i + 1, Operations = operations, ParentId = null, End = r.Next(14400) });
            }

            TimeSpan timeInterval = refTime - new DateTime(2016, 12, 31);
            int time = timeInterval.Days * 1440 + timeInterval.Hours * 60 + timeInterval.Minutes;

            for (int i = 0; i < resourceCount; i++)
                blocks.Add(new Block() { Start = time, End = 0, Duration = time, ResourceId = i + 1, Valid = true });




            pauses.Add(new Pause() { Start = 75, End = 0, Duration = 75 });  // 00:00 - 22:45
            dailyPauses.Add(new Pause() { Start = 360, End = 330, Duration = 30 }); // 18:30 - 18:00
            dailyPauses.Add(new Pause() { Start = 550, End = 540, Duration = 15 }); // 15:00 - 14:45
            dailyPauses.Add(new Pause() { Start = 900, End = 870, Duration = 30 }); // 09:30 - 09:00
            dailyPauses.Add(new Pause() { Start = 1515, End = 1020, Duration = 495 }); // 07:00 - 22:45

            for (int i = 0; i < timeInterval.Days + 1; i++)
            {
                foreach (Pause p in dailyPauses)
                {
                    pauses.Add(new Pause()
                    {
                        Start = p.Start + 1440 * i,
                        End = p.End + 1440 * i,
                        Duration = p.Duration
                    });
                }
            }
            List<DateTime> wDays = null;
            wDays = weekDays(new DateTime(2016, 12, 31), refTime);

            foreach (Block b in blocks)
            {
                updateDuration(b);
            }
        }

        public static void updateDuration(Block b)
        {
            int pauseTime = pauses.Where(x => x.End >= b.End && x.Start <= b.Start).Sum(x => x.Duration);
            b.Duration = b.Duration - pauseTime;
        }

        public static List<DateTime> weekDays(DateTime start, DateTime end)
        {
            List<DateTime> weekDays = new List<DateTime>();
            double day = -1;
            while (start<= end)
            {
                if (end.DayOfWeek == DayOfWeek.Saturday || end.DayOfWeek == DayOfWeek.Sunday)
                    weekDays.Add(end);
                end.AddDays(day);
            }
            return weekDays;
        }



        public static string GetLetter()
        {
            string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            int num = rand.Next(0, chars.Length - 1);
            return chars[num].ToString();
        }
    }
}

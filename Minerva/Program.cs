using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CommandLine;


namespace Minerva.Main
{

    using Minerva.Weekly;
    using Minerva.Project;
    using Minerva.Summary;

    class Program
    {

        static void Phase1()
        {
            //Weekly weekly = new Weekly();
            //ProjectPlan plan = new ProjectPlan();
            //plan.CompareWith(weekly).ReNewProjectPlan();
            WeeklySummary summary = new WeeklySummary();
            summary.Summarize();
        }

        static void Phase2()
        {
            Weekly weekly = new Weekly();
            weekly.ToSortedWeeklyList().Summarize();
        }

        static void Phase3()
        {
            Weekly weekly = new Weekly();
            Summary summary = new Summary(weekly);
            SummaryDocument document = new SummaryDocument(summary);
            document.Perform();
        }

        static void Main(string[] args)
        {
            ModifyInMemory.ActivateMemoryPatching();

            Options options = new Options();
            Parser.Default.ParseArguments<Options>(args)
                   .WithParsed<Options>(o =>
                   {
                       Env.Instance.Initialize(o.ReportDir, o.Phase);
                   });

            switch (Env.Instance.Phase)
            {
                case 1: Phase1(); break;
                case 2: Phase2(); break;
                case 3: Phase3(); break;
                default: Console.WriteLine("please enter 1,2 or 3..."); break;

            }






        }



    }
}

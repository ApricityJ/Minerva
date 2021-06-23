using System;
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
            Console.WriteLine("phase 1 start ... ");

            DevWeeklies devWeeklies = new DevWeeklies();
            devWeeklies.Load().Summarize();

            new ProjectPlan()
                .CompareWith(devWeeklies)
                .ReNewProjectPlan();
        }

        static void Phase2()
        {

        }

        static void Phase3()
        {
        
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

            Console.WriteLine("press any key to exit ...");
            Console.ReadKey();






        }



    }
}

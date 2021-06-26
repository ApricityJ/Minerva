using System;
using CommandLine;


namespace Minerva.Main
{

    using Weekly;
    using Project;
    using Doc;

    class Program
    {
        //比对项目计划与周报中的项目部分，生成草稿
        static void Phase1()
        {
            Console.WriteLine("phase 1 start ... ");

            DevWeeklies devWeeklies = new DevWeeklies();
            devWeeklies.Load().Summarize();

            new ProjectPlan()
                .CompareWith(devWeeklies)
                .ReNewProjectPlan();

            Console.WriteLine("phase 1 complete ... ");
        }

        //生成excel文件 [科技与产品管理部周报]
        static void Phase2()
        {
            Console.WriteLine("phase 2 start ... ");

            new WeeklyDocument()
                .PrepareDataSet()
                .Perform();

            Console.WriteLine("phase 2 complete ... ");
        }

        //生成word文件 [开发工作每周概况]
        static void Phase3()
        {
            new DevBriefDocument()
                .PrepareDataSet()
                .Perform();
        }

        static void Minerva(string[] args)
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

        static void Main(string[] args)
        {
            Minerva(args);
        }



    }
}

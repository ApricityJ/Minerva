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

        static void Main(string[] args)
        {
            ModifyInMemory.ActivateMemoryPatching();

            Options options = new Options();
            Parser.Default.ParseArguments<Options>(args)
                   .WithParsed<Options>(o =>
                   {
                       Env.Instance.Initialize(o.ReportDir);
                   });

            //读取周报(们)
            Weekly weekly = new Weekly();

            //读取项目计划
            ProjectPlan plan = new ProjectPlan();

            //比较并更新项目计划
            plan.CompareWith(weekly).ReNewProjectPlan();

            //生成汇总报告
            Summary summary = new Summary(weekly);
            SummaryDocument document = new SummaryDocument(summary);
            document.Perform();




        }



    }
}

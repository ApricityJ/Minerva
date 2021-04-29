﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CommandLine;

namespace Minerva
{

    using Report;


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

            ProjectPlan projectPlan = new ProjectPlan();
            projectPlan.CompareWithWorkReport().AppendToProjectPlan();
        }


        
    }
}

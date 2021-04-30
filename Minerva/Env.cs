﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Minerva
{
    class Env
    {
        private Env()
        {
        }

        public static Env Instance { get; } = new Env();

        private string FindFileByName(string keyWord)
        {
            DirectoryInfo directory = new DirectoryInfo(WeeklyReportsDir);
            FileInfo[] files = directory.GetFiles(keyWord, SearchOption.AllDirectories);
            if(files.Length > 0)
            {
                return files[0].FullName;
            }
            return "";
        }

        public void Initialize(string dir)
        {
            WeeklyReportsDir = dir;
            ProjectPlanPath = FindFileByName("*项目*");
            ReportDivsion1st = FindFileByName("*开发一部*");
            ReportDivsion2nd = FindFileByName("*开发二部*");
            ReportDivsion3rd = FindFileByName("*开发三部*");
            ReportDivsionData = FindFileByName("*数据分析部*");
        }

        public string WeeklyReportsDir { get; set; }
        public string ProjectPlanPath { get; set; }
        public string ReportDivsion1st { get; set; }
        public string ReportDivsion2nd { get; set; }
        public string ReportDivsion3rd { get; set; }
        public string ReportDivsionData { get; set; }

        public int SummarySequence { get; set; }





    }
}
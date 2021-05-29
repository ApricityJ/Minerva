﻿using System.IO;


namespace Minerva
{
    /// <summary>
    /// 运行环境类，用于存储运行中所需要的常量，全局唯一
    /// </summary>
    class Env
    {
        private Env()
        {
        }

        public static Env Instance { get; } = new Env();

        //根据关键词搜索文件
        private string FindFileByName(string keyWord)
        {
            DirectoryInfo directory = new DirectoryInfo(WeeklyReportsDir);
            FileInfo[] files = directory.GetFiles(keyWord, SearchOption.AllDirectories);
            if (files.Length > 0)
            {
                return files[0].FullName;
            }
            return "";
        }

        public void Initialize(string dir,int phase)
        {
            WeeklyReportsDir = dir;
            Phase = phase;
            ProjectPlanPath = FindFileByName("*项目*");
            ReportDivsion1st = FindFileByName("*开发一部*");
            ReportDivsion2nd = FindFileByName("*开发二部*");
            ReportDivsion3rd = FindFileByName("*开发三部*");
            ReportDivsionData = FindFileByName("*数据分析部*");
            ReportTest = FindFileByName("*测试*");
            ReportSystem = FindFileByName("*系统*");
            ReportNet = FindFileByName("*网络*");
            ReportProduct = FindFileByName("*产品部*");
        }

        public string WeeklyReportsDir { get; set; }
        public string ProjectPlanPath { get; set; }
        public string ReportDivsion1st { get; set; }
        public string ReportDivsion2nd { get; set; }
        public string ReportDivsion3rd { get; set; }
        public string ReportDivsionData { get; set; }
        public string ReportTest { get; set; }
        public string ReportSystem { get; set; }
        public string ReportNet { get; set; }
        public string ReportProduct { get; set; }

        public int SummarySequence { get; set; }

        public int Phase { get; set; }





    }
}

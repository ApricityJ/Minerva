using System.IO;
using System.Collections.Generic;
using System.Linq;

namespace Minerva
{
    using Department;
    /// <summary>
    /// 运行环境类，用于存储运行中所需要的常量，全局唯一
    /// </summary>
    class Env
    {
        private static Dictionary<InnerDepartment, string> keyWordMap = new Dictionary<InnerDepartment, string>()
        {
            {InnerDepartment.DevDivision1st,"*一部*" },
            {InnerDepartment.DevDivision2nd,"*二部*" },
            {InnerDepartment.DevDivision3rd,"*三部*" },
            {InnerDepartment.DevDataSci,"*数据分析*" },
            {InnerDepartment.Test,"*测试*" },
            {InnerDepartment.System,"*系统*" },
            {InnerDepartment.Network,"*网络*" },
            {InnerDepartment.Product,"*产品、*" }
        };


        private Dictionary<InnerDepartment, string> weeklyMap = new Dictionary<InnerDepartment, string>()
        {

        };


        private Env()
        {

        }

        public static Env Instance { get; } = new Env();

        //根据关键词搜索文件
        private string FindFileBy(string keyWord)
        {
            return FindFileBy(RootDir,keyWord);
        }

        private string FindFileBy(string dir, string keyWord)
        {
            DirectoryInfo directory = new DirectoryInfo(dir);
            FileInfo[] files = directory.GetFiles(keyWord, SearchOption.AllDirectories);
            if (files.Length > 0)
            {
                return files[0].FullName;
            }
            return "";
        }

        public void Initialize(string dir, int phase)
        {
            RootDir = dir;
            Phase = phase;

            keyWordMap.ToList().ForEach(pair =>
            {
                weeklyMap.Add(pair.Key, FindFileBy(pair.Value));
            });

            ProjectPlan = FindFileBy("tpl", "*项目计划*");
        }

        public string ToWeeklyPath(InnerDepartment department)
        {
            return weeklyMap[department];
        }

        public string RootDir { get; set; }
        public string ProjectPlan { get; set; }
        public int BriefSequence { get; set; }
        public int Phase { get; set; }





    }
}

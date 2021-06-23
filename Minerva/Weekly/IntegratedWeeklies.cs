using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;

namespace Minerva.Weekly
{
    using Department;
    using Util;

    /// <summary>
    /// 周报集成类，包含一个开发类周报集合和一个基础类周报集合
    /// 用于生成 [科技与产品管理部周报(#Year#)年第(#Week#)期.et]
    /// </summary>

    class IntegratedWeeklies : AbstractWeeklies
    {

        public DevWeeklies DevelopWeeklies { get; set; }

        public BaseWeeklies NonDevelopWeeklies { get; set; }


        private string template = "科技与产品管理部周报(#Year#)年第(#Week#)期.et";
        private string targetPath;

        public IntegratedWeeklies()
        {
            WeeklySet = new Dictionary<InnerDepartment, BaseWeekly>();

            DevelopWeeklies = new DevWeeklies();
            NonDevelopWeeklies = new BaseWeeklies();

        }

        //生成目标周报文件路径，用实际的年和周替换
        private void ToTargetPath()
        {
            targetPath = template.Replace("#Year#", DateUtil.ToCurrentYear())
                .Replace("#Week#", DateUtil.ToWeekOfYear().ToString());
            targetPath = Path.Combine(Env.Instance.RootDir, targetPath);

            if (File.Exists(targetPath))
            {
                File.Delete(targetPath);
            }
        }

        //按内部机构，逐个Load各部门的周报文件
        public override AbstractWeeklies Load()
        {
            
            DevelopWeeklies.Load();
            NonDevelopWeeklies.Load();

            return this;
        }

        public override AbstractWeeklies Summarize()
        {

            DevelopWeeklies.Summarize();
            NonDevelopWeeklies.Summarize();

            return this;
        }


        public override AbstractWeeklies Sort()
        {
            DevelopWeeklies.Sort();
            NonDevelopWeeklies.Sort();

            return this;
        }

        public override AbstractWeeklies Save()
        {
            throw new NotImplementedException();
        }
    }
}

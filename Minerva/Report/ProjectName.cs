using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Minerva.Report
{
    /// <summary>
    /// 项目名称类，用于处理项目名称的模糊匹配等问题
    /// </summary>
    class ProjectName
    {
        private List<string> PlanProjNameList;

        public ProjectName(List<string> nameList)
        {
            PlanProjNameList = nameList;
        }

        //丢弃所有括号内内容
        private string DiscardParentheseWithin(string s)
        {
            s = Regex.Replace(s, @"\(.*\)", "");
            s = Regex.Replace(s, @"（.*）", "");
            return s;
        }

        //计算一个周报中的项目名称和一个项目计划名称的相似度
        private double ToSimilarityBetween(string planProjName, string reportProjName)
        {
            int containsCount = 0;
            reportProjName.ToList().ForEach(c =>
            {
                if (planProjName.Contains(c))
                {
                    containsCount++;
                }
            });

            return containsCount / reportProjName.Length;
        }

        //将一个周报中的项目名称尝试转换为一个标准项目名称(以项目计划为准)
        public string TryToProjectPlanName(string reportProjName)
        {
            double threshold = 0.75;
            string neatName = DiscardParentheseWithin(reportProjName);
            
            //对每个周报项目名称，计算与项目计划名称的相似度，选择大于阈值的项目名称作为候选清单，按相似度倒序排序
            List<string> candidateList = PlanProjNameList.Select(planProjName =>
           {
               return new KeyValuePair<string, double>(planProjName, ToSimilarityBetween(planProjName, reportProjName));
           })
            .Where(pair => { return pair.Value >= threshold; })
            .OrderByDescending(pair => pair.Value)
            .Select(pair => pair.Key)
            .ToList();

            //如果候选清单不为空，返回相似度最高的名称
            if(candidateList.Count > 0)
            {
                return candidateList[0];
            }

            return reportProjName;
        }



    }
}

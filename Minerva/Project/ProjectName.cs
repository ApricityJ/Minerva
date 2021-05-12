using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Minerva.Project
{
    /// <summary>
    /// 项目名称类，用于处理项目名称的模糊匹配
    /// </summary>
    class ProjectName
    {
        private List<string> PlanProjNameList;

        public ProjectName(List<string> nameList)
        {
            PlanProjNameList = nameList;
        }

        //丢弃所有括号内的内容
        //e.g.
        //DiscardParentheseWithin("人社部金保二期工程银行端项目（开发一部-机构部）") -> 人社部金保二期工程银行端项目
        private string DiscardParentheseWithin(string s)
        {
            s = Regex.Replace(s, @"\(.*\)", "");
            s = Regex.Replace(s, @"（.*）", "");
            return s;
        }

        //计算一个周报中的项目名称和一个项目计划名称的相似度
        //计算方法是粗糙和不精确的
        //e.g.
        //ToSimilarityBetween("纳税e贷","纳税e贷") -> 1.0
        //ToSimilarityBetween("纳税e贷","纳税贷") -> 1.0
        //ToSimilarityBetween("纳税e贷","纳税e贷二期") -> 0.67
        //ToSimilarityBetween("纳税e贷","政务e贷") -> 0.5
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
        //e.g.
        //TryToProjectPlanName("纳税e贷") -> 纳税e贷
        //TryToProjectPlanName("纳税贷") -> 纳税e贷
        //TryToProjectPlanName("纳税e贷二期") -> 纳税e贷二期
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
            if (candidateList.Count > 0)
            {
                return candidateList[0];
            }

            return reportProjName;
        }



    }
}

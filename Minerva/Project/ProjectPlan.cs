using Minerva.DAO;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Minerva.Project
{

    using Minerva.Weekly;
    /// <summary>
    /// 项目计划类，处理项目计划相关的操作
    /// </summary>
    class ProjectPlan
    {
        private ExcelDAO<ProjectPlanItem> dao;
        private List<ProjectPlanItem> projectPlanList;
        private List<WeeklyItem> ignoredWeeklyItemList;
        private List<ProjectPlanItem> ignoredProjectPlanList;
        private List<ProjectPlanItem> existedProjectPlanList;
        private List<WeeklyItem> markableWeeklyItemList;


        private ProjectName projectName;

        public ProjectPlan()
        {
            dao = new ExcelDAO<ProjectPlanItem>(Env.Instance.ProjectPlan);
            
            //读取项目计划excel生成项目列表
            projectPlanList = dao.SelectSheetAt(0)
                .Skip(2)
                .ToEntityList(typeof(ProjectPlanItem))
                //已结项和已投产的项目无需考虑
                .Where(item=>!item.IsReleased())
                .ToList();
            projectName = ToProjectName();
        }

        private ProjectName ToProjectName()
        {
            List<string> nameList = projectPlanList.Select(item => item.ProjectName).ToList();
            return new ProjectName(nameList);
        }

        //项目计划与工作周报交叉比对
        //对于 [周报] 中存在，但[项目计划]中不存在的内容，存放 ignoredWeeklyItemList
        //对于 [周报] 中不存在，但[项目计划]中存在的内容，存放 ignoredProjectPlanList
        //对于 [周报] 中立项和需求中的项目，存放 markableWeeklyItemList，留待人工判断
        public ProjectPlan CompareWith(DevWeeklies weeklies)
        {
            //尝试将周报中的项目名称模糊匹配为项目计划中的项目名称
            weeklies.ProjectWorkList.ForEach(item => { item.Name = projectName.TryToProjectPlanName(item.Name); });

            existedProjectPlanList = weeklies.ProjectWorkList
                .Where(item => IsExist(item.Name) && !item.IsProjectApproved())
                .ToList()
                .Select(item=> ToProjectPlanItem(item))
                .ToList();

            ignoredWeeklyItemList = weeklies.ProjectWorkList
                .Where(item => !IsExist(item.Name)&&!item.IsProjectApproved())
                .ToList();

            ignoredProjectPlanList = projectPlanList
                .Where(item => !weeklies.IsExist(item.ProjectName))
                .ToList();

            markableWeeklyItemList = weeklies.ProjectWorkList
                .Where(item => item.IsProjectApproved())
                .ToList();

            return this;
        }

        //抽取字段内容中的"剩余工作量:"之后的值
        private string ToRemainHumanMonth(string schedule)
        {
            return Regex.Match(schedule, @"(?<=剩余工作量：)\S+").Value;
        }

        //抽取字段内容中的"计划投产时间:"之后的值
        private string ToEstimatedTimeRemaining(string schedule)
        {
            return Regex.Match(schedule, @"(?<=计划投产时间：)\S+").Value;
        }

        //根据工作周报内容返回一个项目计划对象，用于周报中存在，项目计划中也存在的项目
        //并使用工作周报内容更新其[剩余工作量]、[计划投产时间]两个属性
        private ProjectPlanItem ToProjectPlanItem(WeeklyItem weeklyItem)
        {
            ProjectPlanItem projectPlanItem = projectPlanList.Find(item => item.ProjectName.Trim().Equals(weeklyItem.Name));
            projectPlanItem.RemainHumanMonth = ToRemainHumanMonth(weeklyItem.Schedule);
            projectPlanItem.EstimatedTimeRemaining = ToEstimatedTimeRemaining(weeklyItem.Schedule);
            return projectPlanItem;
        }

        //根据工作周报内容生成一个新的项目计划对象，用于周报中存在，但项目计划中不存在的项目
        private ProjectPlanItem ToNewProjectPlanItem(WeeklyItem weeklyItem)
        {
            ProjectPlanItem projectPlanItem = new ProjectPlanItem
            {
                ProjectName = weeklyItem.Name,
                ResponsiblePersonnel = weeklyItem.ResponsiblePersonnel,
                HostDivision = weeklyItem.HostDivision,
                RequirementDepartment = weeklyItem.BizDepartment,
                RemainHumanMonth = ToRemainHumanMonth(weeklyItem.Schedule),
                EstimatedTimeRemaining = ToEstimatedTimeRemaining(weeklyItem.Schedule)
            };
            return projectPlanItem;
        }

        //判断一个项目名称是否存在于项目计划中
        private bool IsExist(string projectName)
        {
            return projectPlanList.Any(projectPlanItem => projectPlanItem.ProjectName.Trim().Equals(projectName));
        }

        private void RenewProjectPlanItem(ProjectPlanItem projectPlanItem)
        {
            int rowIndex = dao.SelectSheetAt(0).FindRowIndex(1, projectPlanItem.ProjectName);
            int columnIndex = typeof(ProjectPlanItem).GetProperties().Length;
            dao.SetCellValue(rowIndex, columnIndex - 1, projectPlanItem.RemainHumanMonth);
            dao.SetCellValue(rowIndex, columnIndex, projectPlanItem.EstimatedTimeRemaining);
        }

        //更新项目计划文件
        public ProjectPlan ReNewProjectPlan()
        {
            //对于匹配成功的项目计划，根据周报更新其剩余工作量及预计投产时间
            existedProjectPlanList.ForEach(item => RenewProjectPlanItem(item));

            //将结果追加于3个新的sheet中并保存文件
            dao.RemoveAt("周报(不存在)项目计划(存在)");
            dao.SetCellValues(ignoredProjectPlanList)
                .Name("周报(不存在)项目计划(存在)")
                .Fit();
            dao.Save();

            dao.Close();


            ExcelDAO<WeeklyItem> weeklyItemDAO = new ExcelDAO<WeeklyItem>(Env.Instance.ProjectPlan);

            weeklyItemDAO.RemoveAt("周报(存在)项目计划(不存在)");
            weeklyItemDAO.SetCellValues(ignoredWeeklyItemList)
                .Name("周报(存在)项目计划(不存在)")
                .Fit();
            weeklyItemDAO.Save();


            weeklyItemDAO.RemoveAt("需求与立项中清单");
            weeklyItemDAO.SetCellValues(markableWeeklyItemList)
                .Name("需求与立项中清单")
                .Fit();
            weeklyItemDAO.Save();

            weeklyItemDAO.Close();

            return this;
        }
    }
}

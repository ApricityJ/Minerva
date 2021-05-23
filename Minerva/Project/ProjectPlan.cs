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
        private readonly ExcelDAO<ProjectPlanItem> dao;
        private List<ProjectPlanItem> projectPlanList;
        private List<ProjectPlanItem> ignoredProjectList;
        private List<ProjectPlanItem> existProjectList;

        private ProjectName projectName;

        public ProjectPlan()
        {
            dao = new ExcelDAO<ProjectPlanItem>(Env.Instance.ProjectPlanPath);

            projectPlanList = dao.SelectSheetAt(0)
                .Skip(2)
                .ToEntityList(typeof(ProjectPlanItem));
            projectName = ToProjectName();
        }

        private ProjectName ToProjectName()
        {
            List<string> nameList = projectPlanList.Select(item => item.ProjectName).ToList();
            return new ProjectName(nameList);
        }

        public List<ProjectPlanItem> ToProjectPlanList()
        {
            return projectPlanList;
        }

        //项目计划与工作周报对比
        //对于 周报中存在，但项目计划未能匹配的项目内容，存放ignoredProjectList
        //对于 周报中存在，项目计划匹配成功的项目内容，更新其字段后，存放existProjectList
        public ProjectPlan CompareWith(Weekly weekly)
        {
            weekly.WeeklyList.ForEach(item => { item.Name = projectName.TryToProjectPlanName(item.Name); });

            ignoredProjectList = weekly.WeeklyList
                .Where(item => !IsExistInProjectPlanList(item.Name))
                .Select(item => ToNewProjectPlanItem(item))
                .ToList();

            existProjectList = weekly.WeeklyList
                .Where(item => IsExistInProjectPlanList(item.Name))
                .Select(item => ToProjectPlanItem(item))
                .ToList();

            return this;
        }

        //抽取字段内容中的"剩余工作量:"之后的值
        //注意，此方法同时存在于Summary类中
        private string ToRemainHumanMonth(string schedule)
        {
            return Regex.Match(schedule, @"(?<=剩余工作量：)\S+").Value;
        }

        //抽取字段内容中的"计划投产时间:"之后的值
        //注意，此方法同时存在于Summary类中
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
        private bool IsExistInProjectPlanList(string projectName)
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

        //更新项目计划文件，将结果追加于两个新的sheet中并保存文件
        public ProjectPlan ReNewProjectPlan()
        {
            dao.SetCellValues(ignoredProjectList);
            dao.Fit();
            dao.SetCellValues(existProjectList);
            dao.Fit();
            existProjectList.ForEach(item => RenewProjectPlanItem(item));
            dao.Save();
            return this;
        }
    }
}

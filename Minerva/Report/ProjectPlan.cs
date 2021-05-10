using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Minerva.Report
{


    class ProjectPlan
    {
        private readonly ExcelDesolator<ProjectPlanItem> desolator;
        private List<ProjectPlanItem> projectPlanList;
        private List<ProjectPlanItem> ignoredProjectList;
        private List<ProjectPlanItem> existProjectList;
        private Summary summary;
        private ProjectName projectName;

        public ProjectPlan()
        {
            desolator = new ExcelDesolator<ProjectPlanItem>(Env.Instance.ProjectPlanPath);
            summary = new Summary();
            projectPlanList = desolator.SelectSheetAt(0)
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

        public ProjectPlan CompareWith(WorkReport report)
        {
            report.WorkReportList.ForEach(item => { item.Name = projectName.TryToProjectPlanName(item.Name); });

            ignoredProjectList = report.WorkReportList
                .Where(item => !IsExistInProjectPlanList(item.Name))
                .Select(item => ToNewProjectPlanItem(item))
                .ToList();

            existProjectList = report.WorkReportList
                .Where(item => IsExistInProjectPlanList(item.Name))
                .Select(item => ToProjectPlanItem(item))
                .ToList();

            return this;
        }

        private ProjectPlanItem ToProjectPlanItem(WorkReportItem workReportItem)
        {
            ProjectPlanItem item = projectPlanList.Find(projectPlanItem => projectPlanItem.ProjectName.Trim().Equals(workReportItem.Name));
            item.RemainHumanMonth = summary.ToRemainHumanMonth(workReportItem.Schedule);
            item.EstimatedTimeRemaining = summary.ToEstimatedTimeRemaining(workReportItem.Schedule);
            return item;
        }

        private ProjectPlanItem ToNewProjectPlanItem(WorkReportItem workReportItem)
        {
            ProjectPlanItem projectPlanItem = new ProjectPlanItem
            {
                ProjectName = workReportItem.Name,
                ResponsiblePersonnel = workReportItem.ResponsiblePersonnel,
                HostDivision = workReportItem.HostDivision,
                RequirementDepartment = workReportItem.BizDepartment,
                RemainHumanMonth = summary.ToRemainHumanMonth(workReportItem.Schedule),
                EstimatedTimeRemaining = summary.ToEstimatedTimeRemaining(workReportItem.Schedule)
            };
            return projectPlanItem;
        }


        private bool IsExistInProjectPlanList(string projectName)
        {
            return projectPlanList.Any(projectPlanItem => projectPlanItem.ProjectName.Trim().Equals(projectName));
        }

        public ProjectPlan ReNewProjectPlan()
        {
            desolator.SetCellValues(ignoredProjectList);
            desolator.Fit();
            desolator.SetCellValues(existProjectList);
            desolator.Fit();
            desolator.Save();
            return this;
        }
    }
}

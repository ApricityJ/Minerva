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
        private Summary summary = new Summary();

        public ProjectPlan()
        {
            desolator = new ExcelDesolator<ProjectPlanItem>(Env.Instance.ProjectPlanPath);
            projectPlanList = desolator.SelectSheetAt(0)
                .Skip(2)
                .ToEntityList(typeof(ProjectPlanItem));
        }

        public List<ProjectPlanItem> ToProjectPlanList()
        {
            return projectPlanList;
        }

        public ProjectPlan CompareWithWorkReport(WorkReport report)
        {
            ignoredProjectList = report.WorkReportList
                .Where(workReportItem => IsNotExistInProjectPlanList(workReportItem.Name))
                .Select(workReportItem => ToProjectPlanItem(workReportItem))
                .ToList();

            existProjectList = report.WorkReportList
                .Where(workReportItem => !IsNotExistInProjectPlanList(workReportItem.Name))
                .Select(workReportItem => ToProjectPlanItem(workReportItem))
                .ToList();

            return this;
        }

        private ProjectPlanItem ToProjectPlanItem(WorkReportItem workReportItem)
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


        private bool IsNotExistInProjectPlanList(string projectName)
        {
            return !projectPlanList.Any(projectPlanItem => projectPlanItem.ProjectName.Trim().Equals(projectName));
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

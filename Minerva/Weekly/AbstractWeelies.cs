using System.Collections.Generic;


namespace Minerva.Weekly
{


    using Department;

    abstract class AbstractWeelies
    {

        public Dictionary<InnerDepartment, BaseWeekly> WeeklyCollection { get; set; }


        public AbstractWeelies()
        {

        }


        //加载周报
        public abstract AbstractWeelies LoadWeekly();


       

        //汇总
        public abstract AbstractWeelies Summarize()
        {


           

            return this;
        }

        public AbstractWeelies Save()
        {
            ToTargetPath();
            ExcelDAO<WeeklyItem> dao = new ExcelDAO<WeeklyItem>(targetPath);
            dao.Save();
            return this;
        }



    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Minerva.Weekly
{
    using DAO;
    public class AbstractWeekly
    {

        public string FilePath { get; set; }

        ExcelDAO<WeeklyItem> dao;

        public List<WeeklyItem> Report { get; set; }
        public List<WeeklyItem> CurrentWeekRelease { get; set; }
        public List<WeeklyItem> NextWeekRelease { get; set; }
        public List<WeeklyItem> ExceptionReport { get; set; }

        public AbstractWeekly(string path)
        {
            FilePath = path;
            dao = new ExcelDAO<WeeklyItem>(path);

            //
            Report = dao.SelectSheetAt(0)
                .Skip(2)
                .ToEntityList(typeof(WeeklyItem));


            CurrentWeekRelease = Seperate("本周投产情况");
            NextWeekRelease = Seperate("下周投产计划");
            ExceptionReport = Seperate("生产异常");
        }

        public List<WeeklyItem> Seperate(string keyWord)
        {
            int index = dao.FindRowIndex(0, keyWord);
            if (index > -1)
            {
                return dao.SelectSheetAt(0)
                .Skip(index + 2)
                .ToEntityList(typeof(WeeklyItem));
            }
            else
            {
                return new List<WeeklyItem>();
            }
        }

        
    }
}

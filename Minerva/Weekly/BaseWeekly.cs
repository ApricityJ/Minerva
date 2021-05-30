using System.Collections.Generic;


namespace Minerva.Weekly
{
    using DAO;

    /// <summary>
    /// 基础周报类
    /// 根据目前的周报总结，一份周报至少包括 [工作周报] 与 [生产异常] 两部分内容
    /// (测试部不肯：我怎么会有生产异常？)
    /// 开发1，2，3，4部需要追加本周与下周投产两部分
    /// </summary>
    public class BaseWeekly
    {

        public string FilePath { get; set; }
        public ExcelDAO<WeeklyItem> dao;
        public List<WeeklyItem> CurrentWeekWork { get; set; }
        public List<WeeklyItem> UnnormalCase { get; set; }

        public BaseWeekly(string path)
        {
            FilePath = path;
            dao = new ExcelDAO<WeeklyItem>(path);

            CurrentWeekWork = ExtractBy("工作周报");
            UnnormalCase = ExtractBy("生产异常");
        }


        //根据标题抽取一个区域的内容
        //e.g
        //在一份周报中，按规定必须存在[生产异常]为标题的一个区域的内容
        //ExtractBy("生产异常") -> List<WeeklyItem> 包含该部分内容的List
        public List<WeeklyItem> ExtractBy(string keyWord)
        {
            int index = 0;
            if (!keyWord.Equals("工作周报"))
            {
                index = dao.FindRowIndex(0, keyWord);
            }
             
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

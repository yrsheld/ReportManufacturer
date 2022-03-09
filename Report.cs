using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.RightsManagement;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp14
{
    public class Report  //the main structure of the report
    {
        #region Table1
        public string Author;
        public string Date;
        public string Week;

        public ShortTermPlan TermPlan; 
        public WeekSummary Summary;
        #endregion Table1

        #region Table2
        public List<WeekDetail> Detail;
        public string AllDetail;
        #endregion Table2

        #region Table3
        public NextWeekPlan NextPlan;
        #endregion Table3

        #region Table4
        public FeedBack FeedBack;
        #endregion Table4

        public Report()
        {
            TermPlan = new ShortTermPlan();
            Summary = new WeekSummary();   //TODO
            Detail = new List<WeekDetail>();
            NextPlan = new NextWeekPlan();
            FeedBack = new FeedBack();
            AllDetail = "";
        }
    }

    #region Table1
    public partial class ShortTermPlan
    {
        public Dictionary<string, string> WorkNumber = new Dictionary<string, string>(){};

        public List<string> Plans;
        public string Str;

        public ShortTermPlan()
        {
            Plans = new List<string>();
            Str = "";
        }

        public int Add(string title)   //title from the atom feed
        {
            //Console.WriteLine(title);
            string result;
            if (title.Substring(0,"PICA".Length) == "PICA" || title.Substring(0,"CoreLibrary".Length)=="PICA")
            {
                result = $"[{ WorkNumber["PICAⅡ"]}] PICAⅡ";
            }
            else if (title.Substring(0, "EMMI".Length) == "EMMI")
            {
                result = $"[{ WorkNumber["EMMI"]}] EMMI";
            }
            else
            {
                result = $"[XXXXXXXXX] Other";
            }

            if (!Plans.Contains(result))  //would not add repeat title 
            { 
                Plans.Add(result);  //add the plan
            }
            
            return Plans.IndexOf(result);  
        }
        public int Count()   //return the number of plans
        {
            return Plans.Count();
        }
        public void Combine()  //concatenate all the strs in Plans 
        {
            foreach(var p in Plans)
            {
                Str += p + "\n";
            }
        }
    }
    public partial class WeekSummary
    {
        public List<string> Summarys;
        public string Str;

        public WeekSummary()
        {
            Summarys = new List<string>();
            Str = "";
        }
        public string ToSummary(string title)  
        {
            string target;

            //ToDo: Better way to get the title???
            int label = title.IndexOf("#");
            int comma = title.IndexOf(":");

            if (comma - label < 10)    //a temporal way of judgement 
            {
                target = title.Substring(comma+2);
                return target;
            }
            else
            {
                return "No title";
            }
        }
        public void Combine()  //concatenate all the strs in Plans 
        {
            foreach (var p in Summarys)
            {
                Str += p + "\n";
            }
        }
    }
    #endregion

    #region Table2
    public partial class WeekDetail 
    {
        public List<string> Details;  
        public string Str;

        public WeekDetail()
        {
            Details = new List<string>();
            Str = "";
        }
        public void Add(string original) 
        {
            Details.Add(original);   
        }
        
        public void Combine()  
        {
            foreach (var p in Details)
            {
                Str += p;
            }
        }

    }

    #endregion Table2


    #region Table3
    public partial class NextWeekPlan
    {
        private List<string> NextWeekPlans;
        public string Str;

        public NextWeekPlan()
        {
            NextWeekPlans = new List<string>();
            Str = "";
        }
    }
    #endregion Table3

    #region Table4
    public partial class FeedBack
    {
        private List<string> FeedBacks;
        public string Str;

        public FeedBack()
        {
            FeedBacks = new List<string>();
            Str = "";
        }
    } 
    #endregion Table4
}

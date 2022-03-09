using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using Spire;
using Spire.Doc;
using Spire.Doc.Interface;
using Spire.DocViewer;
using Spire.Doc.Fields;
using System.IO;
using System.IO.Packaging; 
using System.Runtime.Remoting.Messaging;
using System.Runtime.Serialization;
using System.Xml;
using System.Xml.Serialization;
using System.Text.RegularExpressions;
using Spire.Doc.Documents;
using System.Diagnostics;
using System.Net;

namespace WindowsFormsApp14
{
    public partial class Form1 : Form
    {
        Document myDocx = new Spire.Doc.Document(); 
        Object oFilePath;
        Object G_str_path;
        List<FeedData> AllFeedData;  //one FeedData for each URL (store the feeddata for this report)
        List<string> URLs;  //all the URLs
        List<Report> AllReports;
        
        Report MyReport;
        string[] Useless;   //all the titles in Issue Trackers 
        Form2 form2;
        Spire.Doc.Formatting.CharacterFormat myFont;

        public Form1()
        {
            InitializeComponent();
            URLs = new List<string>();
            AllFeedData = new List<FeedData>();
            AllReports = new List<Report>();
            MyReport = new Report();
            AllReports.Add(MyReport);

            oFilePath = "template.docx";
            //oFilePath = "D:\\template.docx";  //the file

            Useless = new string[]{"Due date", "Assignee", "Start date",
                "Tracker", "Priority", "Severity", "Parent task" , "File", "Story Point", "Status", "% Done",
                "Target version", "Tags", "Subject","Precedes", "Description", "Copied From", "Ideal estimated time", 
                "Worst estimated time", "Possible estimated time", "Estimated time", "Related to"
            };
            
        }

        public bool GetAtomData()   //Deserialize all URLs
        {
            AllFeedData.Clear(); //clear the previous one
            var errors = 0;
            foreach(var input in URLs)
            {
                try
                {
                    //1. If it is a Weekly Atom, decompose each atom, then append all of them to the URL pool
                    if (input.Contains("issues.atom?"))
                    {
                        FeedData WeekAtom = new FeedData(input);
                        string url;
                        FeedData SingleFeed;

                        if (WeekAtom.feed.entry == null) continue;

                        foreach (var EachFeed in WeekAtom.feed.entry)
                        {
                            url = EachFeed.linkField.hrefField + ".atom?key=a315f975290796878c381800437af6ca59473acf";
                            SingleFeed = new FeedData(url);
                            SingleFeed.feed.summary = AnalyzeContent(EachFeed.contentField.Text);
                            AllFeedData.Add(SingleFeed);
                        }
                    }
                    //2. If it is a single Atom, directly append to the URL pool (but then, it would not containt summary)
                    else
                    {
                        AllFeedData.Add(new FeedData(input));
                    }
                }
                catch(Exception e)
                {
                    MessageBox.Show("Invalid Atom:\n"+input,"Error", MessageBoxButtons.OK,MessageBoxIcon.Error);
                    errors = 1;
                    break;
                }
            }
            if (errors == 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public void LoadListStyle()
        {
            //set the new liststyle
            ListStyle listStyle0 = new ListStyle(myDocx, ListType.Numbered);
            listStyle0.Name = "listStyle0";
            listStyle0.Levels[0].PatternType = ListPatternType.Arabic;

            //1
            ListStyle listStyle1 = new ListStyle(myDocx, ListType.Numbered);
            listStyle1.Name = "listStyle1";
            listStyle1.Levels[0].PatternType = ListPatternType.Arabic;
            listStyle1.Levels[1].PatternType = ListPatternType.Arabic;
            listStyle1.Levels[1].NumberPrefix = "1.\x0000.";
            //2
            ListStyle listStyle2 = new ListStyle(myDocx, ListType.Numbered);
            listStyle2.Name = "listStyle2";
            listStyle2.Levels[0].PatternType = ListPatternType.Arabic;
            listStyle2.Levels[1].PatternType = ListPatternType.Arabic;
            listStyle2.Levels[1].NumberPrefix = "2.\x0000.";
            //3
            ListStyle listStyle3 = new ListStyle(myDocx, ListType.Numbered);
            listStyle3.Name = "listStyle3";
            listStyle3.Levels[0].PatternType = ListPatternType.Arabic;
            listStyle3.Levels[1].PatternType = ListPatternType.Arabic;
            listStyle3.Levels[1].NumberPrefix = "3.\x0000.";
            //4
            ListStyle listStyle4 = new ListStyle(myDocx, ListType.Numbered);
            listStyle4.Name = "listStyle4";
            listStyle4.Levels[0].PatternType = ListPatternType.Arabic;
            listStyle4.Levels[1].PatternType = ListPatternType.Arabic;
            listStyle4.Levels[1].NumberPrefix = "4.\x0000.";
            //5
            ListStyle listStyle5 = new ListStyle(myDocx, ListType.Numbered);
            listStyle5.Name = "listStyle5";
            listStyle5.Levels[0].PatternType = ListPatternType.Arabic;
            listStyle5.Levels[1].PatternType = ListPatternType.Arabic;
            listStyle5.Levels[1].NumberPrefix = "5.\x0000.";
            //6
            ListStyle listStyle6 = new ListStyle(myDocx, ListType.Numbered);
            listStyle6.Name = "listStyle6";
            listStyle6.Levels[0].PatternType = ListPatternType.Arabic;
            listStyle6.Levels[1].PatternType = ListPatternType.Arabic;
            listStyle6.Levels[1].NumberPrefix = "6.\x0000.";
            //7
            ListStyle listStyle7 = new ListStyle(myDocx, ListType.Numbered);
            listStyle7.Name = "listStyle7";
            listStyle7.Levels[0].PatternType = ListPatternType.Arabic;
            listStyle7.Levels[1].PatternType = ListPatternType.Arabic;
            listStyle7.Levels[1].NumberPrefix = "7.\x0000.";
            //8
            ListStyle listStyle8 = new ListStyle(myDocx, ListType.Numbered);
            listStyle8.Name = "listStyle8";
            listStyle8.Levels[0].PatternType = ListPatternType.Arabic;
            listStyle8.Levels[1].PatternType = ListPatternType.Arabic;
            listStyle8.Levels[1].NumberPrefix = "8.\x0000.";
            //9
            ListStyle listStyle9 = new ListStyle(myDocx, ListType.Numbered);
            listStyle9.Name = "listStyle9";
            listStyle9.Levels[0].PatternType = ListPatternType.Arabic;
            listStyle9.Levels[1].PatternType = ListPatternType.Arabic;
            listStyle9.Levels[1].NumberPrefix = "9.\x0000.";
            //10
            ListStyle listStyle10 = new ListStyle(myDocx, ListType.Numbered);
            listStyle10.Name = "listStyle10";
            listStyle10.Levels[0].PatternType = ListPatternType.Arabic;
            listStyle10.Levels[1].PatternType = ListPatternType.Arabic;
            listStyle10.Levels[1].NumberPrefix = "10.\x0000.";
            //11
            ListStyle listStyle11 = new ListStyle(myDocx, ListType.Numbered);
            listStyle11.Name = "listStyle11";
            listStyle11.Levels[0].PatternType = ListPatternType.Arabic;
            listStyle11.Levels[1].PatternType = ListPatternType.Arabic;
            listStyle11.Levels[1].NumberPrefix = "11.\x0000.";
            //12
            ListStyle listStyle12 = new ListStyle(myDocx, ListType.Numbered);
            listStyle12.Name = "listStyle12";
            listStyle12.Levels[0].PatternType = ListPatternType.Arabic;
            listStyle12.Levels[1].PatternType = ListPatternType.Arabic;
            listStyle12.Levels[1].NumberPrefix = "12.\x0000.";
            //13
            ListStyle listStyle13 = new ListStyle(myDocx, ListType.Numbered);
            listStyle13.Name = "listStyle13";
            listStyle13.Levels[0].PatternType = ListPatternType.Arabic;
            listStyle13.Levels[1].PatternType = ListPatternType.Arabic;
            listStyle13.Levels[1].NumberPrefix = "13.\x0000.";
            //14
            ListStyle listStyle14 = new ListStyle(myDocx, ListType.Numbered);
            listStyle14.Name = "listStyle14";
            listStyle14.Levels[0].PatternType = ListPatternType.Arabic;
            listStyle14.Levels[1].PatternType = ListPatternType.Arabic;
            listStyle14.Levels[1].NumberPrefix = "14.\x0000.";
            //15
            ListStyle listStyle15 = new ListStyle(myDocx, ListType.Numbered);
            listStyle15.Name = "listStyle15";
            listStyle15.Levels[0].PatternType = ListPatternType.Arabic;
            listStyle15.Levels[1].PatternType = ListPatternType.Arabic;
            listStyle15.Levels[1].NumberPrefix = "15.\x0000.";
            //16
            ListStyle listStyle16 = new ListStyle(myDocx, ListType.Numbered);
            listStyle16.Name = "listStyle16";
            listStyle16.Levels[0].PatternType = ListPatternType.Arabic;
            listStyle16.Levels[1].PatternType = ListPatternType.Arabic;
            listStyle16.Levels[1].NumberPrefix = "16.\x0000.";
            //17
            ListStyle listStyle17 = new ListStyle(myDocx, ListType.Numbered);
            listStyle17.Name = "listStyle17";
            listStyle17.Levels[0].PatternType = ListPatternType.Arabic;
            listStyle17.Levels[1].PatternType = ListPatternType.Arabic;
            listStyle17.Levels[1].NumberPrefix = "17.\x0000.";
            //18
            ListStyle listStyle18 = new ListStyle(myDocx, ListType.Numbered);
            listStyle18.Name = "listStyle18";
            listStyle18.Levels[0].PatternType = ListPatternType.Arabic;
            listStyle18.Levels[1].PatternType = ListPatternType.Arabic;
            listStyle18.Levels[1].NumberPrefix = "18.\x0000.";
            //19
            ListStyle listStyle19 = new ListStyle(myDocx, ListType.Numbered);
            listStyle19.Name = "listStyle19";
            listStyle19.Levels[0].PatternType = ListPatternType.Arabic;
            listStyle19.Levels[1].PatternType = ListPatternType.Arabic;
            listStyle19.Levels[1].NumberPrefix = "19.\x0000.";

            //2-1
            ListStyle listStyle101 = new ListStyle(myDocx, ListType.Numbered);
            listStyle101.Name = "listStyle2-1";
            listStyle101.Levels[0].PatternType = ListPatternType.Arabic;
            listStyle101.Levels[1].PatternType = ListPatternType.Arabic;
            listStyle101.Levels[1].NumberPrefix = "1.\x0000.";
            //2-2
            ListStyle listStyle102 = new ListStyle(myDocx, ListType.Numbered);
            listStyle102.Name = "listStyle2-2";
            listStyle102.Levels[0].PatternType = ListPatternType.Arabic;
            listStyle102.Levels[1].PatternType = ListPatternType.Arabic;
            listStyle102.Levels[1].NumberPrefix = "2.\x0000.";
            //2-3
            ListStyle listStyle103 = new ListStyle(myDocx, ListType.Numbered);
            listStyle103.Name = "listStyle2-3";
            listStyle103.Levels[0].PatternType = ListPatternType.Arabic;
            listStyle103.Levels[1].PatternType = ListPatternType.Arabic;
            listStyle103.Levels[1].NumberPrefix = "3.\x0000.";
            //2-4
            ListStyle listStyle104 = new ListStyle(myDocx, ListType.Numbered);
            listStyle104.Name = "listStyle2-4";
            listStyle104.Levels[0].PatternType = ListPatternType.Arabic;
            listStyle104.Levels[1].PatternType = ListPatternType.Arabic;
            listStyle104.Levels[1].NumberPrefix = "4.\x0000.";
            //2-5
            ListStyle listStyle105 = new ListStyle(myDocx, ListType.Numbered);
            listStyle105.Name = "listStyle2-5";
            listStyle105.Levels[0].PatternType = ListPatternType.Arabic;
            listStyle105.Levels[1].PatternType = ListPatternType.Arabic;
            listStyle105.Levels[1].NumberPrefix = "5.\x0000.";
            //2-6
            ListStyle listStyle106 = new ListStyle(myDocx, ListType.Numbered);
            listStyle106.Name = "listStyle2-6";
            listStyle106.Levels[0].PatternType = ListPatternType.Arabic;
            listStyle106.Levels[1].PatternType = ListPatternType.Arabic;
            listStyle106.Levels[1].NumberPrefix = "6.\x0000.";
            //2-7
            ListStyle listStyle107 = new ListStyle(myDocx, ListType.Numbered);
            listStyle107.Name = "listStyle2-7";
            listStyle107.Levels[0].PatternType = ListPatternType.Arabic;
            listStyle107.Levels[1].PatternType = ListPatternType.Arabic;
            listStyle107.Levels[1].NumberPrefix = "7.\x0000.";
            //2-8
            ListStyle listStyle108 = new ListStyle(myDocx, ListType.Numbered);
            listStyle108.Name = "listStyle2-8";
            listStyle108.Levels[0].PatternType = ListPatternType.Arabic;
            listStyle108.Levels[1].PatternType = ListPatternType.Arabic;
            listStyle108.Levels[1].NumberPrefix = "8.\x0000.";
            //2-9
            ListStyle listStyle109 = new ListStyle(myDocx, ListType.Numbered);
            listStyle109.Name = "listStyle2-9";
            listStyle109.Levels[0].PatternType = ListPatternType.Arabic;
            listStyle109.Levels[1].PatternType = ListPatternType.Arabic;
            listStyle109.Levels[1].NumberPrefix = "9.\x0000.";
            //2-10
            ListStyle listStyle110 = new ListStyle(myDocx, ListType.Numbered);
            listStyle110.Name = "listStyle2-10";
            listStyle110.Levels[0].PatternType = ListPatternType.Arabic;
            listStyle110.Levels[1].PatternType = ListPatternType.Arabic;
            listStyle110.Levels[1].NumberPrefix = "10.\x0000.";
            //2-11
            ListStyle listStyle111 = new ListStyle(myDocx, ListType.Numbered);
            listStyle111.Name = "listStyle2-11";
            listStyle111.Levels[0].PatternType = ListPatternType.Arabic;
            listStyle111.Levels[1].PatternType = ListPatternType.Arabic;
            listStyle111.Levels[1].NumberPrefix = "11.\x0000.";
            //2-12
            ListStyle listStyle112 = new ListStyle(myDocx, ListType.Numbered);
            listStyle112.Name = "listStyle2-12";
            listStyle112.Levels[0].PatternType = ListPatternType.Arabic;
            listStyle112.Levels[1].PatternType = ListPatternType.Arabic;
            listStyle112.Levels[1].NumberPrefix = "12.\x0000.";
            //2-13
            ListStyle listStyle113 = new ListStyle(myDocx, ListType.Numbered);
            listStyle113.Name = "listStyle2-13";
            listStyle113.Levels[0].PatternType = ListPatternType.Arabic;
            listStyle113.Levels[1].PatternType = ListPatternType.Arabic;
            listStyle113.Levels[1].NumberPrefix = "13.\x0000.";
            //2-14
            ListStyle listStyle114 = new ListStyle(myDocx, ListType.Numbered);
            listStyle114.Name = "listStyle2-14";
            listStyle114.Levels[0].PatternType = ListPatternType.Arabic;
            listStyle114.Levels[1].PatternType = ListPatternType.Arabic;
            listStyle114.Levels[1].NumberPrefix = "14.\x0000.";
            //2-15
            ListStyle listStyle115 = new ListStyle(myDocx, ListType.Numbered);
            listStyle115.Name = "listStyle2-15";
            listStyle115.Levels[0].PatternType = ListPatternType.Arabic;
            listStyle115.Levels[1].PatternType = ListPatternType.Arabic;
            listStyle115.Levels[1].NumberPrefix = "15.\x0000.";
            //2-16
            ListStyle listStyle116 = new ListStyle(myDocx, ListType.Numbered);
            listStyle116.Name = "listStyle2-16";
            listStyle116.Levels[0].PatternType = ListPatternType.Arabic;
            listStyle116.Levels[1].PatternType = ListPatternType.Arabic;
            listStyle116.Levels[1].NumberPrefix = "16.\x0000.";
            //2-17
            ListStyle listStyle117 = new ListStyle(myDocx, ListType.Numbered);
            listStyle117.Name = "listStyle2-17";
            listStyle117.Levels[0].PatternType = ListPatternType.Arabic;
            listStyle117.Levels[1].PatternType = ListPatternType.Arabic;
            listStyle117.Levels[1].NumberPrefix = "17.\x0000.";
            //2-18
            ListStyle listStyle118 = new ListStyle(myDocx, ListType.Numbered);
            listStyle118.Name = "listStyle2-18";
            listStyle118.Levels[0].PatternType = ListPatternType.Arabic;
            listStyle118.Levels[1].PatternType = ListPatternType.Arabic;
            listStyle118.Levels[1].NumberPrefix = "18.\x0000.";
            //2-19
            ListStyle listStyle119 = new ListStyle(myDocx, ListType.Numbered);
            listStyle119.Name = "listStyle2-19";
            listStyle119.Levels[0].PatternType = ListPatternType.Arabic;
            listStyle119.Levels[1].PatternType = ListPatternType.Arabic;
            listStyle119.Levels[1].NumberPrefix = "19.\x0000.";

            //add to the doc
            myDocx.ListStyles.Add(listStyle0); myDocx.ListStyles.Add(listStyle1);
            myDocx.ListStyles.Add(listStyle2); myDocx.ListStyles.Add(listStyle3);
            myDocx.ListStyles.Add(listStyle4); myDocx.ListStyles.Add(listStyle5);
            myDocx.ListStyles.Add(listStyle6); myDocx.ListStyles.Add(listStyle7);
            myDocx.ListStyles.Add(listStyle8); myDocx.ListStyles.Add(listStyle9);
            myDocx.ListStyles.Add(listStyle10);
            myDocx.ListStyles.Add(listStyle11); myDocx.ListStyles.Add(listStyle12);
            myDocx.ListStyles.Add(listStyle13); myDocx.ListStyles.Add(listStyle14); 
            myDocx.ListStyles.Add(listStyle15); myDocx.ListStyles.Add(listStyle16); 
            myDocx.ListStyles.Add(listStyle17); myDocx.ListStyles.Add(listStyle18); 
            myDocx.ListStyles.Add(listStyle19);

            myDocx.ListStyles.Add(listStyle101);
            myDocx.ListStyles.Add(listStyle102); myDocx.ListStyles.Add(listStyle103);
            myDocx.ListStyles.Add(listStyle104); myDocx.ListStyles.Add(listStyle105);
            myDocx.ListStyles.Add(listStyle106); myDocx.ListStyles.Add(listStyle107);
            myDocx.ListStyles.Add(listStyle108); myDocx.ListStyles.Add(listStyle109);
            myDocx.ListStyles.Add(listStyle110);
            myDocx.ListStyles.Add(listStyle111); myDocx.ListStyles.Add(listStyle112);
            myDocx.ListStyles.Add(listStyle113); myDocx.ListStyles.Add(listStyle114);
            myDocx.ListStyles.Add(listStyle115); myDocx.ListStyles.Add(listStyle116);
            myDocx.ListStyles.Add(listStyle117); myDocx.ListStyles.Add(listStyle118);
            myDocx.ListStyles.Add(listStyle119);
        }
        public void LoadFont()
        {
            myFont = new Spire.Doc.Formatting.CharacterFormat(myDocx);
            myFont.FontName = "微軟正黑體";
            myFont.FontSize = 12;
        }

        public void ConvertToReport()  //Convert AllFeedData to Report
        {
            //load in all the 項目編號清單, font
            LoadListStyle();
            LoadFont();

            ConvertTitle();

            Section section0 = myDocx.Sections[0];
            ITable myTable1 = section0.Tables[0];
            ITable myTable2 = section0.Tables[1];

            Paragraph parTermPlan = myTable1.Rows[2].Cells[0].Paragraphs[0];
            while(myTable1.Rows[4].Cells[0].Paragraphs.Count > 0)
            {
                myTable1.Rows[4].Cells[0].Paragraphs.Remove(myTable1.Rows[4].Cells[0].Paragraphs[0]);
            }
            Paragraph parSummary = null;
            Paragraph parDetail = null;
            Paragraph parEntry = null;
            
            while(myTable2.Rows[1].Cells[0].Paragraphs.Count>0)
            {
                myTable2.Rows[1].Cells[0].Paragraphs.Remove(myTable2.Rows[1].Cells[0].Paragraphs[0]);
            }

            //start to fill in data
            parTermPlan.ListFormat.ListLevelNumber = 0; //the first level
            parTermPlan.ListFormat.ApplyStyle("listStyle0"); //levelStyle  
            parTermPlan.Format.LeftIndent = 20;


            foreach (var f in AllFeedData)
            {
                ConvertTermPlan(f.feed);
                ConvertDetail(f.feed);
            }                                   

            //see through all the f according to its level1Index
            for (int levelIndex = 1; levelIndex<=MyReport.TermPlan.Count();levelIndex++)
            {
                if(levelIndex==1)
                    parTermPlan.AppendText(MyReport.TermPlan.Plans[levelIndex-1]).ApplyCharacterFormat(myFont);
                else
                    parTermPlan.AppendText("\n"+MyReport.TermPlan.Plans[levelIndex - 1]).ApplyCharacterFormat(myFont);
                
                parSummary = myTable1.Rows[4].Cells[0].AddParagraph();
                parSummary.ListFormat.ListLevelNumber = 0; //the first level  
                parSummary.ListFormat.ApplyStyle("listStyle" + levelIndex.ToString()); //levelStyle 
                parSummary.ListFormat.CurrentListLevel.NumberPrefix = levelIndex.ToString() + ".";
                parSummary.ListFormat.CurrentListLevel.StartAt = 1;
                parSummary.ListFormat.IsRestartNumbering = true;
                parSummary.Format.LeftIndent = 20;
                

                string summarytext = "";

                foreach (var f in AllFeedData)
                {
                    if (f.feed.lv1idx == levelIndex)
                    {
                        summarytext = ConvertSummary(f.feed);

                        //Write title
                        if (parSummary.Text == "")
                        {
                            parSummary.AppendText(summarytext).ApplyCharacterFormat(myFont);
                        }
                        else
                        {
                            parSummary.AppendText("\n" + summarytext).ApplyCharacterFormat(myFont);
                        }

                        //for details
                        //Write issue title
                        parDetail = myTable2.Rows[1].Cells[0].AddParagraph();
                        parDetail.ListFormat.ListLevelNumber = 0; //the first level     
                        parDetail.ListFormat.ApplyStyle("listStyle2-" + levelIndex.ToString()); //levelStyle 
                        parDetail.ListFormat.CurrentListLevel.NumberPrefix = levelIndex.ToString() + ".";
                        parDetail.AppendText(summarytext).ApplyCharacterFormat(myFont);
                        parDetail.Format.LeftIndent = 20;


                        ConvertDetail(f.feed);   //create a WeekDetail, push in MyReport.WeekDetail

                        //Write issue Summary (only for Weekly Report)
                        if (!string.IsNullOrEmpty(f.feed.summary))
                        {
                            string[] summaries = f.feed.summary.Split(new string[] { "<table>", "</table>"}, StringSplitOptions.RemoveEmptyEntries);
                            foreach (var s in summaries)
                            {
                                if (s.Contains("<tr>"))
                                {
                                    CreateTableHere(s,20);
                                }
                                else if (s.Contains("<unolist>"))
                                {
                                    string[] untitle1 = s.Split(new string[] { "<unolist>" }, StringSplitOptions.RemoveEmptyEntries);

                                    foreach (var u in untitle1)  
                                    {
                                        string[] unlist1 = u.Split(new string[] { "<title>" }, StringSplitOptions.RemoveEmptyEntries);
                                        
                                            //Wrtie Entry Title (as level2 list)
                                            parEntry = myTable2.Rows[1].Cells[0].AddParagraph();
                                            parEntry.ListFormat.ListLevelNumber = 1; //the first level     
                                            parEntry.ListFormat.ApplyStyle("listStyle2-" + levelIndex.ToString()); //levelStyle 

                                            if (unlist1.Length > 1)
                                            {
                                                parEntry.AppendText(unlist1[0]).ApplyCharacterFormat(myFont);
                                            }
                                            else
                                            {   
                                                parEntry.AppendText(" ").ApplyCharacterFormat(myFont);
                                            }

                                            //Console.WriteLine("listtitle=" + unlist1[0] + "*************");
                                            parEntry.Format.LeftIndent = 70;

                                            //Write Entry Detail (as Paragraph)
                                            parEntry = myTable2.Rows[1].Cells[0].AddParagraph();
                                            if (unlist1.Length > 1)
                                            {
                                                parEntry.AppendText(unlist1[1]).ApplyCharacterFormat(myFont);
                                            }
                                            else
                                            {   
                                                parEntry.AppendText(unlist1[0]).ApplyCharacterFormat(myFont);
                                            }

                                            parEntry.Format.LeftIndent = 70;

                                    }
                                }
                                else
                                {
                                    //Write Entry Detail (as Paragraph)
                                    parEntry = myTable2.Rows[1].Cells[0].AddParagraph();
                                    parEntry.AppendText(s).ApplyCharacterFormat(myFont);
                                    parEntry.Format.LeftIndent = 20;
                                }
                            }
                        }

                        //Write Details
                        foreach (var detail in MyReport.Detail.Last().Details) //one detail, one entry 
                        {
                            string[] details = detail.Split(new string[] { "<table>", "</table>" }, StringSplitOptions.RemoveEmptyEntries);
                            
                            foreach (var d in details)  //d: one entry
                            {
                                
                                if (d.Contains("<tr>"))
                                {
                                    CreateTableHere(d,70);
                                }
                                else if(d.Contains("<unolist>")) 
                                {
                                    string[] untitle = d.Split(new string[] { "<unolist>" }, StringSplitOptions.RemoveEmptyEntries);
            
                                    foreach (var u in untitle) 
                                    {
                                        string[] unlist = u.Split(new string[] { "<title>" }, StringSplitOptions.RemoveEmptyEntries);
                                        
                                            //Wrtie Entry Title (as level2 list)
                                            parEntry = myTable2.Rows[1].Cells[0].AddParagraph();
                                            parEntry.ListFormat.ListLevelNumber = 1; //the first level     
                                            parEntry.ListFormat.ApplyStyle("listStyle2-" + levelIndex.ToString()); //levelStyle 
                                            
                                            if (unlist.Length > 1) {
                                                parEntry.AppendText(unlist[0]).ApplyCharacterFormat(myFont);
                                            }
                                            else
                                            {   
                                                parEntry.AppendText(" ").ApplyCharacterFormat(myFont);
                                            }
                                           
                                            Console.WriteLine("listtitle="+ unlist[0] +"*************");
                                            parEntry.Format.LeftIndent = 70;

                                            //Write Entry Detail (as Paragraph)
                                            parEntry = myTable2.Rows[1].Cells[0].AddParagraph();
                                            if (unlist.Length > 1)
                                            {
                                                parEntry.AppendText(unlist[1]).ApplyCharacterFormat(myFont);
                                            }
                                            else
                                            {   
                                                parEntry.AppendText(unlist[0]).ApplyCharacterFormat(myFont);
                                            }
                                            
                                            parEntry.Format.LeftIndent = 70;

                                    }
                                }
                                else
                                {
                                    parEntry = myTable2.Rows[1].Cells[0].AddParagraph();
                                    parEntry.ListFormat.ListLevelNumber = 1; //the first level     
                                    parEntry.ListFormat.ApplyStyle("listStyle2-" + levelIndex.ToString()); //levelStyle 
                                    
                                    parEntry.AppendText(" ").ApplyCharacterFormat(myFont);
                                    parEntry.Format.LeftIndent = 70;

                                    //Write Entry Detail (as Paragraph)
                                    parEntry = myTable2.Rows[1].Cells[0].AddParagraph();
                                    parEntry.AppendText(d).ApplyCharacterFormat(myFont);
                                    parEntry.Format.LeftIndent = 70;

                                }
                                
                            }
                           
                        }
                    }
                }
            }
        }
        private void CreateTableHere(string s, int indent) 
        {
            Table table = myDocx.Sections[0].Tables[1].Rows[1].Cells[0].AddTable();
            int rows = Regex.Matches(s, "<tr>").Count;
            int cols = Regex.Matches(s, "<td>").Count / rows;
            table.TableFormat.LeftIndent = indent;
            table.ResetCells(rows, cols);
            table.TableFormat.Borders.BorderType = Spire.Doc.Documents.BorderStyle.Hairline;

            string[] tablerow = s.Split(new string[] { "<tr>" }, StringSplitOptions.RemoveEmptyEntries);
            List<string> tabledata = new List<string>();

            foreach (var r in tablerow)
            {
                var newr = r.Trim();
                if (string.IsNullOrEmpty(newr)) continue;

                newr = newr.Substring(0, newr.IndexOf("</tr>")).Trim();
                string[] d = newr.Split(new string[] { "<td>" }, StringSplitOptions.RemoveEmptyEntries);
                tabledata.AddRange(d);
            }
            int k = 0;
            string realdata;


            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    realdata = tabledata[k];
                    int endtag = realdata.IndexOf("</td>");
                    realdata = realdata.Substring(0, endtag);
                    
                    TextRange range = table[i, j].AddParagraph().AppendText(realdata);
                    k++;
                    range.ApplyCharacterFormat(myFont);

                }
            }

        }
        public void ConvertTitle()
        {
            Section section1 = myDocx.Sections[0];
            ITable myTable = section1.Tables[0];

            //Insert Title
            Paragraph name = myTable.Rows[2].Cells[0].Paragraphs[0];
            myTable.Rows[0].Cells[0].Paragraphs[0].AppendText(textBox1.Text);
            myTable.Rows[0].Cells[1].Paragraphs[0].AppendText(textBox2.Text);
            myTable.Rows[0].Cells[2].Paragraphs[0].AppendText(textBox3.Text);
        }

        public void ConvertTermPlan(Feed f)   //Create all the strings of Table1
        {
            string title = f.entry[0].title;   //only look at the first entry's title (cuz all the title should be the same)
            f.lv1idx = MyReport.TermPlan.Add(title) + 1;
        }
        public string ConvertSummary(Feed f)
        {
            string title = f.entry[0].title;
            return MyReport.Summary.ToSummary(title);    
        }
        
        public void ConvertDetail(Feed f)  //Deserialize the entry's content (xml) 
        {
            string currentEntry = "";   //each entry --> one Detail
            WeekDetail weekDetail = new WeekDetail();
            
            foreach (var e in f.entry)  //for each entry   
            {
                string result = e.contentField.Text.Trim();  //get the original entry's content
                Console.WriteLine("result="+ result);
                currentEntry = AnalyzeContent(result);

                if (!string.IsNullOrEmpty(currentEntry))
                {
                    weekDetail.Add(currentEntry);
                }
            }
            
            MyReport.Detail.Add(weekDetail);  
        }

        private string AnalyzeContent(string content)
        {
            string currentEntry = "";
            //Console.WriteLine("content:\n"+content);

            //1. Remove all the IssueTracker titles (ex: Dated time, Assignee...)
            foreach (string t in Useless)
            {
                content = Regex.Replace(content, $@"<li><strong>{t}.+</li>", "");
            }

            //2. Take off all the bold   
            //content = Regex.Replace(content, $@"</?strong>", "");

            //3. Take off all the linkfield
            content = Regex.Replace(content, $" <a .*\">", "");
            content = Regex.Replace(content, $@"</a>", "");

            //4. Take off the &nbsp;
            content = Regex.Replace(content, "&nbsp;", "");

            //5. Take off the color span
            content = Regex.Replace(content,"<span style=(.*?)>","");
            content = Regex.Replace(content, "</span>", "");

            content = Regex.Replace(content, "&#38; ", "&");
            content = Regex.Replace(content, "</p>","<pend>");
            //5. Categorize by unordered list & paragraphs 
            string[] series = content.Split(new string[] { "<ul>", "</ul>", "<p>"}, StringSplitOptions.RemoveEmptyEntries);
            
            
            int index = 0, add = 0; string ss, input="";

            //6. for each list or paragraphs, find its listitem, and put into MyDetail
            foreach (var a in series)
            {
                string cur = a;

                //unorderedlist 
                //TODO:(need to find better way to label)
                cur = Regex.Replace(cur, @"<li><strong>(.*?)</strong><br />", "<unolist><title>$1<title>");
                cur = Regex.Replace(cur, @"<li><strong>(.*?)</strong>", "<unolist><title>$1<title>");
                //Console.WriteLine("\nAfter**************************\n"+cur);

                //cur = Regex.Replace(cur, "<li><strong>", "<unolist>");

                cur = Regex.Replace(cur, $@"</?strong>", "");

                string[] lines = cur.Split(new string[] { "<li>", "</li>" }, StringSplitOptions.RemoveEmptyEntries);
                index = 0;
                foreach (var st in lines)
                {
                    input = "";
                    ss = st.Trim();
                    if (string.IsNullOrEmpty(ss)) continue;
                    ss = Regex.Replace(ss, "<pend>","\n");
                    index += add;

                    if (ss.Contains("<ol>"))      //start to label idx
                    {
                        add = 1;
                        //input = "<olist>";
                        continue;
                    }
                    else if (ss.Contains("</ol>"))     //end labeling idx
                    {
                        index = 0; add = 0;
                        continue;
                    }

                    //add to MyReport 
                    if (index > 0)
                    {
                        input += (index.ToString() + ". " + ss.Replace("<br />", "\n"))+"\n";
                    }
                    else
                    {
                        //input = ss;
                        input += (ss.Replace("<br />", "\n"));
                    }
                    currentEntry += input;
                }
            }
            return currentEntry;
        }
 
        private void button1_Click(object sender, EventArgs e)  //Create Report
        {
            //create a new Report
            if (AllReports.Count > 1)
            {
                AllReports.Add(new Report());
                MyReport = AllReports.Last();  
            }

            URLs.Clear();
            if (listBox1.Items.Count > 0)
            {
                foreach (var url in listBox1.Items)
                {
                    URLs.Add(url.ToString());  //add all the url in listbox in URLs
                }
            }

            //All URLs --> Feed
            if (GetAtomData())
            {

                ThreadPool.QueueUserWorkItem(
                    (P_temp) =>             //use lambda expression
                    {
                        try
                        {
                            myDocx.LoadFromFile(@oFilePath.ToString(), FileFormat.Docx);

                            ConvertToReport();

                            G_str_path = string.Format(@"{0}{1}", "D:\\", DateTime.Now.ToString("yyyyMMdd hh:mm:ss") + ".docx");

                            myDocx.SaveToFile(@G_str_path.ToString(), FileFormat.Docx);

                            this.Invoke(            
                            (MethodInvoker)(() =>          
                            {
                                MessageBox.Show("Successful！");
                            }));

                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    });
            }
        }

        private void Add_Click(object sender, EventArgs e)
        {
            if (inputURLbox.Text != null)
            {
                string input = inputURLbox.Text;
                if (!listBox1.Items.Contains(input))   //if not added yet
                {
                    listBox1.Items.Add(input);
                }
                else
                {
                    MessageBox.Show("Already added");
                }
                inputURLbox.Text = "";
            }
        }

        private void Remove_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedItems.Count > 0)
            {
                string checkURL = this.listBox1.SelectedItem.ToString();
                this.listBox1.Items.Remove(checkURL);
            }
        }

        private void Clear_Click(object sender, EventArgs e)
        {
            DialogResult myResult = MessageBox.Show("Clear all the URLs?","",MessageBoxButtons.OKCancel);
            if (myResult == DialogResult.OK)
            {
                listBox1.Items.Clear();
            }
        }

        private void btn_Display_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Word2007-2010 files (*.docx)|*.docx|All files (*.*)|*.*";
            dialog.Title = "Select docx file";
            dialog.Multiselect = false;
            dialog.InitialDirectory = System.IO.Path.GetFullPath(@"D:\\".ToString());
            
            DialogResult result = dialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                try
                {
                    form2 = new Form2();
                    form2.Visible = true;

                    form2.docDocumentViewer1.LoadFromFile(dialog.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            }
        }
}

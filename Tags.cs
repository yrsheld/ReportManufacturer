using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace WindowsFormsApp14
{
    //The structure of AtomFeed Data (xml)
    //However, it cannot be used due to some tags mismatching in AtomFeed

    [XmlRoot("Content")]
    public class Content   //stand for one entry's content  
    {
        private unorderedList[] ULfield;
        private orderedList[] OLfield;
        private paragraph[] PARAfield; 

        [XmlElement("ul")]
        public unorderedList[] UL
        {
            get
            {
                return this.ULfield;
            }
            set
            {
                this.ULfield = value;
            }
        }

        [XmlElement("ol")]
        public orderedList[] OL
        {
            get
            {
                return this.OLfield;
            }
            set
            {
                this.OLfield = value;
            }
        }

        [XmlElement("p")]
        public paragraph[] PARA 
        {
            get
            {
                return this.PARAfield;
            }
            set
            {
                this.PARAfield = value;
            }
        }
    }

    public partial class unorderedList
    {
        //[XmlElement("li")]  //list items
        //public string[] listitem { get; set; }
        [XmlText]
        public string Text { get; set; }
    }
    public partial class orderedList
    {
        [XmlElement("li")]
        public string[] listitem { get; set; }
    }
    public partial class paragraph
    {
        public string Text { get; set; }   
    }
}

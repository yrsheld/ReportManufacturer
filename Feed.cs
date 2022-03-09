using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Net;
using System.IO;
using System.Xml;
using System.Xml.Serialization;
using System.Xml.Schema;
using System.Text;

namespace WindowsFormsApp14
{
    [XmlRoot("feed", Namespace = "http://www.w3.org/2005/Atom")]
    public partial class Feed
    {
        public int lv1idx;
        public int lv2idx;
        public string summary;

        private object titleField;

        private feedLink[] linkField;

        private string idField;

        private string iconField;

        private System.DateTime updatedField;

        private feedAuthor authorField;

        private feedEntry[] entryField;

        /// <remarks/>
        [XmlElement("title")]
        public object title
        {
            get
            {
                return this.titleField;
            }
            set
            {
                this.titleField = value;
            }
        }

        /// <remarks/>
        [XmlElement("link")]
        public feedLink[] link
        {
            get
            {
                return this.linkField;
            }
            set
            {
                this.linkField = value;
            }
        }

        [XmlElement("id")]
        public string id
        {
            get
            {
                return this.idField;
            }
            set
            {
                this.idField = value;
            }
        }

        [XmlElement("icon")]
        public string icon
        {
            get
            {
                return this.iconField;
            }
            set
            {
                this.iconField = value;
            }
        }

        [XmlElement("updated")]
        public System.DateTime updated
        {
            get
            {
                return this.updatedField;
            }
            set
            {
                this.updatedField = value;
            }
        }

        [XmlElement("author")]
        public feedAuthor author
        {
            get
            {
                return this.authorField;
            }
            set
            {
                this.authorField = value;
            }
        }

        /// <remarks/>
        [XmlElement("entry")]
        public feedEntry[] entry
        {
            get
            {
                return this.entryField;
            }
            set
            {
                this.entryField = value;
            }
        }
    }


    public partial class feedLink
    {
        [XmlAttribute("rel")]
        public string rel { get; set; }
        [XmlAttribute("type")]
        public string type { get; set; }
        [XmlAttribute("href")]
        public string href { get; set; }
    }

    public partial class feedAuthor
    {
        [XmlElement("name")]
        public string name { get; set; }
    }


    public partial class feedEntry
    {
        [XmlElement("title")]
        public string title { get; set; }

        [XmlElement("link")]
        public feedEntryLink linkField { get; set; }

        [XmlElement("id")]
        public string id { get; set; }

        [XmlElement("updated")]
        public System.DateTime updatedField { get; set; }
        [XmlElement("author")]
        public feedEntryAuthor authorField { get; set; }
        [XmlElement("content")]
        public feedEntryContent contentField { get; set; }

    }

    public partial class feedEntryLink
    {
        [XmlAttribute("rel")]
        public string relField { get; set; }
        [XmlAttribute("href")]
        public string hrefField { get; set; }
    }


    public partial class feedEntryAuthor
    {
        [XmlElement("name")]
        public string name { get; set; }

    }

    public partial class feedEntryContent
    {
        [XmlAttribute("type")]
        public string Type { get; set; }

        [XmlText]
        public string Text { get; set; }
    }

    public class FeedData
    {
        public string address;
        public Feed feed;

        public FeedData() { }
        public FeedData(string url)  //create the feed at the same time
        {
            address = url;
            DeserializeObject();
        }
        public void DeserializeObject()
        {
            
                WebRequest myRequest = WebRequest.Create(@address);
                myRequest.Method = "GET";
                WebResponse myResponse = myRequest.GetResponse();
                StreamReader sr = new StreamReader(myResponse.GetResponseStream());

                //Convert to string
                string xmlstr = sr.ReadToEnd();
                sr.Close();
                myResponse.Close();

                //Convert string to xml
                XmlDocument xdoc = new XmlDocument();
                xdoc.LoadXml(xmlstr);

                XmlNodeReader reader = new XmlNodeReader(xdoc.DocumentElement);
                XmlSerializer ser = new XmlSerializer(typeof(Feed));


                feed = (Feed)ser.Deserialize(reader);  //feed has all the deserialized data
                reader.Close();
           
        }
    }
}


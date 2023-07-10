using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Xml.Linq;

namespace DocViewer
{
    class XmlReader
    {
        private String[] titles;
        private String[] texts;

        public String[] getTitles() { return titles; }
        public String[] getTexts()  { return texts; }

        public Boolean readXml(String fileName)
        {
            XDocument xml = null;

            try
            {
                xml = XDocument.Load(fileName);
            }
            catch (FileNotFoundException)
            {
                return false;
            }
            XElement table = xml.Element("document");
            var rows = table.Elements("chapter");
            
            int chapterNum = 0;
            foreach (XElement row in rows)
            {
                chapterNum++;
            }

            titles = new String[chapterNum];
            texts = new String[chapterNum];

            int i = 0;
            foreach (XElement row in rows)
            {

                XElement title = row.Element("title");
                titles[i] = title.Value;

                XElement text = row.Element("text");
                texts[i] = text.Value.Replace("EOS", " ");

                i++;
            }

            return true;
        }
        
    }
}

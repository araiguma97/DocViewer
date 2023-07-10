using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace DocViewer
{
    class XmlWriter
    {
        private String[] ids;
        private String[] titles;
        private String[] texts;
        private String[] terms;
        private String[] sims;

        public XmlWriter (String[] ids, String[] titles, String[] texts, String[] terms, String[] sims) {
            this.ids = ids;
            this.titles = titles;
            this.texts = texts;
            this.terms = terms;
            this.sims = sims;
        } 

        public void writeXml(String fileName) {

            XElement root = new XElement("document");
            for (int i = 0; i < ids.Length; i++)
            {
                XElement chapter = new XElement("chapter");
                XElement id = new XElement("no", ids[i]);
                chapter.Add(id);
                XElement title = new XElement("title", titles[i]);
                chapter.Add(title);
                XElement text = new XElement("text", texts[i]);
                chapter.Add(text);
                if (terms[i] != null)
                {
                    XElement term = new XElement("term", terms[i]);
                    chapter.Add(term);
                }
                if (sims[i] != null)
                {
                    XElement sim = new XElement("sim", sims[i]);
                    chapter.Add(sim);
                }
                root.Add(chapter);
            }

            root.Save(fileName);
        }
    }
}

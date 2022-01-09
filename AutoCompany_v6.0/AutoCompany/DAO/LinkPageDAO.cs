using AutoCompany.Actions;
using AutoCompany.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

namespace AutoCompany.DAO
{
    class LinkPageDAO
    {
        private XElement xDoc;
        private string FileName = @"..\..\Assets\Data\LinkPage.xml";
        public LinkPageDAO()
        {
            xDoc = XElement.Load(FileName);
        }
        public bool Add(LinkPage linkPage, int type)
        {
            string typeString = type == (int)App.LINKPAGE.LIST ? "List" : "Selected";
            var xmlTree = new XElement("Page",
                new XElement("Name", linkPage.Name),
                new XElement("Link", linkPage.Link)
            );
            xDoc.Element(typeString).Add(xmlTree);
            xDoc.Save(FileName);
            return true;
        }
        public bool Delete(LinkPage linkPage, int type)
        {
            string typeString = type == (int)App.LINKPAGE.LIST ? "List" : "Selected";
            foreach (XElement Vocabulary in xDoc.Elements(typeString).Elements("Page"))
            {
                if (Vocabulary.Element("Name").Value.ToUpper().ToString().Equals(linkPage.Name.ToUpper()))
                {
                    Vocabulary.Remove();
                    xDoc.Save(FileName);
                    return true;
                }
            }
            return false;
        }
        public List<LinkPage> ReadAll(int type)
        {
            List<LinkPage> linkPage = new List<LinkPage>();
            if (type == (int)App.LINKPAGE.LIST)
            {
                foreach (XElement linkpageXML in XElement.Load(FileName).Elements("List").Elements("Page"))
                {
                    linkPage.Add(new LinkPage()
                    {
                        Name = linkpageXML.Element("Name").Value.ToString(),
                        Link = linkpageXML.Element("Link").Value.ToString()
                    });
                }
                return linkPage.OrderBy(x => x.Name).ToList();
            }
            else
            {
                foreach (XElement linkpageXML in XElement.Load(FileName).Elements("Selected").Elements("Page"))
                {
                    linkPage.Add(new LinkPage()
                    {
                        Name = linkpageXML.Element("Name").Value.ToString(),
                        Link = linkpageXML.Element("Link").Value.ToString()
                    });
                }
                return linkPage;
            }
        }
        public LinkPage Read(LinkPage linkPage, int type)
        {
            string typeString = type == (int)App.LINKPAGE.LIST ? "List" : "Selected";
            foreach (XElement linkPageXML in XElement.Load(FileName).Elements(typeString).Elements("Page"))
            {
                if (linkPageXML.Element("Name").Value.ToString().Equals(linkPage.Name))
                {
                    return new LinkPage()
                    {
                        Name = linkPageXML.Element("Name").Value.ToString(),
                        Link = linkPageXML.Element("Link").Value.ToString()
                    };
                }
            }
            return new LinkPage();
        }
    }
}

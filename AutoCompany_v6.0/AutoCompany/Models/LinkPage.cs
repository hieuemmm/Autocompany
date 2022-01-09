using AutoCompany.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoCompany.Models
{
    public class LinkPage
    {
        public LinkPage()
        {
        }
        public LinkPage(string name)
        {
            Name = name;
        }
        public string Name = "";
        public string Link = "";
        public List<string> LinkChild = new List<string>();
        public List<Companny> Compannies = new List<Companny>();
    }
}

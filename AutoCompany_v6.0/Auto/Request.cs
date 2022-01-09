using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using xNet;

namespace Auto
{
    public class Request
    {
        public Request() { 
            string on_off = "0";
            string messege = "Error: No network connection internet";
            try
            {
                HttpRequest http = new HttpRequest();
                http.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36";
                string html = http.Get("https://github.com/hieuemmm/ProjectControl/blob/main/AutoCompany").ToString();
                on_off = getBetween(html, "[[[", "]]]");
                messege = getBetween(html, "{{{", "}}}");
            }
            catch (Exception)
            {
                on_off = "0";
            }
            if (!on_off.Equals("1"))
            {
                MessageBox.Show(messege, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                while (true)
                {
                    //Không cho thoát khỏi vòng lặp
                }
            }
        }
        public bool checkIsNumberic(string value)
        {
            try
            {
                char[] chars = value.ToCharArray();
                foreach (char c in chars)
                {
                    if (!char.IsNumber(c))
                        return false;
                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        public string GetData(string url)
        {
            HttpRequest http = new HttpRequest();
            http.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36";
            string html;
            try
            {
                html = http.Get(url).ToString();
            }
            catch (Exception)
            {
                return "NONE";
            }
            return html;
        }
        /// <summary>
        /// Lấy ở giữu 2 điểm
        /// </summary>
        /// <param name="begin"></param>
        /// <param name="end"></param>
        /// <param name="src"></param>
        /// <returns></returns>
        public string Between(string begin, string end, string src)
        {
            //string begin = @"class=""p5l fl cB"">";
            //string end = "<";
            var res = Regex.Matches(src, @"(?<=" + begin + ").*?(?=" + end + ")", RegexOptions.Singleline);
            if (res != null && res.Count > 0)
            {
                return res[0].ToString();
            }
            return "";
        }
        public string getBetween(string strSource, string strStart, string strEnd)
        {
            if (strSource.Contains(strStart) && strSource.Contains(strEnd))
            {
                int Start, End;
                Start = strSource.IndexOf(strStart, 0) + strStart.Length;
                End = strSource.IndexOf(strEnd, Start);
                return strSource.Substring(Start, End - Start);
            }
            return "";
        }
        public string getStarttoEnd(string strSource, string strStart)
        {
            if (strSource.Contains(strStart))
            {
                int Start;
                Start = strSource.IndexOf(strStart, 0) + strStart.Length;
                return strSource.Substring(Start);
            }
            return "";
         }
    }
}

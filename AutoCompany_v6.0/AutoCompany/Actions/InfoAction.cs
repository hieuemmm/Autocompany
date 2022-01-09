using AutoCompany.Model;
using AutoCompany.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using xNet;

namespace AutoCompany.Actions
{
    public class InfoAction : Auto.Request
    {
        public Companny GetInfoCompany(string LinkCompany)
        {
            Companny companny = new Companny();
            List<string> listcompany = new List<string>(); ;
            var html = GetData(LinkCompany);
            companny.MST = getBetween(html, "<title>", "</title>").Substring(0, 10);
            //Vùng Xanh đậm
            html = getBetween(html, @"<div class=""jumbotron"">", "<h4>Doanh nghiệp mới cập nhật:</h4>");
            companny.Name = getBetween(html, @"<span title=""", @"""");
            //thu hẹp phạm vi source vì bị trùng
            html = getStarttoEnd(html, "Địa chỉ:");
            companny.Address = html.Substring(1, html.IndexOf("<br/>") - 1);
            companny.RepresentativeName = getBetween(html, "Đại diện pháp luật: ", "<br/>");
            companny.LicenseDate = getBetween(html, "Ngày cấp giấy phép: ", "<br/>");
            string base64ImageSDT = getBetween(html, "data:image/png;base64,", @"""");
            SaveImageBMP(base64ImageSDT);
            companny.SDT = BMP_To_Text(base64ImageSDT);
            return companny;
        }
        public List<string> GetListCompany(LinkPage linkPage)
        {
            List<string> listcompany = new List<string>(); ;
            var html = GetData(linkPage.Link);
            if (html.Equals("NONE"))
            {
                return new List<string>();
            }
            //tách các ví dụ từ cấu trúc ul>li
            var begin = @"href=";
            var end = @"</a>";
            var res = Regex.Matches(html, @"(?<=" + begin + ").*?(?=" + end + ")", RegexOptions.Singleline);
            foreach (var item in res)
            {
                string IT = item.ToString();
                string IT_MST = IT.Substring(item.ToString().Length - 10);
                if (checkIsNumberic(IT_MST))
                {
                    listcompany.Add(item.ToString().Substring(1, item.ToString().Length - 13));
                }
            }
            return listcompany;
        }
        //SDT
        private string BMP_To_Text(String base64Image)
        {
            using (Bitmap bitmapsource = (Bitmap)Image.FromFile(FileAction.AbsolutePath(@"..\..\Assets\SDT\SDT.bmp")))
            {
                string[] searchs = { "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "-", "+", "-V2" };
                List<PointImage> pointArray = new List<PointImage>();
                foreach (string searchstring in searchs)
                {
                    using (Bitmap bitmapSearch = (Bitmap)Image.FromFile(FileAction.AbsolutePath(@"..\..\Assets\SDT\" + searchstring + ".bmp")))
                    {
                        List<Point> points = FindBitmapsEntry(bitmapsource, bitmapSearch);
                        foreach (Point p in points)
                        {
                            pointArray.Add(new PointImage
                            {
                                X = p.X,
                                number = searchstring
                            });
                        }
                    }
                }
                pointArray = pointArray.OrderBy(t => t.X).ToList();
                string result = "";
                foreach (PointImage pointImage in pointArray)
                {
                    result += pointImage.number;
                }
                //Bỏ tất cả SAU dấu trừ
                if (result.IndexOf("-") != -1)
                {
                    return result.Substring(0, result.IndexOf("-"));
                }
                return result;
            };
        }
        private void SaveImageBMP(String base64Image)
        {
            string filePath = FileAction.AbsolutePath(@"..\..\Assets\SDT\SDT.bmp");
            File.WriteAllBytes(filePath, Convert.FromBase64String(base64Image));
        }
        public static List<Point> FindBitmapsEntry(Bitmap sourceBitmap, Bitmap serchingBitmap)
        {
            #region Arguments check

            if (sourceBitmap == null || serchingBitmap == null)
                throw new ArgumentNullException();

            //if (sourceBitmap.PixelFormat != serchingBitmap.PixelFormat)
            //    throw new ArgumentException("Pixel formats arn't equal");

            if (sourceBitmap.Width < serchingBitmap.Width || sourceBitmap.Height < serchingBitmap.Height)
                throw new ArgumentException("Size of serchingBitmap bigger then sourceBitmap");

            #endregion

            var pixelFormatSize = Image.GetPixelFormatSize(sourceBitmap.PixelFormat) / 8;


            // Copy sourceBitmap to byte array
            var sourceBitmapData = sourceBitmap.LockBits(new Rectangle(0, 0, sourceBitmap.Width, sourceBitmap.Height),
                ImageLockMode.ReadOnly, sourceBitmap.PixelFormat);
            var sourceBitmapBytesLength = sourceBitmapData.Stride * sourceBitmap.Height;
            var sourceBytes = new byte[sourceBitmapBytesLength];
            Marshal.Copy(sourceBitmapData.Scan0, sourceBytes, 0, sourceBitmapBytesLength);
            sourceBitmap.UnlockBits(sourceBitmapData);

            // Copy serchingBitmap to byte array
            var serchingBitmapData =
                serchingBitmap.LockBits(new Rectangle(0, 0, serchingBitmap.Width, serchingBitmap.Height),
                    ImageLockMode.ReadOnly, serchingBitmap.PixelFormat);
            var serchingBitmapBytesLength = serchingBitmapData.Stride * serchingBitmap.Height;
            var serchingBytes = new byte[serchingBitmapBytesLength];
            Marshal.Copy(serchingBitmapData.Scan0, serchingBytes, 0, serchingBitmapBytesLength);
            serchingBitmap.UnlockBits(serchingBitmapData);

            var pointsList = new List<Point>();

            // Serching entries
            // minimazing serching zone
            // sourceBitmap.Height - serchingBitmap.Height + 1
            for (var mainY = 0; mainY < sourceBitmap.Height - serchingBitmap.Height + 1; mainY++)
            {
                var sourceY = mainY * sourceBitmapData.Stride;

                for (var mainX = 0; mainX < sourceBitmap.Width - serchingBitmap.Width + 1; mainX++)
                {// mainY & mainX - pixel coordinates of sourceBitmap
                 // sourceY + sourceX = pointer in array sourceBitmap bytes
                    var sourceX = mainX * pixelFormatSize;

                    var isEqual = true;
                    for (var c = 0; c < pixelFormatSize; c++)
                    {// through the bytes in pixel
                        if (sourceBytes[sourceX + sourceY + c] == serchingBytes[c])
                            continue;
                        isEqual = false;
                        break;
                    }

                    if (!isEqual) continue;

                    var isStop = false;

                    // find fist equalation and now we go deeper) 
                    for (var secY = 0; secY < serchingBitmap.Height; secY++)
                    {
                        var serchY = secY * serchingBitmapData.Stride;

                        var sourceSecY = (mainY + secY) * sourceBitmapData.Stride;

                        for (var secX = 0; secX < serchingBitmap.Width; secX++)
                        {// secX & secY - coordinates of serchingBitmap
                         // serchX + serchY = pointer in array serchingBitmap bytes

                            var serchX = secX * pixelFormatSize;

                            var sourceSecX = (mainX + secX) * pixelFormatSize;

                            for (var c = 0; c < pixelFormatSize; c++)
                            {// through the bytes in pixel
                                if (sourceBytes[sourceSecX + sourceSecY + c] == serchingBytes[serchX + serchY + c]) continue;

                                // not equal - abort iteration
                                isStop = true;
                                break;
                            }

                            if (isStop) break;
                        }

                        if (isStop) break;
                    }

                    if (!isStop)
                    {// serching bitmap is founded!!
                        pointsList.Add(new Point(mainX, mainY));
                    }
                }
            }

            return pointsList;
        }
    }
}

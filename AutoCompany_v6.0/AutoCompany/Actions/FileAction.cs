using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoCompany.Actions
{
    public class FileAction
    {
        /// <summary>
        /// Xóa hết tất cả file trong folder
        /// </summary>
        /// <param name="directory">@"C:\..."</param>
        public static void ClearFolder(DirectoryInfo directory)
        {
            foreach (FileInfo file in directory.GetFiles()) file.Delete();
            foreach (DirectoryInfo subDirectory in directory.GetDirectories()) subDirectory.Delete(true);
        }
        /// <summary>
        /// FileAction.AbsolutePath(@"..\..\Assets\Data\" + BtnChooseImageObstacle.AccessibleDescription)
        /// </summary>
        /// <param name="RelativePath"></param>
        /// <returns></returns>
        public static string AbsolutePath(string RelativePath)
        {
            return Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, RelativePath));
        }
    }
}

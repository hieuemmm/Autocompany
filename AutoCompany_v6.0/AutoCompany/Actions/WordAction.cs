using AutoCompany.Actions;
using AutoCompany.Model;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoCompany
{
    public class WordAction
    {
        private string PATH_VIETTEL;
        private string PATH_FAST;
        private string PATH_NCCA1;
        private string PATH_NCCA2;
        private string VINCA1;
        private string VINCA2;
        private string FPT1;
        private string FPT2;
        private string FOLDER_TODAY;

        public WordAction()
        {
            PATH_VIETTEL = FileAction.AbsolutePath(@"..\..\Assets\Templace\VIETTEL\HD chukiso.docx");
            PATH_FAST = FileAction.AbsolutePath(@"..\..\Assets\Templace\FAST\DK01.docx");
            PATH_NCCA1 = FileAction.AbsolutePath(@"..\..\Assets\Templace\NCCA\GXN.docx");
            PATH_NCCA2 = FileAction.AbsolutePath(@"..\..\Assets\Templace\NCCA\PHIEUDANGKI.docx");
            VINCA1 = FileAction.AbsolutePath(@"..\..\Assets\Templace\VINCA\Giaydangky.docx");
            VINCA2 = FileAction.AbsolutePath(@"..\..\Assets\Templace\VINCA\Giayxacnhan.docx");
            FPT1 = FileAction.AbsolutePath(@"..\..\Assets\Templace\FPT\Giayxacnhan-thaythe-BBBG.docx");
            FPT2 = FileAction.AbsolutePath(@"..\..\Assets\Templace\FPT\Phiếu đăng ký.docx");
            FOLDER_TODAY = DateTime.Now.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture).Replace(char.Parse("/"), char.Parse("-"));
            //Tạo Folder hôm nay để sẵn sàng hứng data bất cứ lúc nào
            CreateFolder(FileAction.AbsolutePath(@"..\..\Assets\KetQua\" + FOLDER_TODAY));
        }
        public void CreateWordDocument(List<Companny> compannies)
        {
            foreach (Companny companny in compannies)
            {
                //Tạo Folder cho công ty
                string FolderCompany = FileAction.AbsolutePath(@"..\..\Assets\KetQua\" + FOLDER_TODAY + @"\" + companny.NameTo2Word() + " - " + companny.TypeTEMPLATE);
                CreateFolder(FolderCompany);
                switch (companny.TypeTEMPLATE)
                {
                    case "VIETTEL":
                        CreateWord(
                            PATH_VIETTEL,
                            FolderCompany + @"\HD chukiso.docx",
                            companny
                        );
                        companny.StatusGET = "(1)";
                        break;
                    case "FAST":
                        CreateWord(
                            PATH_FAST,
                            FolderCompany + @"\DK01.docx",
                            companny
                        );
                        companny.StatusGET = "(1)";
                        break;
                    case "NCCA":
                        CreateWord(
                            PATH_NCCA1,
                            FolderCompany + @"\GXN.docx",
                            companny
                        );
                        CreateWord(
                            PATH_NCCA2,
                            FolderCompany + @"\PHIEUDANGKI.docx",
                            companny
                        );
                        companny.StatusGET = "(2)";
                        break;
                    case "VINCA":
                        CreateWord(
                            VINCA1,
                            FolderCompany + @"\Giaydangky.docx",
                            companny
                        );
                        CreateWord(
                            PATH_NCCA2,
                            FolderCompany + @"\Giayxacnhan.docx",
                            companny
                        );
                        companny.StatusGET = "(2)";
                        break;
                    case "FPT":
                        CreateWord(
                            FPT1,
                            FolderCompany + @"\Giayxacnhan-thaythe-BBBG.docx",
                            companny
                        );
                        CreateWord(
                            FPT2,
                            FolderCompany + @"\Phiếu đăng ký.docx",
                            companny
                        );
                        companny.StatusGET = "(2)";
                        break;
                    default:
                        companny.StatusGET = "T.Bại";
                        break;
                }
                FormMain.companniesTemplace = compannies;
            }
        }
        private void CreateFolder(string Dir)
        {
            if (!Directory.Exists(Dir))
            {
                Directory.CreateDirectory(Dir);
            }
        }
        //Find and Replace Method
        private void FindAndReplace(Microsoft.Office.Interop.Word.Application wordApp, object ToFindText, object replaceWithText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundLike = false;
            object nmatchAllforms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiactitics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            wordApp.Selection.Find.Execute(ref ToFindText,
                ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundLike,
                ref nmatchAllforms, ref forward,
                ref wrap, ref format, ref replaceWithText,
                ref replace, ref matchKashida,
                ref matchDiactitics, ref matchAlefHamza,
                ref matchControl);
        }
        //Create the Doc Method
        public void CreateWord(object filename, object SaveAs, Companny companny)
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            object missing = Missing.Value;
            Microsoft.Office.Interop.Word.Document myWordDoc = null;

            if (File.Exists((string)filename))
            {
                object readOnly = false;
                object isVisible = false;
                wordApp.Visible = false;

                myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                        ref missing, ref missing, ref missing,
                                        ref missing, ref missing, ref missing,
                                        ref missing, ref missing, ref missing,
                                        ref missing, ref missing, ref missing, ref missing);
                myWordDoc.Activate();

                //find and replace
                try
                {
                    FindAndReplace(wordApp, "<TENCONGTY>", companny.Name);
                    FindAndReplace(wordApp, "<DIACHI>", companny.Address);
                    FindAndReplace(wordApp, "<MASOTHUE>", companny.MST);
                    FindAndReplace(wordApp, "<TENGIAMDOC>", companny.RepresentativeName);
                    FindAndReplace(wordApp, "<NGAYCAP>", companny.OperationDate);
                    FindAndReplace(wordApp, "<TINHTHANH>", companny.AddressToCity());
                }
                catch (Exception)
                {
                }

                myWordDoc.SaveAs2(ref SaveAs, ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing);
            }
            else
            {
                MessageBox.Show("Không tìm thấy File mẫu " + filename.ToString());
            }

            myWordDoc.Close();
            wordApp.Quit();
        }
    }
}

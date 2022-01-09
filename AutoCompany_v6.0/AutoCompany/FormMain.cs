using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using AutoCompany.Actions;
using AutoCompany.DAO;
using AutoCompany.Model;
using AutoCompany.Models;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using Keys = System.Windows.Forms.Keys;

namespace AutoCompany
{
    public partial class FormMain : Form
    {
        private List<Companny> compannies;
        public static List<Companny> companniesTemplace;
        public static bool iRun = true;// Các biến cho nút lấy dữ liệu
        public static bool iClose = false;
        public static bool iRunTemplate = true;//được phép thêm xóa sửa list
        public static int counterTimeTemplate = 0;
        public FormMain()
        {
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;
            this.tabControl.SelectedIndex = (int)App.TABCONTROL.TAB_INFO;
            compannies = new List<Companny>();
            companniesTemplace = new List<Companny>();
        }
        private void FormMain_Load(object sender, EventArgs e)
        {
            dataGridViewCreateTemplate.Columns[0].Width = 30;
            dataGridViewCreateTemplate.Columns[1].Width = 90;
            dataGridViewCreateTemplate.Columns[2].Width = 90;
            dataGridViewCreateTemplate.Columns[3].Width = 55;
            dataGridViewCreateTemplate.Columns[4].Width = 40;
            textBoxLimitDate.Text = DateTime.Now.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
            richTextBox1.Text = System.IO.File.ReadAllText(FileAction.AbsolutePath(@"..\..\Assets\Data\ThongTinPhanMem.md"));
            LoadDGView(companniesTemplace);
            LoadListBox((int)App.LINKPAGE.LIST);
            LoadListBox((int)App.LINKPAGE.SELECTED);
        }
        public void LoadDGView(List<Companny> list)
        {
            dataGridViewCreateTemplate.Rows.Clear();
            int idx = 1;
            foreach (Companny companny in list)
            {
                var index = dataGridViewCreateTemplate.Rows.Add();
                dataGridViewCreateTemplate.Rows[index].Cells["STT"].Value = idx++;
                dataGridViewCreateTemplate.Rows[index].Cells["MST"].Value = companny.MST;
                dataGridViewCreateTemplate.Rows[index].Cells["NameDN"].Value = companny.NameTo2Word();
                dataGridViewCreateTemplate.Rows[index].Cells["Type"].Value = companny.TypeTEMPLATE;
                dataGridViewCreateTemplate.Rows[index].Cells["Status"].Value = companny.StatusGET;
            }
        }
        public void LoadListBox(int type)
        {
            List<LinkPage> listLinkPage = new LinkPageDAO().ReadAll(type);
            if (type == (int)App.LINKPAGE.LIST)
            {
                listBoxList.Items.Clear();
            }
            else
            {
                listBoxSelected.Items.Clear();
            }
            foreach (LinkPage linkPage in listLinkPage)
            {
                if (type == (int)App.LINKPAGE.LIST)
                {
                    listBoxList.Items.Add(linkPage.Name);
                }
                else
                {
                    listBoxSelected.Items.Add(linkPage.Name);
                }
            }
        }
        private List<Companny> GetListInfo()
        {
            FirefoxAction firefoxAction = new FirefoxAction(this);
            companniesTemplace = firefoxAction.GetInfoListCompany(companniesTemplace);
            return companniesTemplace;
        }
        private void save_temp()
        {
            if (iRunTemplate)
            {
                if (TextSave_Temp.Text.Trim().Length == 10 && checkIsNumberic(TextSave_Temp.Text.Trim()))
                {
                    string TypeTemplate = "";
                    if (radioButtonVIETTEL.Checked)
                        TypeTemplate = radioButtonVIETTEL.Text;
                    if (radioButtonFPT.Checked)
                        TypeTemplate = radioButtonFPT.Text;
                    if (radioButtonVINCA.Checked)
                        TypeTemplate = radioButtonVINCA.Text;
                    if (radioButtonFAST.Checked)
                        TypeTemplate = radioButtonFAST.Text;
                    if (radioButtonNCCA.Checked)
                        TypeTemplate = radioButtonNCCA.Text;

                    string MST = TextSave_Temp.Text.Trim();
                    Companny companny = new Companny(MST);
                    companny.TypeTEMPLATE = TypeTemplate;
                    companny.StatusGET = "x";
                    TextSave_Temp.Text = "";
                    companniesTemplace.Add(companny);
                    LoadDGView(companniesTemplace);
                }
                else
                {
                    TextSave_Temp.SelectAll();
                    TextSave_Temp.Focus();
                    MessageBox.Show("Mã số thuế không hợp lệ (Hợp lệ: 10 số)");
                }
            }
            else
            {
                MessageBox.Show("Có một tiến trình tạo hồ sơ đang chạy, hãy đợi...");
            }
        }
        private bool checkIsNumberic(string value)
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
        private void SaveInfo()
        {
            if (TextSaveInfo.Text.Trim().Length > 0)
            {
                string[] NameCity = TextSaveInfo.Text.Split(char.Parse("/"));
                string NameLink = "KHÔNG-TÊN";
                try
                {
                    NameLink = NameCity[NameCity.Count() - 2].ToUpper();
                    TextSaveInfo.Text = "";
                }
                catch (Exception)
                {
                    TextSaveInfo.SelectAll();
                    TextSaveInfo.Focus();
                    MessageBox.Show("Link Sai");
                    return;
                }
                new LinkPageDAO().Add(
                        new LinkPage
                        {
                            Name = NameLink,
                            Link = TextSaveInfo.Text.Trim()
                        },
                        (int)App.LINKPAGE.LIST
                    );
                LoadListBox((int)App.LINKPAGE.LIST);
            }
            else
            {
                MessageBox.Show("Vui lòng điền Link");
            }
        }
        public void NotifyInfo()
        {
            Thread threadNotifi = new Thread(() =>
            {
                string Status = "";
                if (App.ExcelCompany == App.ToTalCompany && App.TookLink == App.ToTalLink)
                {
                    Status = "Tiến trình hoàn tất!";
                }
                labelInfo.Text =
                "Link          : " + App.TookLink + "/" + App.ToTalLink + Environment.NewLine +
                "Công ty    : " + App.ToTalCompany + Environment.NewLine +
                "Excel        : " + App.ExcelCompany + "/" + App.ToTalCompany + Environment.NewLine +
                Status;
            });
            threadNotifi.IsBackground = true;
            threadNotifi.Start();

        }
        //SO SÁNH NGÀY
        private bool ConditionDate(string datetime1, String datetime2)
        {
            DateTime date1 = DateTime.ParseExact(datetime1, "dd/MM/yyyy", CultureInfo.InvariantCulture);
            DateTime date2 = DateTime.ParseExact(datetime2, "dd/MM/yyyy", CultureInfo.InvariantCulture);
            int result = DateTime.Compare(date1, date2);
            if (result <= 0)
                return true;
            return false;
        }
        //LẤY THÔNG TIN
        private void GetInfoCompany()
        {
            //RESET THỐNG KÊ
            App.ToTalCompany = 0;
            App.ToTalLink = 0;
            App.TookLink = 0;
            App.ExcelCompany = 0;
            labelInfoError.Visible = false;
            labelInfoError.Text = "Lịch sử";
            if (textBoxLimitDate.Text.ToString().Length == 10)
            {
                buttonGetInfo.Text = "Dừng tiến trình...";
                List<LinkPage> linkPageselected = new LinkPageDAO().ReadAll((int)App.LINKPAGE.SELECTED);

                App.ToTalLink = linkPageselected.Count;//THỐNG KÊ
                NotifyInfo();
                Thread ThreadRunInfo = new Thread(() =>
                {
                    foreach (LinkPage linkPage in linkPageselected)
                    {
                        if (iClose)
                        {
                            iRun = true;
                            iClose = false;
                            break;
                        }
                        InfoAction infoAction = new InfoAction();
                        linkPage.LinkChild = infoAction.GetListCompany(linkPage);
                        if (linkPage.LinkChild.Count < 1)
                        {
                            labelInfoError.Visible = true;
                            labelInfoError.Text = labelInfoError.Text + Environment.NewLine + "Link của " + linkPage.Name + " bị sai.";
                            App.TookLink++;//THỐNG KÊ
                            NotifyInfo();
                            continue;
                        }
                        //Lấy thông tin từng công ty
                        foreach (string link in linkPage.LinkChild)
                        {
                            if (iClose)
                            {
                                break;
                            }
                            Companny compannyChildPage = infoAction.GetInfoCompany(link);

                            if (ConditionDate(textBoxLimitDate.Text, compannyChildPage.LicenseDate))
                            {
                                linkPage.Compannies.Add(compannyChildPage);
                                App.ToTalCompany++;//THỐNG KÊ
                                NotifyInfo();
                                continue;
                            }
                            break;
                        }
                        App.TookLink++;//THỐNG KÊ
                        NotifyInfo();
                    }
                    if (iRun && !iClose)// bấm dừng
                    {
                        App.ToTalCompany = 0;
                        App.ToTalLink = 0;
                        App.TookLink = 0;
                        App.ExcelCompany = 0;
                        NotifyInfo();
                        buttonGetInfo.Text = "Lấy thông tin";
                        buttonViewResultInfo.Visible = false;
                        labelInfo.Text = "Nếu khoảng cách  là 10.000 bước, chỉ cần bạn bước 1 bước, AutoCompany sẽ chạy Taxi đến chở bạn đi";
                    }
                    else
                    {
                        iRun = true;
                        iClose = false;
                        ExcelAction excelAction = new ExcelAction();
                        excelAction.CreateExcelCompany(linkPageselected, this);
                        buttonGetInfo.Text = "Lấy thông tin";
                        buttonViewResultInfo.Visible = true;
                    }
                    //Xong hết việc thì reset lại
                    compannies = new List<Companny>();
                });
                ThreadRunInfo.IsBackground = true;
                ThreadRunInfo.Start();
            }
            else
            {
                MessageBox.Show(@"Vui lòng điền ngày đúng định dạng: dd/mm/yyyy. Ví dụ: 05/05/2020. Để lấy kết quả từ ngày đó đến ngày hiện tại");
            }
        }
        //TẠO HỢP ĐỒNG
        private void CreateTemplate()
        {
            iRunTemplate = !iRunTemplate;
            int timeout = companniesTemplace.Count * 30 + 100;
            counterTimeTemplate = timeout;
            buttonCreateTemplate.Enabled = false;
            if (companniesTemplace.Count > 0)
            {
                buttonCreateTemplate.Text = "Tiến trình lấy (" + timeout + " giây)...";
                timer.Start();
            }
            Thread acceptClient = new Thread(() =>
            {
                WordAction wordAction = new WordAction();
                if (companniesTemplace.Count > 0)
                {
                    companniesTemplace = GetListInfo();
                    wordAction.CreateWordDocument(companniesTemplace);
                }
                else
                {
                    MessageBox.Show("Không có công ty nào");
                }
                counterTimeTemplate = 0;
                buttonCreateTemplate.Text = "Tạo mẫu hợp đồng";
                buttonCreateTemplate.Enabled = true;
                iRunTemplate = !iRunTemplate;
            });
            acceptClient.IsBackground = true;
            acceptClient.Start();
        }
        //SỰ KIỆN
        private void tabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            System.Media.SoundPlayer player = new System.Media.SoundPlayer(FileAction.AbsolutePath(@"..\..\Assets\MP3\HatRu.wav"));
            if (this.tabControl.SelectedIndex == (int)App.TABCONTROL.TAB_INFO)
            {
                App.TabCurrent = (int)App.TABCONTROL.TAB_INFO;
                player.Stop();
            }
            else if (this.tabControl.SelectedIndex == (int)App.TABCONTROL.TAB_TEMPLATE)
            {
                App.TabCurrent = (int)App.TABCONTROL.TAB_TEMPLATE;
                player.Stop();
            }
            else if (this.tabControl.SelectedIndex == (int)App.TABCONTROL.TAB_INFO_APP)
            {
                App.TabCurrent = (int)App.TABCONTROL.TAB_INFO_APP;
                player.Play();
            }
        }
        private void buttonCreateTemplate_Click(object sender, EventArgs e)
        {
            CreateTemplate();
        }
        private void buttonSave_Temp_Click(object sender, EventArgs e)
        {
            save_temp();
        }
        private void dataGridViewCreateTemplate_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            if (iRunTemplate)
            {
                if (e.Row.Index >= 0)
                {
                    companniesTemplace.RemoveAt(e.Row.Index);
                }
            }
            else
            {
                MessageBox.Show("Có một tiến trình tạo hồ sơ đang chạy, hãy đợi...");
            }
        }
        private void dataGridViewCreateTemplate_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (iRunTemplate)
            {
                if (e.RowIndex >= 0)
                {
                    DataGridViewRow row = new DataGridViewRow();
                    row = dataGridViewCreateTemplate.Rows[e.RowIndex];
                    TextSave_Temp.Text = row.Cells[1].Value.ToString();
                    string typeTemplace = row.Cells[2].Value.ToString();
                    if (typeTemplace.Equals("VIETTEL"))
                        radioButtonVIETTEL.Checked = true;
                    if (typeTemplace.Equals("FPT"))
                        radioButtonFPT.Checked = true;
                    if (typeTemplace.Equals("NCCA"))
                        radioButtonNCCA.Checked = true;
                    if (typeTemplace.Equals("FAST"))
                        radioButtonFAST.Checked = true;
                    if (typeTemplace.Equals("VINCA"))
                        radioButtonVINCA.Checked = true;
                    TextSave_Temp.SelectAll();
                    TextSave_Temp.Focus();
                    companniesTemplace.RemoveAt(e.RowIndex);
                    LoadDGView(companniesTemplace);
                }
            }
            else
            {
                MessageBox.Show("Có một tiến trình tạo hồ sơ đang chạy, hãy đợi...");
            }
        }
        private void TextSave_Temp_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
            (e.KeyChar != '.'))
            {
                e.Handled = true;
            }

            // only allow one decimal point
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }
        private void listBoxList_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            int index = this.listBoxList.IndexFromPoint(e.Location);
            if (index != System.Windows.Forms.ListBox.NoMatches)
            {
                LinkPageDAO linkPageDAO = new LinkPageDAO();
                LinkPage linkPage = linkPageDAO.Read(new LinkPage(listBoxList.SelectedItem.ToString()), (int)App.LINKPAGE.LIST);
                linkPageDAO.Delete(new LinkPage(listBoxList.SelectedItem.ToString()), (int)App.LINKPAGE.LIST);
                linkPageDAO.Add(linkPage, (int)App.LINKPAGE.SELECTED);
                LoadListBox((int)App.LINKPAGE.LIST);
                LoadListBox((int)App.LINKPAGE.SELECTED);
            }
        }
        private void listBoxSelected_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            int index = this.listBoxSelected.IndexFromPoint(e.Location);
            if (index != System.Windows.Forms.ListBox.NoMatches)
            {
                LinkPageDAO linkPageDAO = new LinkPageDAO();
                LinkPage linkPage = linkPageDAO.Read(new LinkPage(listBoxSelected.SelectedItem.ToString()), (int)App.LINKPAGE.SELECTED);
                linkPageDAO.Delete(new LinkPage(listBoxSelected.SelectedItem.ToString()), (int)App.LINKPAGE.SELECTED);
                linkPageDAO.Add(linkPage, (int)App.LINKPAGE.LIST);
                LoadListBox((int)App.LINKPAGE.LIST);
                LoadListBox((int)App.LINKPAGE.SELECTED);
            }
        }
        private void buttonSaveInfo_Click(object sender, EventArgs e)
        {
            SaveInfo();
        }
        private void listBoxList_MouseClick(object sender, MouseEventArgs e)
        {
            int index = this.listBoxList.IndexFromPoint(e.Location);
            if (index != System.Windows.Forms.ListBox.NoMatches)
            {
                App.SelectListBoxCurrent = (int)App.LISTBOX.LIST;
            }
        }
        private void listBoxSelected_MouseClick(object sender, MouseEventArgs e)
        {
            int index = this.listBoxSelected.IndexFromPoint(e.Location);
            if (index != System.Windows.Forms.ListBox.NoMatches)
            {
                App.SelectListBoxCurrent = (int)App.LISTBOX.SELECTED;
            }
        }
        private void buttonGetInfo_Click(object sender, EventArgs e)
        {
            if (iRun)
            {
                iRun = !iRun;
                GetInfoCompany();
            }
            else
            {
                iClose = !iClose;
            }
        }
        private void buttonViewResultInfo_MouseHover(object sender, EventArgs e)
        {
            labelbuttonViewResultInfo.Visible = true;
        }
        private void buttonViewResultInfo_MouseLeave(object sender, EventArgs e)
        {
            labelbuttonViewResultInfo.Visible = false;
        }
        private void buttonViewResultInfo_Click(object sender, EventArgs e)
        {
            Process.Start(FileAction.AbsolutePath(@"..\..\Assets\Excel\Data.xlsx"));
        }
        private void timer_Tick(object sender, EventArgs e)
        {
            Thread TimerThread = new Thread(() =>
            {
                counterTimeTemplate--;
                if (counterTimeTemplate < 1)
                {
                    timer.Stop();
                    LoadDGView(companniesTemplace);
                }
                else
                {
                    if (companniesTemplace.Last().StatusGET.Equals("OK") && counterTimeTemplate % 4 == 0)
                    {
                        LoadDGView(companniesTemplace);
                    }
                    buttonCreateTemplate.Text = "Tiến trình lấy (" + counterTimeTemplate + " giây)...";
                }
            });
            TimerThread.IsBackground = true;
            TimerThread.Start();
        }
        private void buttonTemplaceReset_Click(object sender, EventArgs e)
        {
            if (iRunTemplate)
            {
                companniesTemplace = new List<Companny>();
                LoadDGView(companniesTemplace);
            }
            else
            {
                MessageBox.Show("Có một tiến trình tạo hồ sơ đang chạy, hãy đợi...");
            }
        }
        private void buttonViewResultTemplace_MouseHover(object sender, EventArgs e)
        {
            labelViewResultTemplace.Visible = true;
        }
        private void buttonViewResultTemplace_MouseLeave(object sender, EventArgs e)
        {
            labelViewResultTemplace.Visible = false;
        }
        private void buttonViewResultTemplace_Click(object sender, EventArgs e)
        {
            Process.Start("explorer.exe", FileAction.AbsolutePath(@"..\..\Assets\KetQua"));
        }
        private void FormMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                FirefoxAction.firefoxDriver.Close();
                FirefoxAction.firefoxDriver.Quit();
                Application.Exit();
            }
            catch (Exception)
            {
            }
        }
        //CALENDAR
        private void monthCalendar_DateChanged(object sender, DateRangeEventArgs e)
        {
            monthCalendar.Visible = false;
            var startDate = monthCalendar.SelectionStart.Date.ToString("dd/MM/yyyy");
            textBoxLimitDate.Text = startDate;
        }
        private void textBoxLimitDate_Click(object sender, EventArgs e)
        {
            monthCalendar.SetDate(DateTime.Now);
            monthCalendar.Visible = true;
        }
        private void textBoxLimitDate_Leave(object sender, EventArgs e)
        {
            monthCalendar.Visible = false;
        }
        private void labelInfo_Click(object sender, EventArgs e)
        {
            monthCalendar.Visible = false;
        }
        private void panel3_Click(object sender, EventArgs e)
        {
            monthCalendar.Visible = false;
        }
        //ROBOT
        private void pictureBox2_Click(object sender, EventArgs e)
        {
            this.tabControl.SelectedIndex = (int)App.TABCONTROL.TAB_INFO_APP;
        }
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            this.tabControl.SelectedIndex = (int)App.TABCONTROL.TAB_INFO_APP;
        }
        //HOT KEY
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.S))
            {
                if (App.TabCurrent == (int)App.TABCONTROL.TAB_TEMPLATE)
                {
                    save_temp();
                }
                if (App.TabCurrent == (int)App.TABCONTROL.TAB_INFO)
                {
                    SaveInfo();
                }
                return true;
            }
            if (keyData == (Keys.Control | Keys.D))
            {
                if (App.TabCurrent == (int)App.TABCONTROL.TAB_INFO)
                {
                    TextSaveInfo.Text = Clipboard.GetText();
                    SaveInfo();
                }
                if (App.TabCurrent == (int)App.TABCONTROL.TAB_TEMPLATE)
                {
                    TextSave_Temp.Text = Clipboard.GetText();
                    save_temp();
                }
                return true;
            }
            if (keyData == (Keys.Enter))
            {
                if (App.TabCurrent == (int)App.TABCONTROL.TAB_INFO)
                {
                    if (iRun)
                    {
                        iRun = !iRun;
                        GetInfoCompany();
                    }
                    else
                    {
                        iClose = !iClose;
                    }
                }
                if (App.TabCurrent == (int)App.TABCONTROL.TAB_TEMPLATE)
                {
                    CreateTemplate();
                }
                return true;
            }
            if (keyData == (Keys.Control | Keys.Back))
            {
                if (App.SelectListBoxCurrent == (int)App.LISTBOX.LIST)
                {
                    string NameDetele = listBoxList.SelectedItem.ToString();
                    LinkPageDAO linkPageDAO = new LinkPageDAO();
                    linkPageDAO.Delete(new LinkPage(listBoxList.SelectedItem.ToString()), (int)App.LINKPAGE.LIST);
                    LoadListBox((int)App.LINKPAGE.LIST);
                    labelInfo.Text = "Đã xóa - " + NameDetele;
                }
                if (App.SelectListBoxCurrent == (int)App.LISTBOX.SELECTED)
                {
                    string NameDetele = listBoxSelected.SelectedItem.ToString();
                    LinkPageDAO linkPageDAO = new LinkPageDAO();
                    linkPageDAO.Delete(new LinkPage(listBoxSelected.SelectedItem.ToString()), (int)App.LINKPAGE.SELECTED);
                    LoadListBox((int)App.LINKPAGE.SELECTED);
                    labelInfo.Text = "Đã xóa - " + NameDetele;
                }
                return true;
            }
            if (keyData == (Keys.Control | Keys.Space))
            {
                if (App.TabCurrent == (int)App.TABCONTROL.TAB_INFO)
                {
                    Process.Start(FileAction.AbsolutePath(@"..\..\Assets\Excel\Data.xlsx"));
                }
                if (App.TabCurrent == (int)App.TABCONTROL.TAB_TEMPLATE)
                {
                    Process.Start("explorer.exe", FileAction.AbsolutePath(@"..\..\Assets\KetQua"));
                }
                return true;
            }
            if (keyData == (Keys.Control | Keys.M))
            {
                this.tabControl.SelectedIndex = (int)App.TABCONTROL.TAB_INFO_APP;
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
    }
}

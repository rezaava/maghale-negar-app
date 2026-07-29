using MaghaleNegar.Constants;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using static MaghaleNegar.DedicatedFunctions;

namespace MaghaleNegar.Forms.MaghaleNegarManager.CreateDocument
{
    public class InfoListItem
    {
        public int Index { get; set; }
        public string AuthorName { get; set; }
        public string AffiliationFa { get; set; }
        public string AffiliationEn { get; set; }
    }

    public partial class CreateDocumentSlide6 : UserControl
    {
        // Action ها
        public Action TransitionDocumentManagerRequest { get; set; }
        public Action CloseForm { get; set; }
        public Action TransitionMoveNextCommand { get; set; }

        // متغیرها
        private List<string> authorNames = new List<string>();
        private string selectedUniversityType = "";
        private string _titleEn = "";
        private string _titleFa = "";
        private List<InfoListItem> infoList = new List<InfoListItem>();

        public CreateDocumentSlide6()
        {
            InitializeComponent();
            dgInfoList.ItemsSource = infoList;
            ValidateControls();
        }

        #region مقداردهی

        public void initializeVariables(
            List<string> authorNames,
            string titleEn,
            string titleFa) 
        {
            this.authorNames = authorNames ?? new List<string>();
            this._titleEn = titleEn ?? "";
            this._titleFa = titleFa ?? "";

            cmbNameList.ItemsSource = this.authorNames;
            if (this.authorNames.Count > 0)
                cmbNameList.SelectedIndex = 0;
            else
                cmbNameList.SelectedIndex = -1;

            SetDefaultUniversityState();
            ValidateControls();
        }

        #endregion

        #region رویدادها

        private void CmbNameList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ValidateControls();
        }

        private void ChkUniversityType_Checked(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.CheckBox chk = sender as System.Windows.Controls.CheckBox;

            if (chk == chkAzad)
            {
                chkDolati.IsChecked = false;
                selectedUniversityType = "Azad";
                SetAzadUniversityState();
            }
            else if (chk == chkDolati)
            {
                chkAzad.IsChecked = false;
                selectedUniversityType = "Dolati";
                SetDolatiUniversityState();
            }

            ValidateControls();
        }

        private void ChkUniversityType_Unchecked(object sender, RoutedEventArgs e)
        {
            if (chkAzad.IsChecked == false && chkDolati.IsChecked == false)
            {
                selectedUniversityType = "";
                SetDefaultUniversityState();
                ValidateControls();
            }
        }

        private void AnyTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            ValidateControls();
        }

        #endregion

        #region تغییر وضعیت بر اساس نوع دانشگاه

        private void SetDefaultUniversityState()
        {
            lblFaculty.Visibility = Visibility.Visible;
            txtFacultyFa.Visibility = Visibility.Visible;
            txtFacultyEn.Visibility = Visibility.Visible;
            gridFaculty.Visibility = Visibility.Visible;

            lblUniversity.Visibility = Visibility.Visible;
            txtUniversityFa.Visibility = Visibility.Visible;
            txtUniversityEn.Visibility = Visibility.Visible;
            gridUniversity.Visibility = Visibility.Visible;

            lblFaculty.Text = "نام دانشکده:";
        }

        private void SetAzadUniversityState()
        {
            lblFaculty.Text = "نام واحد:";

            lblUniversity.Visibility = Visibility.Collapsed;
            txtUniversityFa.Visibility = Visibility.Collapsed;
            txtUniversityEn.Visibility = Visibility.Collapsed;
            gridUniversity.Visibility = Visibility.Collapsed;

            lblFaculty.Visibility = Visibility.Visible;
            txtFacultyFa.Visibility = Visibility.Visible;
            txtFacultyEn.Visibility = Visibility.Visible;
            gridFaculty.Visibility = Visibility.Visible;
        }

        private void SetDolatiUniversityState()
        {
            lblFaculty.Text = "نام دانشکده:";

            lblFaculty.Visibility = Visibility.Visible;
            txtFacultyFa.Visibility = Visibility.Visible;
            txtFacultyEn.Visibility = Visibility.Visible;
            gridFaculty.Visibility = Visibility.Visible;

            lblUniversity.Visibility = Visibility.Visible;
            txtUniversityFa.Visibility = Visibility.Visible;
            txtUniversityEn.Visibility = Visibility.Visible;
            gridUniversity.Visibility = Visibility.Visible;
        }

        #endregion

        #region ساخت نام فایل

        private string GetFileNameFromTitle()
        {
            string titleEn = _titleEn;

            if (string.IsNullOrEmpty(titleEn))
            {
                return $"مقاله_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.docx";
            }

            char[] invalidChars = System.IO.Path.GetInvalidFileNameChars();
            foreach (char c in invalidChars)
            {
                titleEn = titleEn.Replace(c.ToString(), "");
            }

            string[] words = titleEn.Split(new char[] { ' ', '\t', '\n', '\r' },
                                           StringSplitOptions.RemoveEmptyEntries);

            int maxWords = 5;
            if (words.Length > maxWords)
            {
                words = words.Take(maxWords).ToArray();
            }

            string fileName = string.Join("_", words);

            if (string.IsNullOrEmpty(fileName))
            {
                fileName = $"مقاله_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}";
            }

            return $"{fileName}.docx";
        }

        #endregion

        #region ساخت پیش‌نمایش

        private string GetResponsibleAuthorEmail()
        {
            if (cmbNameList.SelectedItem != null)
            {
                string selected = cmbNameList.SelectedItem.ToString();
                int start = selected.IndexOf('(');
                int end = selected.IndexOf(')');
                if (start != -1 && end != -1 && end > start)
                {
                    return selected.Substring(start + 1, end - start - 1);
                }
            }
            return "";
        }

        private string GetPreviewTextFa()
        {
            // ====== گرفتن مستقیم از تکست‌باکس‌ها ======
            string degree = txtAcademicDegreeFa.Text.Trim();
            string group = txtGroupFa.Text.Trim();
            string faculty = txtFacultyFa.Text.Trim();
            string university = txtUniversityFa.Text.Trim();
            string city = txtCityFa.Text.Trim();

            if (string.IsNullOrEmpty(selectedUniversityType))
            {
                return "⚠️ لطفاً نوع دانشگاه را انتخاب کنید...";
            }

            if (selectedUniversityType == "Azad")
            {
                return $"{degree}، گروه {group}، واحد {faculty}، دانشگاه آزاد اسلامی {city}، ایران";
            }
            else // Dolati
            {
                return $"{degree}، گروه {group}، دانشکده {faculty} ... دانشگاه {university} ... {city}، ایران";
            }
        }

        private string GetPreviewTextEn()
        {
            // ====== گرفتن مستقیم از تکست‌باکس‌ها ======
            string degree = txtAcademicDegreeEn.Text.Trim();
            string group = txtGroupEn.Text.Trim();
            string faculty = txtFacultyEn.Text.Trim();
            string university = txtUniversityEn.Text.Trim();
            string city = txtCityEn.Text.Trim();
            string email = GetResponsibleAuthorEmail();

            if (string.IsNullOrEmpty(selectedUniversityType))
            {
                return "⚠️ Please select university type...";
            }

            if (selectedUniversityType == "Azad")
            {
                return $"{degree}, Department of {group}, {faculty} Branch, Islamic Azad University, {city}, Iran ({email})";
            }
            else // Dolati
            {
                return $"{degree}, Department of {group}, Faculty of {faculty}, University of {university}, {city}, Iran ({email})";
            }
        }

        private void UpdatePreview()
        {
            txtPreviewFa.Text = GetPreviewTextFa();
            txtPreviewEn.Text = GetPreviewTextEn();
        }

        #endregion

        #region لیست اطلاعات

        private void btnAddToList_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string authorName = cmbNameList.SelectedItem?.ToString() ?? "";
                string affiliationFa = txtAcademicDegreeFa.Text.Trim();
                string affiliationEn = txtAcademicDegreeEn.Text.Trim();

                if (string.IsNullOrEmpty(authorName))
                {
                    MessageBox.Show("لطفاً یک نویسنده انتخاب کنید.", "خطا",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (string.IsNullOrEmpty(affiliationFa) || affiliationFa.Contains("⚠️") ||
                    string.IsNullOrEmpty(affiliationEn) || affiliationEn.Contains("⚠️"))
                {
                    MessageBox.Show("لطفاً اطلاعات دانشگاهی را کامل کنید.", "خطا",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var newItem = new InfoListItem
                {
                    Index = infoList.Count + 1,
                    AuthorName = authorName,
                    AffiliationFa = affiliationFa,
                    AffiliationEn = affiliationEn
                };

                infoList.Add(newItem);
                RefreshDataGrid();
                ClearFields();
                ValidateControls();

                MessageBox.Show($"✅ اطلاعات نویسنده '{authorName}' با موفقیت به لیست اضافه شد.", "موفق",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطا در افزودن به لیست: {ex.Message}", "خطا",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void btnDeleteItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Button btn = sender as Button;
                var item = btn?.Tag as InfoListItem;

                if (item != null && infoList.Contains(item))
                {
                    if (MessageBox.Show($"آیا از حذف آیتم '{item.AuthorName}' مطمئن هستید؟",
                        "تأیید حذف", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                    {
                        infoList.Remove(item);
                        for (int i = 0; i < infoList.Count; i++)
                        {
                            infoList[i].Index = i + 1;
                        }
                        RefreshDataGrid();
                        ValidateControls();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطا در حذف: {ex.Message}", "خطا",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void RefreshDataGrid()
        {
            dgInfoList.ItemsSource = null;
            dgInfoList.ItemsSource = infoList;
        }

        private void ClearFields()
        {
            // ====== فقط فیلدهای دانشگاهی پاک شوند، کامبوباکس نویسنده دست نخورد ======
            chkAzad.IsChecked = false;
            chkDolati.IsChecked = false;
            selectedUniversityType = "";
            txtAcademicDegreeFa.Text = "";
            txtAcademicDegreeEn.Text = "";
            txtGroupFa.Text = "";
            txtGroupEn.Text = "";
            txtFacultyFa.Text = "";
            txtFacultyEn.Text = "";
            txtUniversityFa.Text = "";
            txtUniversityEn.Text = "";
            txtCityFa.Text = "";
            txtCityEn.Text = "";

            SetDefaultUniversityState();
            UpdatePreview();
        }

        #endregion

        #region دکمه ایجاد مقاله

        private async void btnCreateDocument_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                btnForward.IsEnabled = false;
                btnForward.Content = "⏳ در حال ایجاد...";

                var wordApp = Globals.ThisAddIn.Application;

                // ====== 1. گرفتن تمپلیت ======
                string resourceName = "MaghaleNegar.Templates.MainTemplate.docx";
                Assembly assembly = Assembly.GetExecutingAssembly();
                System.IO.Stream stream = assembly.GetManifestResourceStream(resourceName);

                if (stream == null)
                {
                    string[] allResources = assembly.GetManifestResourceNames();
                    throw new Exception($"فایل تمپلیت پیدا نشد!\n{string.Join("\n", allResources)}");
                }

                string templatesPath = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "Microsoft", "Templates", "MaghaleNegarTemplates");
                System.IO.Directory.CreateDirectory(templatesPath);
                string templatePath = System.IO.Path.Combine(templatesPath, "MainTemplate.docx");

                using (System.IO.FileStream fileStream = new System.IO.FileStream(templatePath, System.IO.FileMode.Create, System.IO.FileAccess.Write))
                {
                    stream.CopyTo(fileStream);
                }

                // ====== 2. بستن همه اسناد باز ======
                try
                {
                    while (wordApp.Documents.Count > 0)
                    {
                        Document doc = wordApp.Documents[1];
                        if (doc != null)
                        {
                            bool isBlank = string.IsNullOrEmpty(doc.FullName) && doc.Characters.Count < 3;
                            if (isBlank)
                            {
                                doc.Close(WdSaveOptions.wdDoNotSaveChanges);
                            }
                            else
                            {
                                doc.ActiveWindow.Visible = false;
                                break;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"خطا در بستن اسناد: {ex.Message}");
                }

                // ====== قبل از ایجاد سند ======
                var previousAlerts = wordApp.DisplayAlerts;
                wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;

                //try
                //{
                    // ====== 3. ایجاد سند جدید ======
                    Document newDoc = wordApp.Documents.Add(templatePath);

                    // ====== 3.5. جاگذاری اطلاعات ======
                    FillDocumentContent(newDoc);

                // ====== 4. ذخیره ======
                string workspacePath = Properties.Settings.Default.WorkSpaceDirectory;
                if (string.IsNullOrEmpty(workspacePath))
                {
                    workspacePath = System.IO.Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                        "MaghaleNegarWorkspace");
                    Properties.Settings.Default.WorkSpaceDirectory = workspacePath;
                    Properties.Settings.Default.Save();
                }

                System.IO.Directory.CreateDirectory(workspacePath);

                string fileName = GetFileNameFromTitle();
                string savePath = System.IO.Path.Combine(workspacePath, fileName);

                newDoc.SaveAs2(savePath);
                //}
                //finally
                //{
                //    wordApp.DisplayAlerts = previousAlerts;
                //}

                // ====== 5. نمایش سند جدید و مخفی کردن بقیه ======
                wordApp.Visible = true;

                foreach (Document doc in wordApp.Documents)
                {
                    try
                    {
                        if (doc != newDoc)
                        {
                            doc.ActiveWindow.Visible = false;
                        }
                    }
                    catch { }
                }

                newDoc.Activate();
                newDoc.ActiveWindow.Visible = true;

                // ====== 6. بستن سند خالی باقی‌مونده ======
                try
                {
                    for (int i = wordApp.Documents.Count; i >= 1; i--)
                    {
                        Document doc = wordApp.Documents[i];
                        if (doc != newDoc)
                        {
                            bool isBlank = string.IsNullOrEmpty(doc.FullName) && doc.Characters.Count < 3;
                            if (isBlank)
                            {
                                doc.Close(WdSaveOptions.wdDoNotSaveChanges);
                            }
                        }
                    }
                }
                catch { }

                // ====== 7. بستن فرم ======
                CloseForm?.Invoke();

                // ====== 8. تنظیمات اولیه سند ======
                SetupNewDocument(newDoc);

                // ====== 9. تنظیم Ribbon ======
                string ribbonTitle = $"{StringConstant.NameOfProject}";
                Ribbon.InitializeRibbon(ribbonTitle);
                Ribbon.setTabProperties(ribbonTitle, true);
                Ribbon.RibbonControlsVisibility(true);

                // ====== 10. نمایش پیام موفقیت ======
                MessageBox.Show("✅ مقاله با موفقیت ایجاد شد!", "موفق",
                    MessageBoxButton.OK, MessageBoxImage.Information);

                TransitionMoveNextCommand?.Invoke();
            }
            catch (Exception ex)
            {
                Globals.ThisAddIn.Application.Visible = true;
                MessageBox.Show($"❌ خطا در ایجاد مقاله:\n{ex.Message}", "خطا",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                btnForward.IsEnabled = true;
                btnForward.Content = "🚀 ایجاد مقاله";
            }
        }

        #endregion

        #region SetupNewDocument

        private void SetupNewDocument(Document doc)
        {
            try
            {
                doc.Content.LanguageID = WdLanguageID.wdPersian;

                doc.PageSetup.TopMargin = Globals.ThisAddIn.Application.CentimetersToPoints(2.5f);
                doc.PageSetup.BottomMargin = Globals.ThisAddIn.Application.CentimetersToPoints(2.5f);
                doc.PageSetup.LeftMargin = Globals.ThisAddIn.Application.CentimetersToPoints(2.5f);
                doc.PageSetup.RightMargin = Globals.ThisAddIn.Application.CentimetersToPoints(2.5f);

                DedicatedFunctions.addVariable(doc, VariableIdentifierIDs._variable_id_GUID.ToString(), StringConstant.GUID);
                DedicatedFunctions.addVariable(doc, VariableTypeIDs._variable_type_Document.ToString(), ((int)DocumentTypes.Nothing).ToString());
                DedicatedFunctions.addVariable(doc, VariableIdentifierIDs._variable_id_Hardware.ToString(), DedicatedFunctions.getUUID());

                doc.Save();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"خطا در SetupNewDocument: {ex.Message}");
            }
        }

        #endregion

        #region متدهای عمومی

        public void close()
        {
            CloseForm?.Invoke();
        }

        public void resetControls()
        {
            // ====== ریست کامل ======
            cmbNameList.SelectedIndex = -1;
            cmbNameList.ItemsSource = null;
            chkAzad.IsChecked = false;
            chkDolati.IsChecked = false;
            selectedUniversityType = "";
            authorNames = new List<string>();
            infoList.Clear();

            txtAcademicDegreeFa.Text = "";
            txtAcademicDegreeEn.Text = "";
            txtGroupFa.Text = "";
            txtGroupEn.Text = "";
            txtFacultyFa.Text = "";
            txtFacultyEn.Text = "";
            txtUniversityFa.Text = "";
            txtUniversityEn.Text = "";
            txtCityFa.Text = "";
            txtCityEn.Text = "";

            SetDefaultUniversityState();
            RefreshDataGrid();
            UpdatePreview();
            UpdateStatus("⚠️ لطفاً تمام فیلدها را تکمیل کنید", "#FF9800");
        }

        #endregion

        #region Helpers

        private void ValidateControls()
        {
            bool isValid = true;

            // ====== بررسی همه فیلدها ======
            if (cmbNameList.SelectedItem == null)
                isValid = false;

            if (string.IsNullOrEmpty(selectedUniversityType))
                isValid = false;

            if (string.IsNullOrEmpty(txtAcademicDegreeFa.Text))
                isValid = false;

            if (string.IsNullOrEmpty(txtAcademicDegreeEn.Text))
                isValid = false;

            if (string.IsNullOrEmpty(txtGroupFa.Text))
                isValid = false;

            if (string.IsNullOrEmpty(txtGroupEn.Text))
                isValid = false;

            if (selectedUniversityType == "Azad")
            {
                if (string.IsNullOrEmpty(txtFacultyFa.Text))
                    isValid = false;

                if (string.IsNullOrEmpty(txtFacultyEn.Text))
                    isValid = false;
            }
            else if (selectedUniversityType == "Dolati")
            {
                if (string.IsNullOrEmpty(txtFacultyFa.Text))
                    isValid = false;

                if (string.IsNullOrEmpty(txtFacultyEn.Text))
                    isValid = false;

                if (string.IsNullOrEmpty(txtUniversityFa.Text))
                    isValid = false;

                if (string.IsNullOrEmpty(txtUniversityEn.Text))
                    isValid = false;
            }

            if (string.IsNullOrEmpty(txtCityFa.Text))
                isValid = false;

            if (string.IsNullOrEmpty(txtCityEn.Text))
                isValid = false;

            // ====== دکمه افزودن به لیست ======
            btnAddToList.IsEnabled = isValid;

            // ====== دکمه ایجاد مقاله ======
            btnForward.IsEnabled = infoList.Count > 0;

            // ====== به‌روزرسانی پیش‌نمایش ======
            UpdatePreview();

            // ====== به‌روزرسانی وضعیت ======
            if (infoList.Count > 0)
                UpdateStatus($"✅ {infoList.Count} نویسنده به لیست اضافه شد! آماده ایجاد مقاله.", "#2E7D32");
            else if (isValid)
                UpdateStatus("✅ تمام اطلاعات تکمیل شد! روی 'افزودن به لیست' کلیک کنید.", "#2196F3");
            else
                UpdateStatus("⚠️ لطفاً تمام فیلدها را تکمیل کنید.", "#FF9800");
        }

        private void UpdateStatus(string message, string color)
        {
            lblStatus.Text = message;
            lblStatus.Foreground = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString(color);
        }

        #endregion

        /// <summary>
        /// جاگذاری اطلاعات در ContentControlهای تمپلیت
        /// </summary>
        private void FillDocumentContent(Document doc)
        {
            try
            {
                // ====== 1. عنوان مقاله (فارسی) ======
                string titleFa = _titleFa; // از اسلاید ۳ بگیر
                SetContentControlText(doc, "TitleFa", titleFa);

                // ====== 2. نام نویسنده‌ها (فارسی) ======
                string allAuthorsFa = string.Join("، ", infoList.Select(a => a.AuthorName));
                SetContentControlText(doc, "AuthorNamesFa", allAuthorsFa);

                
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"خطا در FillDocumentContent: {ex.Message}");
            }
        }

        /// <summary>
        /// تنظیم متن یک ContentControl با Tag مشخص
        /// </summary>
        private void SetContentControlText(Document doc, string tag, string text)
        {
            try
            {
                foreach (Microsoft.Office.Interop.Word.ContentControl cc in doc.ContentControls)
                {
                    if (cc.Tag == tag)
                    {
                        cc.Range.Text = text;
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"خطا در SetContentControlText برای {tag}: {ex.Message}");
            }
        }
    }
}
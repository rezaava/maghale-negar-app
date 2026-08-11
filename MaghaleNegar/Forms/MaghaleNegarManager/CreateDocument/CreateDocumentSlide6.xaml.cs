using DocumentFormat.OpenXml.Spreadsheet;
using MaghaleNegar.Constants;
using MaghaleNegar.Forms.MaghaleNegarManager.DocumentManager.View;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using static MaghaleNegar.DedicatedFunctions;

namespace MaghaleNegar.Forms.MaghaleNegarManager.CreateDocument
{
    public partial class CreateDocumentSlide6 : UserControl
    {
        // Action ها
        public Action TransitionDocumentManagerRequest { get; set; }
        public Action CloseForm { get; set; }
        public Action TransitionMoveNextCommand { get; set; }

        // اطلاعات از اسلایدهای قبلی
        private List<InfoListItem> infoList = new List<InfoListItem>();
        private List<InfoListItem> fullInfoList = new List<InfoListItem>();
        private string _titleFa = "";
        private string _titleEn = "";
        private string selectedUniversityType = "";
        private string _academicDegreeFa = "";
        private string _groupFa = "";
        private string _facultyFa = "";
        private string _universityFa = "";
        private string _cityFa = "";
        private string _previewText = "";
        private string _titleEnForFile = "";



        private bool isCompleted = false;
        private string DocumentName = "";

        public CreateDocumentSlide6()
        {
            InitializeComponent();
        }

        #region مقداردهی

        public void initializeVariables(
            List<InfoListItem> infoList,
            string titleFa,
            string titleEn,
            string universityType,
            string academicDegreeFa,
            string groupFa,
            string facultyFa,
            string universityFa,
            string cityFa,
            string previewText,
            string titleEnForFile,
            string documentName)
        {
            this.infoList = infoList ?? new List<InfoListItem>();
            this.fullInfoList = infoList ?? new List<InfoListItem>();
            this._titleFa = titleFa ?? "";
            this._titleEn = titleEn ?? "";
            this.selectedUniversityType = universityType ?? "";
            this._academicDegreeFa = academicDegreeFa ?? "";
            this._groupFa = groupFa ?? "";
            this._facultyFa = facultyFa ?? "";
            this._universityFa = universityFa ?? "";
            this._cityFa = cityFa ?? "";
            this._previewText = previewText ?? "";
            this._titleEnForFile = titleEnForFile ?? "";
            this.DocumentName = documentName ?? GetFileNameFromTitle();


            // پر کردن اطلاعات
            FillInformation();

            // ====== پر کردن کامبوباکس انتخاب نویسنده ======
            cmbAuthorSelector.ItemsSource = infoList.Select(a => a.AuthorName).ToList();
            if (infoList.Count > 0)
            {
                cmbAuthorSelector.SelectedIndex = 0;
                DisplayAuthorInfo(0);
            }

            ValidateControls();
        }

        #endregion

        #region Fill Information

        private void FillInformation()
        {
            // اطلاعات مقاله
            txtTitleFa.Text = $"📌 عنوان (فارسی): {_titleFa}";
            txtTitleEn.Text = $"📌 عنوان (انگلیسی): {_titleEn}";

            // اطلاعات نویسندگان
            dgAuthors.ItemsSource = infoList;

            // اطلاعات دانشگاهی
            string universityTypeText = selectedUniversityType == "Azad" ? "دانشگاه آزاد" : "دانشگاه دولتی";
            
        }

        #endregion

        #region نمایش اطلاعات نویسنده

        private void DisplayAuthorInfo(int selectedIndex)
        {
            try
            {
                if (selectedIndex < 0 || selectedIndex >= fullInfoList.Count)
                {
                    ClearAuthorInfo();
                    return;
                }

                var author = fullInfoList[selectedIndex];
                if (author == null)
                {
                    ClearAuthorInfo();
                    return;
                }

                txtAuthorName.Text = $"👤 نام: {author.AuthorName}";
                txtAuthorNameEn.Text = $"👤 Name: {author.AuthorNameEn}";
                txtAuthorAffiliationFa.Text = $"🏛️ وابستگی علمی (فارسی): {author.AffiliationFa}";
                txtAuthorAffiliationEn.Text = $"🏛️ Affiliation (English): {author.AffiliationEn}";
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"خطا در DisplayAuthorInfo: {ex.Message}");
            }
        }

        private void ClearAuthorInfo()
        {
            txtAuthorName.Text = "👤 نام: -";
            txtAuthorNameEn.Text = "👤 Name: -";
            txtAuthorAffiliationFa.Text = "🏛️ وابستگی علمی (فارسی): -";
            txtAuthorAffiliationEn.Text = "🏛️ Affiliation (English): -";
        }

        #endregion

        #region Events

        private void CmbAuthorSelector_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int selectedIndex = cmbAuthorSelector.SelectedIndex;

            if (selectedIndex >= 0 && selectedIndex < fullInfoList.Count)
            {
                DisplayAuthorInfo(selectedIndex);
            }
            else
            {
                ClearAuthorInfo();
            }
        }

        private void ChkConfirm_Checked(object sender, RoutedEventArgs e)
        {
            ValidateControls();
        }

        private void ChkConfirm_Unchecked(object sender, RoutedEventArgs e)
        {
            ValidateControls();
        }

        private async void btnCreateDocument_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                btnCreateDocument.IsEnabled = false;
                btnCreateDocument.Content = "⏳ در حال ایجاد...";

                var wordApp = Globals.ThisAddIn.Application;

                // ====== 1. گرفتن تمپلیت ======
                string resourceName = "MaghaleNegar.Templates.MainTemplate.docx";
                Assembly assembly = Assembly.GetExecutingAssembly();

                using (System.IO.Stream stream = assembly.GetManifestResourceStream(resourceName))
                {
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

                    // ====== 2. بستن اسناد باز ======
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

                    try
                    {
                        // ====== 3. ایجاد سند جدید ======
                        Document newDoc = wordApp.Documents.Add(templatePath);

                        // ====== 4. جاگذاری اطلاعات ======
                        FillDocumentContent(newDoc);

                        // ====== 5. ذخیره ======
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

                        // ====== 6. نمایش سند ======
                        wordApp.Visible = true;
                        foreach (Document doc in wordApp.Documents)
                        {
                            try
                            {
                                if (doc != newDoc)
                                    doc.ActiveWindow.Visible = false;
                            }
                            catch { }
                        }
                        newDoc.Activate();
                        newDoc.ActiveWindow.Visible = true;

                        // ====== 7. بستن سند خالی ======
                        try
                        {
                            for (int i = wordApp.Documents.Count; i >= 1; i--)
                            {
                                Document doc = wordApp.Documents[i];
                                if (doc != newDoc)
                                {
                                    bool isBlank = string.IsNullOrEmpty(doc.FullName) && doc.Characters.Count < 3;
                                    if (isBlank)
                                        doc.Close(WdSaveOptions.wdDoNotSaveChanges);
                                }
                            }
                        }
                        catch { }

                        // ====== 8. بستن فرم ======
                        CloseForm?.Invoke();

                        // ====== 9. تنظیمات اولیه ======
                        SetupNewDocument(newDoc);

                        // ====== 10. تنظیم Ribbon ======
                        string ribbonTitle = $"{StringConstant.NameOfProject}";
                        Ribbon.InitializeRibbon(ribbonTitle);
                        Ribbon.setTabProperties(ribbonTitle, true);
                        Ribbon.RibbonControlsVisibility(true);

                        // ====== ✅ 11. آپلود در سرور ======
                        bool uploadSuccess = await SaveToServer(newDoc);

                       

                        // ====== 12. انتقال به مرحله بعد ======
                        TransitionMoveNextCommand?.Invoke();
                    }
                    finally
                    {
                        wordApp.DisplayAlerts = previousAlerts;
                    }
                }
            }
            catch (Exception ex)
            {
                Globals.ThisAddIn.Application.Visible = true;
                MessageBox.Show($"❌ خطا در ایجاد مقاله:\n{ex.Message}", "خطا",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                btnCreateDocument.IsEnabled = true;
                btnCreateDocument.Content = "🚀 ایجاد مقاله";
            }
        }

        #endregion

        #region SetupNewDocument

        private void SetupNewDocument(Document doc)
        {
            try
            {
                doc.Content.LanguageID = WdLanguageID.wdPersian;

                doc.PageSetup.TopMargin = Globals.ThisAddIn.Application.CentimetersToPoints(4.5f);
                doc.PageSetup.BottomMargin = Globals.ThisAddIn.Application.CentimetersToPoints(2.5f);
                doc.PageSetup.LeftMargin = Globals.ThisAddIn.Application.CentimetersToPoints(2.5f);
                doc.PageSetup.RightMargin = Globals.ThisAddIn.Application.CentimetersToPoints(4.9f);

                DedicatedFunctions.addVariable(doc, VariableIdentifierIDs._variable_id_GUID.ToString(), StringConstant.GUID);
                DedicatedFunctions.addVariable(doc, VariableTypeIDs._variable_type_Document.ToString(), ((int)DocumentTypes.Nothing).ToString());
                DedicatedFunctions.addVariable(doc, VariableIdentifierIDs._variable_id_Hardware.ToString(), DedicatedFunctions.getUUID());

                // ====== متغیرهای سرور ======
                DedicatedFunctions.addVariable(doc, VariableServerIDs._variable_server_UserToken.ToString(), Properties.Settings.Default.UserToken);
                string version = BugReport.AssemblyVersion.Replace(".", "");
                DedicatedFunctions.addVariable(doc, VariableServerIDs._variable_server_VersionNumber.ToString(), version);
                DedicatedFunctions.addVariable(doc, VariableVersionIDs._variable_version_AddIn.ToString(), Properties.Settings.Default.VersionAddin.ToString());

                // ====== متغیرهای اطلاعات مقاله ======
                DedicatedFunctions.addVariable(doc, VariableFieldIDs._variable_field_Title_Fa.ToString(), _titleFa);
                DedicatedFunctions.addVariable(doc, VariableFieldIDs._variable_field_Title_En.ToString(), _titleEn);
                DedicatedFunctions.addVariable(doc, VariableFieldIDs._variable_field_University_Fa.ToString(), _universityFa);
                DedicatedFunctions.addVariable(doc, VariableFieldIDs._variable_field_AcademicDegree_Fa.ToString(), _academicDegreeFa);
                DedicatedFunctions.addVariable(doc, VariableFieldIDs._variable_field_Group_Fa.ToString(), _groupFa);
                DedicatedFunctions.addVariable(doc, "CityFa", _cityFa);
                // ====== متغیرهای نویسنده اول ======
                if (infoList != null && infoList.Count > 0)
                {
                    var firstAuthor = infoList[0];
                    DedicatedFunctions.addVariable(doc, VariableFieldIDs._variable_field_Author_Fa.ToString(), firstAuthor.AuthorName);
                    DedicatedFunctions.addVariable(doc, VariableFieldIDs._variable_field_Author_En.ToString(), firstAuthor.AuthorNameEn);
                    DedicatedFunctions.addVariable(doc, VariableFieldIDs._variable_field_University_Fa.ToString(), firstAuthor.AffiliationFa);
                    DedicatedFunctions.addVariable(doc, VariableFieldIDs._variable_field_University_En.ToString(), firstAuthor.AffiliationEn);
                }

                doc.Save();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"خطا در SetupNewDocument: {ex.Message}");
            }
        }

        #endregion

        #region Helpers

        private void ValidateControls()
        {
            bool isValid = chkConfirm.IsChecked == true;
            btnCreateDocument.IsEnabled = isValid;

            if (isValid)
            {
                lblStatus.Text = "✅ اطلاعات تأیید شد! آماده ایجاد مقاله.";
                lblStatus.Foreground = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#2E7D32");
            }
            else
            {
                lblStatus.Text = "⚠️ لطفاً اطلاعات را بررسی و تأیید کنید.";
                lblStatus.Foreground = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#FF9800");
            }
        }

        private void FillDocumentContent(Document doc)
        {
            try
            {
                // ====== عنوان مقاله ======
                SetContentControlText(doc, "TitleFa", _titleFa);
                SetContentControlText(doc, "TitleEn", _titleEn);

                // ====== نویسندگان فارسی + Footnote ======
                InsertAuthorsWithFootnotes(doc, "AuthorNamesFa", false);

                // ====== نویسندگان انگلیسی ======
                InsertAuthorsWithFootnotes(doc, "AuthorNamesEn", true);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"خطا در FillDocumentContent: {ex}");
                MessageBox.Show(
                    $"خطا در ایجاد اطلاعات نویسندگان:\n{ex.Message}",
                    "خطا",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }
        private void InsertAuthorsWithFootnotes(Document doc)
        {
            try
            {
                if (infoList == null || infoList.Count == 0)
                    return;

                // ====== نویسندگان فارسی ======
                if (doc.Bookmarks.Exists("AuthorNamesFa"))
                {
                    Range authorRange = doc.Bookmarks["AuthorNamesFa"].Range;

                    // پاک کردن محتوای قبلی Bookmark
                    authorRange.Text = "";

                    // بعد از تغییر Text، Bookmark ممکن است از بین برود
                    // بنابراین Range را دوباره از محل Bookmark اصلی نمی‌توان گرفت.
                    // موقعیت شروع را قبل از تغییر نگه می‌داریم.
                }

                // بهتر است Bookmark را با یک Range مشخص کنترل کنیم
                Bookmark bookmark = doc.Bookmarks["AuthorNamesFa"];
                Range range = bookmark.Range;

                int startPosition = range.Start;

                // پاک کردن محتوای Bookmark
                range.Text = "";

                // ایجاد Range جدید در محل Bookmark
                Range insertRange = doc.Range(startPosition, startPosition);

                for (int i = 0; i < infoList.Count; i++)
                {
                    var author = infoList[i];

                    if (i > 0)
                    {
                        insertRange.InsertAfter("، ");
                        insertRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    }

                    // نام نویسنده
                    insertRange.InsertAfter(author.AuthorName);
                    insertRange.Collapse(WdCollapseDirection.wdCollapseEnd);

                    // ====== ایجاد Footnote واقعی Word ======
                    string footnoteText = author.AffiliationFa;

                    if (string.IsNullOrWhiteSpace(footnoteText))
                        footnoteText = author.AffiliationEn;

                    if (string.IsNullOrWhiteSpace(footnoteText))
                        footnoteText = "وابستگی علمی مشخص نیست";

                    // ایجاد Footnote دقیقاً بعد از نام نویسنده
                    Footnote footnote = doc.Footnotes.Add(
                        insertRange,
                        false,
                        footnoteText
                    );

                    // تنظیم فونت Footnote
                    footnote.Range.Font.Size = 10;
                    footnote.Range.ParagraphFormat.Alignment =
                        WdParagraphAlignment.wdAlignParagraphRight;

                    // رفتن به انتهای Reference ایجاد شده
                    insertRange = doc.Range(
                        footnote.Reference.End,
                        footnote.Reference.End
                    );
                }

                // ====== نویسندگان انگلیسی ======
                if (doc.Bookmarks.Exists("AuthorNamesEn"))
                {
                    Bookmark bookmarkEn = doc.Bookmarks["AuthorNamesEn"];
                    Range rangeEn = bookmarkEn.Range;

                    int startPositionEn = rangeEn.Start;

                    rangeEn.Text = "";

                    Range insertRangeEn = doc.Range(
                        startPositionEn,
                        startPositionEn
                    );

                    for (int i = 0; i < infoList.Count; i++)
                    {
                        var author = infoList[i];

                        if (i > 0)
                        {
                            insertRangeEn.InsertAfter(", ");
                            insertRangeEn.Collapse(
                                WdCollapseDirection.wdCollapseEnd
                            );
                        }

                        // نام انگلیسی نویسنده
                        insertRangeEn.InsertAfter(author.AuthorNameEn);
                        insertRangeEn.Collapse(
                            WdCollapseDirection.wdCollapseEnd
                        );

                        // وابستگی انگلیسی
                        string footnoteText = author.AffiliationEn;

                        if (string.IsNullOrWhiteSpace(footnoteText))
                            footnoteText = author.AffiliationFa;

                        if (string.IsNullOrWhiteSpace(footnoteText))
                            footnoteText = "Affiliation not specified";

                        // Footnote واقعی
                        Footnote footnote = doc.Footnotes.Add(
                            insertRangeEn,
                            false,
                            footnoteText
                        );

                        footnote.Range.Font.Size = 10;
                        footnote.Range.ParagraphFormat.Alignment =
                            WdParagraphAlignment.wdAlignParagraphLeft;

                        insertRangeEn = doc.Range(
                            footnote.Reference.End,
                            footnote.Reference.End
                        );
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(
                    $"خطا در InsertAuthorsWithFootnotes: {ex.Message}"
                );
            }
        }

        private void SetBookmarkText(Document doc, string bookmarkName, string text)
        {
            try
            {
                if (doc.Bookmarks.Exists(bookmarkName))
                {
                    doc.Bookmarks[bookmarkName].Range.Text = text;
                }
            }
            catch (Exception ex)
            {
                //Debug.WriteLine($"خطا در SetBookmarkText برای {bookmarkName}: {ex.Message}");
            }
        }
        private void InsertAuthorsWithFootnotes(
            Document doc,
            string bookmarkName,
            bool isEnglish)
        {
            try
            {
                if (infoList == null || infoList.Count == 0)
                    return;

                if (!doc.Bookmarks.Exists(bookmarkName))
                {
                    Debug.WriteLine($"Bookmark پیدا نشد: {bookmarkName}");
                    return;
                }

                // محدوده Bookmark
                Range bookmarkRange = doc.Bookmarks[bookmarkName].Range;

                // موقعیت شروع Bookmark
                int start = bookmarkRange.Start;

                // حذف متن قبلی
                bookmarkRange.Text = "";

                // Range جدید برای درج نویسندگان
                Range insertRange = doc.Range(start, start);

                for (int i = 0; i < infoList.Count; i++)
                {
                    var author = infoList[i];

                    if (author == null)
                        continue;

                    // ==========================================
                    // جداکننده بین نویسندگان
                    // ==========================================
                    if (i > 0)
                    {
                        insertRange.InsertAfter(isEnglish ? ", " : "، ");

                        insertRange.Collapse(
                            WdCollapseDirection.wdCollapseEnd);
                    }

                    // ==========================================
                    // نام نویسنده
                    // ==========================================
                    string authorName = isEnglish
                        ? author.AuthorNameEn
                        : author.AuthorName;

                    if (string.IsNullOrWhiteSpace(authorName))
                        authorName = "-";

                    insertRange.InsertAfter(authorName);

                    insertRange.Collapse(
                        WdCollapseDirection.wdCollapseEnd);

                    // ==========================================
                    // متن Footnote
                    // ==========================================
                    string footnoteText = isEnglish
                        ? author.AffiliationEn
                        : author.AffiliationFa;

                    // اگر متن زبان موردنظر خالی بود
                    // از زبان دیگر استفاده کن
                    if (string.IsNullOrWhiteSpace(footnoteText))
                    {
                        footnoteText = isEnglish
                            ? author.AffiliationFa
                            : author.AffiliationEn;
                    }

                    if (string.IsNullOrWhiteSpace(footnoteText))
                    {
                        footnoteText = isEnglish
                            ? "Affiliation not specified"
                            : "وابستگی علمی مشخص نیست";
                    }

                    // ==========================================
                    // ایجاد Footnote واقعی Word
                    // ==========================================

                    object reference = Type.Missing;
                    object text = footnoteText;

                    Footnote footnote = doc.Footnotes.Add(
                        insertRange,
                        ref reference,
                        ref text
                    );

                    // ==========================================
                    // تنظیم Footnote
                    // ==========================================

                    footnote.Range.Font.Size = 10;

                    if (isEnglish)
                    {
                        footnote.Range.ParagraphFormat.Alignment =
                            WdParagraphAlignment.wdAlignParagraphLeft;
                    }
                    else
                    {
                        footnote.Range.ParagraphFormat.Alignment =
                            WdParagraphAlignment.wdAlignParagraphRight;
                    }

                    // ==========================================
                    // ادامه درج بعد از Reference
                    // ==========================================

                    insertRange = doc.Range(
                        footnote.Reference.End,
                        footnote.Reference.End
                    );
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(
                    $"خطا در InsertAuthorsWithFootnotes ({bookmarkName}): {ex}");

                throw;
            }
        }
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

        private string GetFileNameFromTitle()
        {
            string titleEn = _titleEnForFile;

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

        public void resetControls()
        {
            chkConfirm.IsChecked = false;
            btnCreateDocument.IsEnabled = false;
            lblStatus.Text = "⚠️ لطفاً اطلاعات را بررسی و تأیید کنید.";
            lblStatus.Foreground = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#FF9800");
        }

        public void close()
        {
            CloseForm?.Invoke();
        }

        #endregion


        #region آپلود در سرور

        private async System.Threading.Tasks.Task<bool> SaveToServer(Document doc)
        {
            try
            {
                if (doc == null)
                    return false;

                ShowLoading();

                string token = Properties.Settings.Default.UserToken;
                if (string.IsNullOrEmpty(token))
                {
                    Debug.WriteLine("⚠️ توکن کاربر یافت نشد!");
                    HideLoading(0);
                    return false;
                }

                // ====== 1. دریافت اطلاعات سند ======
                DocumentTypes documentType = DedicatedFunctions.getDocumentType(doc);
                JsonObject jsonVariables = DedicatedFunctions.variablesToJsonServer(doc);

                // ====== 2. اضافه کردن اطلاعات نویسنده‌ها ======
                if (infoList != null && infoList.Count > 0)
                {
                    var authorsList = new System.Text.Json.Nodes.JsonArray();
                    foreach (var author in infoList)
                    {
                        var authorObj = new JsonObject
                        {
                            ["name"] = author.AuthorName,
                            ["nameEn"] = author.AuthorNameEn,
                            ["affiliation"] = author.AffiliationFa,
                            ["affiliationEn"] = author.AffiliationEn
                        };
                        authorsList.Add(authorObj);
                    }
                    jsonVariables["authors"] = authorsList;
                }

                // ====== 3. اضافه کردن اطلاعات مقاله ======
                jsonVariables["titleFa"] = _titleFa;
                jsonVariables["titleEn"] = _titleEn;
                jsonVariables["documentName"] = DocumentName;
                jsonVariables["universityType"] = selectedUniversityType;
                jsonVariables["academicDegree"] = _academicDegreeFa;
                jsonVariables["group"] = _groupFa;
                jsonVariables["faculty"] = _facultyFa;
                jsonVariables["university"] = _universityFa;
                jsonVariables["city"] = _cityFa;

                // ====== 4. دریافت چکیده ======
                Microsoft.Office.Interop.Word.ContentControl[] abstractContentControl =
                    DedicatedFunctions.getContentControls(doc, ContentControlNames._field_Abstract_Fa.ToString());
                if (abstractContentControl != null && abstractContentControl.Length != 0)
                {
                    Range rangeAbstract = abstractContentControl[0].Range;
                    if (rangeAbstract != null)
                    {
                        string abstractText = rangeAbstract.Text.Trim();
                        if (!string.IsNullOrEmpty(abstractText))
                        {
                            if (jsonVariables.ContainsKey(VariableFieldIDs._variable_field_Abstract_Fa.ToString()))
                                jsonVariables[VariableFieldIDs._variable_field_Abstract_Fa.ToString()] = abstractText;
                            else
                                jsonVariables.Add(VariableFieldIDs._variable_field_Abstract_Fa.ToString(), abstractText);
                        }
                    }
                }

                // ====== 5. دریافت کلمات کلیدی ======
                Microsoft.Office.Interop.Word.ContentControl[] keywordsContentControl =
                    DedicatedFunctions.getContentControls(doc, ContentControlNames._field_Keywords_Fa.ToString());
                if (keywordsContentControl != null && keywordsContentControl.Length != 0)
                {
                    Range rangeKeywords = keywordsContentControl[0].Range;
                    if (rangeKeywords != null)
                    {
                        string keywordsText = rangeKeywords.Text.Trim();
                        if (!string.IsNullOrEmpty(keywordsText))
                        {
                            if (jsonVariables.ContainsKey("KeywordsFa"))
                                jsonVariables["KeywordsFa"] = keywordsText;
                            else
                                jsonVariables.Add("KeywordsFa", keywordsText);
                        }
                    }
                }

                // ====== 6. ساخت URL ======
                string urlParameters = $"save/maghalenegar/3?type={(int)documentType}&name={DocumentName}&config={jsonVariables.ToString()}";
                var formData = new MultipartFormDataContent();

                // ====== 7. اضافه کردن فایل ======
                // ====== ✅ متغیر response رو اینجا تعریف کن ======
                HttpResponseMessage response = null;

                using (var fileStream = new FileStream(doc.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    var fileContent = new StreamContent(fileStream);
                    formData.Add(fileContent, "file", "documentfile.docx");

                    // ====== 8. ارسال به سرور ======
                    response = await DedicatedFunctions.httpAsyncPostRequestAsync(
                        StringConstant.PrimaryServerApiBaseAddress,
                        urlParameters,
                        token,
                        formData);
                }

                // ====== 9. بررسی پاسخ (خارج از using) ======
                if (response != null && response.IsSuccessStatusCode)
                {
                    string result = await response.Content.ReadAsStringAsync();
                    try
                    {
                        JsonDocument document = JsonDocument.Parse(result);
                        JsonElement root = document.RootElement;

                        // ====== ذخیره ID و تاریخ به‌روزرسانی ======
                        if (root.TryGetProperty("id", out JsonElement idElement))
                        {
                            int documentID = idElement.GetInt32();
                            DedicatedFunctions.setORAddStaticVariableValue(doc,
                                VariableServerIDs._variable_server_DocumentID.ToString(),
                                documentID.ToString());
                        }

                        if (root.TryGetProperty("updated", out JsonElement updatedAtElement))
                        {
                            string updatedAt = updatedAtElement.GetString();
                            DedicatedFunctions.setORAddStaticVariableValue(doc,
                                VariableServerIDs._variable_server_UpdatedAt.ToString(), updatedAt);
                            DedicatedFunctions.setORAddStaticVariableValue(doc,
                                VariableServerIDs._variable_server_UpdatedFile.ToString(), updatedAt);
                            DedicatedFunctions.setORAddStaticVariableValue(doc,
                                VariableServerIDs._variable_server_UpdatedConfig.ToString(), updatedAt);
                        }

                        // ====== ذخیره نهایی سند ======
                        doc.Save();
                        Debug.WriteLine("✅ مقاله با موفقیت در سرور ذخیره شد!");

                        // ====== موفقیت ======
                        HideLoading(1, "✅ مقاله با موفقیت ایجاد شد!");
                        return true;
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"خطا در پردازش پاسخ سرور: {ex.Message}");
                        HideLoading(0);
                        return false;
                    }
                }
                else
                {
                    Debug.WriteLine($"❌ خطا در آپلود: {(response != null ? response.StatusCode.ToString() : "No response")}");
                    HideLoading(0);
                    return false;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"❌ خطا در SaveToServer: {ex.Message}");
                HideLoading(0);
                return false;
            }
        }


        #endregion



        #region Loading

        private LoadingForm loadingForm;

        private void ShowLoading()
        {
            Dispatcher.Invoke(() =>
            {
                if (loadingForm == null)
                {
                    loadingForm = new LoadingForm();
                    loadingForm.Show();
                }
            });
        }

        private void HideLoading(int status = 1, string successMessage = "")
        {
            Dispatcher.Invoke(() =>
            {
                if (loadingForm != null)
                {
                    loadingForm.closeForm(status);
                    loadingForm = null;
                }

                // ====== نمایش پیام موفقیت با تاخیر 1 ثانیه ======
                if (status == 1 && !string.IsNullOrEmpty(successMessage))
                {
                   

                    // ====== بستن فرم بعد از 1 ثانیه ======
                    System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
                    timer.Interval = 1000;  // 1 ثانیه
                    timer.Tick += (s, args) =>
                    {
                        timer.Stop();
                        timer.Dispose();
                        CloseForm?.Invoke();
                    };
                    timer.Start();
                }
            });
        }

        #endregion



    }
}
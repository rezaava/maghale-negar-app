using MaghaleNegar.Constants;
using MaghaleNegar.Constants.ComboBoxData;
using MaghaleNegar.Forms.MaghaleNegarManager.CreateDocument.Models;
using MaterialDesignThemes.Wpf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using static MaghaleNegar.Forms.MaghaleNegarManager.CreateDocument.Models.CreateDocumentControlModel;

namespace MaghaleNegar.Forms.MaghaleNegarManager.CreateDocument
{
    public class AuthorInfo
    {
        public string Name { get; set; }
        public string NameEn { get; set; }
        public string Email { get; set; }
        public bool IsResponsible { get; set; }
    }

    public partial class CreateDocumentSlide3 : UserControl
    {
        private List<CreateDocumentControlModel> textBoxControlModels;
        private List<AuthorInfo> authorsList = new List<AuthorInfo>();

        private DocumentTypes documentType;

        //Properties
        public Universities University { get; private set; }
        public string UniversityFa { get; private set; }
        public string UniversityEn { get; private set; }
        public string BranchFa { get; private set; }
        public string BranchEn { get; private set; }
        public string DepartmentFa { get; private set; }
        public string DepartmentEn { get; private set; }
        public string GroupFa { get; private set; }
        public string GroupEn { get; private set; }
        public string FieldOfStudyFa { get; private set; }
        public string FieldOfStudyEn { get; private set; }
        public string AreaOfStudyFa { get; private set; }
        public string AreaOfStudyEn { get; private set; }
        public string AcademicDegreeFa { get; private set; }
        public string AcademicDegreeEn { get; private set; }
        public string Email { get; private set; }
        public bool IsResponsibleAuthor { get; private set; }

        public Action TransitionMoveNextCommand { get; set; }

        public List<string> AuthorNames { get; private set; } = new List<string>();

        public List<string> AuthorNamesEn { get; private set; } = new List<string>();


        // متغیرهای حالت ویرایش
        private bool isEditingMode = false;
        private AuthorInfo editingAuthor = null;
        private AuthorInfo originalAuthor = null;

        public CreateDocumentSlide3()
        {
            InitializeComponent();

            btnForward.IsEnabled = false;

            // مقداردهی لیست نویسندگان
            authorsList = new List<AuthorInfo>();
            dgAuthors.ItemsSource = authorsList;

            // تنظیم فیلدها
            textBoxControlModels = new List<CreateDocumentControlModel>()
            {
                new CreateDocumentControlModel(txtBoxFieldOfStudy, CreateDocumentControlModel.ControlLevels.Essential),
                new CreateDocumentControlModel(txtBoxFieldOfStudyEn, CreateDocumentControlModel.ControlLevels.Essential),
                new CreateDocumentControlModel(txtBoxAreaOfStudy, CreateDocumentControlModel.ControlLevels.Essential),
                new CreateDocumentControlModel(txtBoxAreaOfStudyEn, CreateDocumentControlModel.ControlLevels.Essential),
                new CreateDocumentControlModel(txtBoxEmail, CreateDocumentControlModel.ControlLevels.Optional),
            };

            foreach (var controlModel in textBoxControlModels)
            {
                TextBox textBox = controlModel.Control as TextBox;
                textBox.TextChanged += TextBox_TextChanged;
                textBox.GotFocus += TextBox_GotFocus;
                textBox.LostFocus += TextBox_LostFocus;
            }

            // تنظیم اولیه ایمیل
            txtBoxEmail.IsEnabled = false;
            txtBoxEmail.Text = "";

            // رویدادهای CheckBox
            chkResponsibleAuthor.Checked += ChkResponsibleAuthor_Checked;
            chkResponsibleAuthor.Unchecked += ChkResponsibleAuthor_Unchecked;

            // رویداد کلیک دکمه بعدی
            btnForward.Click += BtnForward_Click;

            // تنظیم اولیه وضعیت دکمه‌ها
            UpdateButtonsState();
        }

        #region دکمه‌ها

        private void BtnForward_Click(object sender, RoutedEventArgs e)
        {
            if (validateControls())
            {
                SaveData();
                TransitionMoveNextCommand?.Invoke();
            }
        }

        private void SaveData()
        {
            FieldOfStudyFa = txtBoxFieldOfStudy.Text;
            FieldOfStudyEn = txtBoxFieldOfStudyEn.Text;
            AreaOfStudyFa = txtBoxAreaOfStudy.Text;
            AreaOfStudyEn = txtBoxAreaOfStudyEn.Text;
            Email = txtBoxEmail.Text;
            IsResponsibleAuthor = chkResponsibleAuthor.IsChecked ?? false;

            // ====== این بخش رو اضافه کن ======
            AuthorNames = new List<string>();
            AuthorNamesEn = new List<string>();

            foreach (var author in authorsList)
            {
                
                string displayName = author.Name;
                if (!string.IsNullOrEmpty(author.Email))
                {
                    displayName += $" ({author.Email})";
                }
                AuthorNames.Add(displayName);
            }

            foreach (var author in authorsList)
            {
                AuthorNamesEn.Add(author.NameEn);
            }
        }

        private void btnAddAuthor_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string authorName = txtBoxAreaOfStudy.Text.Trim();
                string authorNameEn = txtBoxAreaOfStudyEn.Text.Trim();
                string email = txtBoxEmail.Text.Trim();
                bool isResponsible = chkResponsibleAuthor.IsChecked ?? false;

                // بررسی نام نویسنده
                if (string.IsNullOrEmpty(authorName))
                {
                    MessageBox.Show("لطفاً نام نویسنده را وارد کنید.", "خطا",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (string.IsNullOrEmpty(authorNameEn))
                {
                    MessageBox.Show("لطفاً نام انگلیسی نویسنده را وارد کنید.", "خطا",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // ✅ اگر در حالت ویرایش هستیم
                if (isEditingMode && editingAuthor != null)
                {
                    // بررسی: اگر ایمیل وارد شده و قبلاً در لیست وجود دارد (به جز خود نویسنده در حال ویرایش)
                    if (!string.IsNullOrEmpty(email) && authorsList.Any(a => a.Email == email && a != editingAuthor))
                    {
                        MessageBox.Show("این ایمیل قبلاً برای نویسنده دیگری ثبت شده است.",
                            "خطا", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    // به‌روزرسانی نویسنده
                    editingAuthor.Name = authorName;
                    editingAuthor.NameEn = authorNameEn;
                    editingAuthor.Email = email;
                    editingAuthor.IsResponsible = isResponsible;

                    // بررسی: اگر نویسنده مسئول است و قبلاً نویسنده مسئول دیگری وجود دارد
                    if (isResponsible && authorsList.Any(a => a.IsResponsible && a != editingAuthor))
                    {
                        MessageBox.Show("قبلاً یک نویسنده مسئول ثبت شده است. ابتدا نویسنده مسئول قبلی را ویرایش کنید.",
                            "خطا", MessageBoxButton.OK, MessageBoxImage.Warning);
                        chkResponsibleAuthor.IsChecked = false;
                        return;
                    }

                    // اضافه کردن نویسنده ویرایش‌شده به لیست
                    authorsList.Add(editingAuthor);
                    RefreshDataGrid();

                    // پاک کردن فیلدها
                    txtBoxAreaOfStudy.Text = "";
                    txtBoxAreaOfStudyEn.Text = "";
                    txtBoxEmail.Text = "";
                    txtBoxEmail.IsEnabled = false;
                    chkResponsibleAuthor.IsChecked = false;

                    // بازگشت به حالت عادی
                    isEditingMode = false;
                    editingAuthor = null;
                    originalAuthor = null;
                    UpdateButtonsState();

                    // پاک کردن خطاها
                    normalControl(txtBoxAreaOfStudy);
                    normalControl(txtBoxAreaOfStudyEn);

                    validateControls();

                    MessageBox.Show($"نویسنده '{authorName}' با موفقیت ویرایش شد.", "موفق",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                // ✅ حالت عادی - افزودن نویسنده جدید
                // بررسی: اگر قبلاً نویسنده مسئول وجود دارد و کاربر می‌خواهد یکی دیگر اضافه کند
                if (isResponsible && authorsList.Any(a => a.IsResponsible))
                {
                    MessageBox.Show("قبلاً یک نویسنده مسئول ثبت شده است. ابتدا نویسنده مسئول قبلی را ویرایش کنید.",
                        "خطا", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // بررسی: اگر ایمیل وارد شده و قبلاً در لیست وجود دارد
                if (!string.IsNullOrEmpty(email) && authorsList.Any(a => a.Email == email))
                {
                    MessageBox.Show("این ایمیل قبلاً برای نویسنده دیگری ثبت شده است.",
                        "خطا", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // ایجاد نویسنده جدید
                var newAuthor = new AuthorInfo
                {
                    Name = authorName,
                    NameEn = authorNameEn,
                    Email = email,
                    IsResponsible = isResponsible
                };

                authorsList.Add(newAuthor);
                RefreshDataGrid();

                // پاک کردن فیلدها
                txtBoxAreaOfStudy.Text = "";
                txtBoxAreaOfStudyEn.Text = "";
                txtBoxEmail.Text = "";
                txtBoxEmail.IsEnabled = false;
                chkResponsibleAuthor.IsChecked = false;

                // پاک کردن خطاها
                normalControl(txtBoxAreaOfStudy);
                normalControl(txtBoxAreaOfStudyEn);

                validateControls();

                MessageBox.Show($"نویسنده '{authorName}' با موفقیت به لیست اضافه شد.", "موفق",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطا در افزودن نویسنده: {ex.Message}", "خطا",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void btnEditAuthor_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // ✅ اگر در حالت ویرایش هستیم، اجازه نده
                if (isEditingMode)
                {
                    MessageBox.Show("لطفاً ابتدا ویرایش فعلی را کامل کنید یا انصراف دهید.",
                        "توجه", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                Button btn = sender as Button;
                var author = btn?.Tag as AuthorInfo;

                if (author == null || !authorsList.Contains(author))
                {
                    MessageBox.Show("نویسنده مورد نظر یافت نشد.", "خطا",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // تنظیم حالت ویرایش
                isEditingMode = true;
                editingAuthor = author;

                // ذخیره کپی از نویسنده اصلی برای بازگشت در صورت انصراف
                originalAuthor = new AuthorInfo
                {
                    Name = author.Name,
                    NameEn = author.NameEn,
                    Email = author.Email,
                    IsResponsible = author.IsResponsible
                };

                // پر کردن فیلدها با اطلاعات نویسنده
                txtBoxAreaOfStudy.Text = author.Name;
                txtBoxAreaOfStudyEn.Text = author.NameEn;
                txtBoxEmail.Text = author.Email;
                chkResponsibleAuthor.IsChecked = author.IsResponsible;

                if (author.IsResponsible)
                {
                    txtBoxEmail.IsEnabled = true;
                }

                // حذف نویسنده از لیست
                authorsList.Remove(author);
                RefreshDataGrid();

                // به‌روزرسانی وضعیت دکمه‌ها
                UpdateButtonsState();

                // فوکوس روی فیلد نام
                txtBoxAreaOfStudy.Focus();
                validateControls();

                MessageBox.Show($"نویسنده '{author.Name}' برای ویرایش انتخاب شد. پس از ویرایش، روی دکمه 'ویرایش نویسنده' کلیک کنید.",
                    "ویرایش", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطا در ویرایش نویسنده: {ex.Message}", "خطا",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void btnDeleteAuthor_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // ✅ اگر در حالت ویرایش هستیم، اجازه حذف نده
                if (isEditingMode)
                {
                    MessageBox.Show("لطفاً ابتدا ویرایش را کامل کنید یا روی 'انصراف از ویرایش' کلیک کنید.",
                        "توجه", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                Button btn = sender as Button;
                var author = btn?.Tag as AuthorInfo;

                if (author != null && authorsList.Contains(author))
                {
                    string message = $"آیا از حذف نویسنده '{author.Name}' مطمئن هستید؟";

                    if (author.IsResponsible)
                    {
                        message += "\n\n⚠️ این نویسنده مسئول است. پس از حذف، باید نویسنده مسئول جدیدی انتخاب کنید.";
                    }

                    if (MessageBox.Show(message, "تأیید حذف",
                        MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                    {
                        authorsList.Remove(author);
                        RefreshDataGrid();
                        validateControls();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطا در حذف نویسنده: {ex.Message}", "خطا",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void btnCancelEdit_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // ✅ اگر نویسنده اصلی وجود دارد، آن را به لیست برگردان
                if (originalAuthor != null)
                {
                    // بررسی: آیا نویسنده قبلاً در لیست نیست؟
                    bool exists = authorsList.Any(a =>
                        a.Name == originalAuthor.Name &&
                        a.NameEn == originalAuthor.NameEn &&
                        a.Email == originalAuthor.Email);

                    if (!exists)
                    {
                        authorsList.Add(originalAuthor);
                        RefreshDataGrid();
                    }
                }

                // بازگشت به حالت عادی
                isEditingMode = false;
                editingAuthor = null;
                originalAuthor = null;
                UpdateButtonsState();

                // پاک کردن فیلدها
                txtBoxAreaOfStudy.Text = "";
                txtBoxAreaOfStudyEn.Text = "";
                txtBoxEmail.Text = "";
                txtBoxEmail.IsEnabled = false;
                chkResponsibleAuthor.IsChecked = false;

                // پاک کردن خطاها
                normalControl(txtBoxAreaOfStudy);
                normalControl(txtBoxAreaOfStudyEn);

                validateControls();

                MessageBox.Show("ویرایش لغو شد و نویسنده به لیست بازگردانده شد.", "انصراف",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطا در انصراف از ویرایش: {ex.Message}", "خطا",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void RefreshDataGrid()
        {
            dgAuthors.ItemsSource = null;
            dgAuthors.ItemsSource = authorsList;
        }

        private void UpdateButtonsState()
        {
            // در حالت ویرایش، دکمه‌ها و DataGrid را غیرفعال کن
            dgAuthors.IsEnabled = !isEditingMode;
            

            // تغییر ظاهر دکمه افزودن
            if (isEditingMode)
            {
                btnAddAuthor.Content = "✏️ ویرایش نویسنده";
                btnCancelEdit.Visibility = Visibility.Visible;
            }
            else
            {
                btnAddAuthor.Content = "➕ افزودن به لیست نویسندگان";
                btnCancelEdit.Visibility = Visibility.Collapsed;
            }
        }

        #endregion

        #region CheckBox Events
        private void ChkResponsibleAuthor_Checked(object sender, RoutedEventArgs e)
        {
            // ✅ اگر در حالت ویرایش هستیم و نویسنده در حال ویرایش، اجازه تغییر بده
            if (isEditingMode && editingAuthor != null)
            {
                txtBoxEmail.IsEnabled = true;
                txtBoxEmail.Focus();

                var emailControlModel = textBoxControlModels.FirstOrDefault(m => m.Control == txtBoxEmail);
                if (emailControlModel != null)
                {
                    emailControlModel.ControlLevel = ControlLevels.Essential;
                    emailControlModel.Validate = validateTextBox(txtBoxEmail, ControlLevels.Essential, false);
                }
                validateControls();
                return;
            }

            // ✅ حالت عادی - بررسی نویسنده مسئول تکراری
            if (authorsList.Any(a => a.IsResponsible))
            {
                MessageBox.Show("قبلاً یک نویسنده مسئول ثبت شده است. ابتدا نویسنده مسئول قبلی را ویرایش کنید.",
                    "خطا", MessageBoxButton.OK, MessageBoxImage.Warning);

                chkResponsibleAuthor.IsChecked = false;
                return;
            }

            txtBoxEmail.IsEnabled = true;
            txtBoxEmail.Focus();

            var emailModel = textBoxControlModels.FirstOrDefault(m => m.Control == txtBoxEmail);
            if (emailModel != null)
            {
                emailModel.ControlLevel = ControlLevels.Essential;
                emailModel.Validate = validateTextBox(txtBoxEmail, ControlLevels.Essential, false);
            }
            validateControls();
        }

        private void ChkResponsibleAuthor_Unchecked(object sender, RoutedEventArgs e)
        {
            txtBoxEmail.IsEnabled = false;
            txtBoxEmail.Text = "";

            var emailModel = textBoxControlModels.FirstOrDefault(m => m.Control == txtBoxEmail);
            if (emailModel != null)
            {
                emailModel.ControlLevel = ControlLevels.Optional;
                emailModel.Validate = true;
                normalControl(txtBoxEmail);
            }
            validateControls();
        }
        #endregion

        #region TextBox Events
        private void TextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            TextBox textBox = (TextBox)sender;

            for (int i = 0; i < textBoxControlModels.Count; i++)
            {
                if (textBox == (TextBox)textBoxControlModels[i].Control)
                {
                    textBoxControlModels[i].Validate = validateTextBox(textBox, textBoxControlModels[i].ControlLevel, false);
                    validateControls();
                    return;
                }
            }
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = (TextBox)sender;

            if (textBox == txtBoxAreaOfStudy)
            {
                if (!string.IsNullOrEmpty(textBox.Text))
                {
                    textBoxControlModels.Where(a => a.Control == txtBoxAreaOfStudyEn).FirstOrDefault().ControlLevel = ControlLevels.Essential;
                }
                else
                {
                    textBoxControlModels.Where(a => a.Control == txtBoxAreaOfStudyEn).FirstOrDefault().ControlLevel = ControlLevels.Optional;
                }

                textBoxControlModels.Where(a => a.Control == txtBoxAreaOfStudyEn).FirstOrDefault().Validate = validateTextBox(txtBoxAreaOfStudyEn, textBoxControlModels.Where(a => a.Control == txtBoxAreaOfStudyEn).FirstOrDefault().ControlLevel, false);
                validateControls();
            }
            else if (textBox == txtBoxAreaOfStudyEn)
            {
                if (!string.IsNullOrEmpty(textBox.Text))
                {
                    textBoxControlModels.Where(a => a.Control == txtBoxAreaOfStudy).FirstOrDefault().ControlLevel = ControlLevels.Essential;
                }
                else
                {
                    textBoxControlModels.Where(a => a.Control == txtBoxAreaOfStudy).FirstOrDefault().ControlLevel = ControlLevels.Optional;
                }

                textBoxControlModels.Where(a => a.Control == txtBoxAreaOfStudy).FirstOrDefault().Validate = validateTextBox(txtBoxAreaOfStudy, textBoxControlModels.Where(a => a.Control == txtBoxAreaOfStudy).FirstOrDefault().ControlLevel, false);
                validateControls();
            }

            for (int i = 0; i < textBoxControlModels.Count; i++)
            {
                if (textBox == (TextBox)textBoxControlModels[i].Control)
                {
                    textBoxControlModels[i].Validate = validateTextBox(textBox, textBoxControlModels[i].ControlLevel, true);
                    if (textBoxControlModels[i].Validate)
                        validateControls();
                    return;
                }
            }
        }

        private void TextBox_GotFocus(object sender, RoutedEventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            string tag = textBox.Tag.ToString();

            if (!string.IsNullOrEmpty(tag.Trim()))
            {
                if (tag == "Persian")
                    DedicatedFunctions.changeKeyboardLanguage(KeyboardLanguage.Persian);
                else if (tag == "English")
                    DedicatedFunctions.changeKeyboardLanguage(KeyboardLanguage.English);
            }
        }
        #endregion

        #region Validators
        private bool validateControls()
        {
            bool isValid = true;

            // 1. بررسی فیلدهای متنی
            foreach (CreateDocumentControlModel controlModel in textBoxControlModels)
            {
                if (!controlModel.Validate && controlModel.ControlLevel != CreateDocumentControlModel.ControlLevels.Optional)
                {
                    isValid = false;
                    break;
                }
            }

            // 2. بررسی لیست نویسندگان (حداقل یک نویسنده)
            if (isValid && authorsList.Count == 0)
            {
                isValid = false;

            }
            else
            {
                normalControl(txtBoxAreaOfStudy);
                normalControl(txtBoxAreaOfStudyEn);
            }

            // 3. بررسی: آیا نویسنده مسئول وجود دارد؟
            if (isValid && !authorsList.Any(a => a.IsResponsible))
            {
                isValid = false;

            }

            // 4. مقداردهی نهایی
            if (isValid)
            {
                FieldOfStudyFa = txtBoxFieldOfStudy.Text;
                FieldOfStudyEn = txtBoxFieldOfStudyEn.Text;
                AreaOfStudyFa = txtBoxAreaOfStudy.Text;
                AreaOfStudyEn = txtBoxAreaOfStudyEn.Text;
                Email = txtBoxEmail.Text;
                IsResponsibleAuthor = chkResponsibleAuthor.IsChecked ?? false;

                btnForward.IsEnabled = true;
                return true;
            }
            else
            {
                FieldOfStudyFa = "";
                FieldOfStudyEn = "";
                AreaOfStudyFa = "";
                AreaOfStudyEn = "";
                Email = "";
                IsResponsibleAuthor = false;

                btnForward.IsEnabled = false;
                return false;
            }
        }

        private bool validateTextBox(TextBox textBox, ControlLevels controlLevel, bool onlyReturn)
        {
            if (controlLevel == ControlLevels.Optional)
            {
                normalControl(textBox);
                return true;
            }

            if (!string.IsNullOrEmpty(textBox.Text) && !string.IsNullOrWhiteSpace(textBox.Text) && textBox.Text.Length > 2)
            {
                normalControl(textBox);
                return true;
            }
            else if (string.IsNullOrEmpty(textBox.Text) || string.IsNullOrWhiteSpace(textBox.Text))
            {
                if (!onlyReturn)
                    errorControl(textBox, "فیلد نباید خالی باشد");
                return false;
            }
            else
            {
                if (!onlyReturn)
                    errorControl(textBox, "حداقل 3 حرف وارد کنید");
                return false;
            }
        }
        #endregion

        #region Functions

        internal void resetControls()
        {
            #region variables
            foreach (CreateDocumentControlModel controlModel in textBoxControlModels)
            {
                controlModel.Validate = false;
                if (controlModel.Control == txtBoxEmail)
                {
                    controlModel.ControlLevel = ControlLevels.Optional;
                }
            }

            FieldOfStudyFa = "";
            FieldOfStudyEn = "";
            AreaOfStudyFa = "";
            AreaOfStudyEn = "";
            Email = "";
            IsResponsibleAuthor = false;

            // ====== این خط رو اضافه کن ======
            AuthorNames = new List<string>();

            // پاک کردن لیست نویسندگان
            authorsList.Clear();
            RefreshDataGrid();
            #endregion

            #region controls
            Dispatcher.Invoke(() =>
            {
                btnForward.IsEnabled = false;

                foreach (CreateDocumentControlModel controlModel in textBoxControlModels)
                {
                    TextBox textBox = (TextBox)controlModel.Control;
                    resetTextBoxControl(textBox);
                }

                txtBoxFieldOfStudy.Text = "";
                txtBoxFieldOfStudyEn.Text = "";
                txtBoxAreaOfStudy.Text = "";
                txtBoxAreaOfStudyEn.Text = "";
                txtBoxEmail.Text = "";
                txtBoxEmail.IsEnabled = false;
                chkResponsibleAuthor.IsChecked = false;
            });
            #endregion
        }

        private void resetTextBoxControl(TextBox control)
        {
            control.Text = "";
            normalControl(control);
        }

        private void errorControl(Control control, string hintText)
        {
            HintAssist.SetHelperText(control, hintText);
            control.Foreground = System.Windows.Media.Brushes.Red;

            Thickness margin = new Thickness(0, 0, 0, 20);
            if (control == txtBoxFieldOfStudy || control == txtBoxFieldOfStudyEn)
                gridFieldOfStudy.Margin = margin;
            else if (control == txtBoxAreaOfStudy || control == txtBoxAreaOfStudyEn)
                gridAreaOfStudy.Margin = margin;
            else if (control == txtBoxEmail)
                gridEmail.Margin = margin;
        }

        private void normalControl(Control control)
        {
            HintAssist.SetHelperText(control, "");
            control.Foreground = System.Windows.Media.Brushes.Black;

            Thickness margin = new Thickness(0);
            if (control == txtBoxFieldOfStudy && txtBoxFieldOfStudyEn.Foreground != System.Windows.Media.Brushes.Red)
                gridFieldOfStudy.Margin = margin;
            else if (control == txtBoxFieldOfStudyEn && txtBoxFieldOfStudy.Foreground != System.Windows.Media.Brushes.Red)
                gridFieldOfStudy.Margin = margin;
            else if (control == txtBoxAreaOfStudy && txtBoxAreaOfStudyEn.Foreground != System.Windows.Media.Brushes.Red)
                gridAreaOfStudy.Margin = margin;
            else if (control == txtBoxAreaOfStudyEn && txtBoxAreaOfStudy.Foreground != System.Windows.Media.Brushes.Red)
                gridAreaOfStudy.Margin = margin;
            else if (control == txtBoxEmail)
                gridEmail.Margin = margin;
        }

        internal void initializeVariables(DocumentTypes documentType)
        {
            this.documentType = documentType;
            validateControls();
        }

        #endregion
    }
}

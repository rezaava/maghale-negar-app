using MaghaleNegar.Forms.MaghaleNegarManager.DocumentManager.View;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Threading;

namespace MaghaleNegar.Forms.MaghaleNegarManager.CreateDocument
{
    public partial class CreateDocumentControl : System.Windows.Controls.UserControl, Interfaces.IStatusFormRequest, Interfaces.IChangeTransitionDocumentManager, INotifyPropertyChanged
    {
        //implement interface
        public Action TransitionDocumentManagerRequest { get; set; }

        public Action CloseFormRequest { get; set; }
        public Action AlwaysOnTopEnableRequest { get; set; }
        public Action AlwaysOnTopDisableRequest { get; set; }
        public Action MinimizeStateFormRequest { get; set; }
        public Action NormalStateFormRequest { get; set; }
        public Action MaximizeStateFormRequest { get; set; }

        private int previousSelectedTransition = 0;
        private bool isTransitionMovementForward = false;

        private int m_progress;
        public int Progress
        {
            get { return m_progress; }
            set
            {
                m_progress = value;
                OnPropertyChanged("Progress");
            }
        }

        public ObservableCollection<string> Steps
        {
            get;
            set;
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        public CreateDocumentControl()
        {
            InitializeComponent();

            Steps = new ObservableCollection<string>();

            // ====== سه مرحله ======
            Steps.Add("مشخصات مقاله");    // ایندکس 0 = اسلاید 3
            Steps.Add("وابستگی علمی");    // ایندکس 1 = اسلاید 5
            Steps.Add("ساخت مقاله");      // ایندکس 2 = اسلاید 6

            Progress = 33;
            DataContext = this;

            btnExitCreateDocument.Click += BtnExitCreateDocument_Click;
            toggleButtonAlwaysOnTop.Click += ToggleButtonAlwaysOnTop_Click;
            btnMinimize.Click += (b1, e) => { MinimizeStateFormRequest?.Invoke(); };
            btnMaximize.Click += (b1, e) => { MaximizeStateFormRequest?.Invoke(); };

            transitionCreateDocument.SelectionChanged += TransitionCreateDocument_SelectionChanged;

            // ====== رفتن به اسلاید اول ======
            transitionCreateDocument.SelectedIndex = 0;

            // ====== رویداد دکمه بعدی در اسلاید 3 (رفتن به اسلاید 5) ======
            createDocumentSlide3.TransitionMoveNextCommand += () =>
            {
                transitionCreateDocument.SelectedIndex = 1;
            };

            // ====== رویداد دکمه بعدی در اسلاید 5 (رفتن به اسلاید 6) ======
            createDocumentSlide5.TransitionMoveNextCommand += () =>
            {
                try
                {
                    // ====== گرفتن اطلاعات از اسلاید 5 ======
                    var infoList = createDocumentSlide5.GetInfoList();
                    var titleFa = createDocumentSlide5.GetTitleFa();
                    var titleEn = createDocumentSlide5.GetTitleEn();
                    var universityType = createDocumentSlide5.GetUniversityType();
                    var academicDegreeFa = createDocumentSlide5.GetAcademicDegreeFa();
                    var groupFa = createDocumentSlide5.GetGroupFa();
                    var facultyFa = createDocumentSlide5.GetFacultyFa();
                    var universityFa = createDocumentSlide5.GetUniversityFa();
                    var cityFa = createDocumentSlide5.GetCityFa();
                    var previewText = createDocumentSlide5.GetPreviewText();
                    var titleEnForFile = createDocumentSlide5.GetTitleEnForFile();
                    string documentName = createDocumentSlide3.DocumentName ?? "";

                    // ====== انتقال به اسلاید 6 ======
                    createDocumentSlide6.initializeVariables(
                        infoList,
                        titleFa,
                        titleEn,
                        universityType,
                        academicDegreeFa,
                        groupFa,
                        facultyFa,
                        universityFa,
                        cityFa,
                        previewText,
                        titleEnForFile,
                        documentName
                    );

                    transitionCreateDocument.SelectedIndex = 2;
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show($"خطا در انتقال به تأیید نهایی: {ex.Message}",
                        "خطا", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                }
            };

            // ====== رویدادهای اسلاید 6 ======
            createDocumentSlide6.TransitionDocumentManagerRequest += () =>
            {
                createDocumentSlide3.resetControls();
                createDocumentSlide5.resetControls();
                createDocumentSlide6.resetControls();
                Progress = 33;

                Dispatcher.Invoke(() =>
                {
                    transitionCreateDocument.SelectedIndex = 0;
                });
                this.TransitionDocumentManagerRequest?.Invoke();
            };

            createDocumentSlide6.CloseForm += () =>
            {
                CloseFormRequest?.Invoke();
            };

            string versionCustomized = BugReport.AssemblyVersion;
            versionCustomized = versionCustomized.Replace("0", "۰").Replace("1", "۱").Replace("2", "۲").Replace("3", "۳").Replace("4", "۴").Replace("5", "۵")
                .Replace("6", "۶").Replace("7", "۷").Replace("8", "۸").Replace("9", "۹");

            lblVersion.Text = "نسخه: " + versionCustomized;
        }

        private void ToggleButtonAlwaysOnTop_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            if (toggleButtonAlwaysOnTop.IsChecked == true)
                AlwaysOnTopEnableRequest?.Invoke();
            else
                AlwaysOnTopDisableRequest?.Invoke();
        }

        private void BtnExitCreateDocument_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            // ====== اگر در اسلاید 5 هستیم ======
            if (transitionCreateDocument.SelectedIndex == 1)
            {
                createDocumentSlide5.close();
                transitionCreateDocument.SelectedIndex = 0;
                return;
            }

            // ====== اگر در اسلاید 6 هستیم ======
            if (transitionCreateDocument.SelectedIndex == 2)
            {
                createDocumentSlide6.close();
                transitionCreateDocument.SelectedIndex = 0;
                return;
            }

            // ====== اسلاید 3 ======
            createDocumentSlide3.resetControls();
            createDocumentSlide5.resetControls();
            createDocumentSlide6.resetControls();
            Progress = 33;

            transitionCreateDocument.SelectedIndex = 0;
            TransitionDocumentManagerRequest?.Invoke();
        }

        private void BtnMaximize_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            // ====== پیدا کردن فرم والد ======
            Window parentWindow = Window.GetWindow(this);
            if (parentWindow != null)
            {
                if (parentWindow.WindowState == WindowState.Normal)
                    parentWindow.WindowState = WindowState.Maximized;
                else
                    parentWindow.WindowState = WindowState.Normal;
            }
        }

        private void BtnMinimize_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            Window parentWindow = Window.GetWindow(this);
            if (parentWindow != null)
            {
                parentWindow.WindowState = WindowState.Minimized;
            }
        }

        private void TransitionCreateDocument_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (transitionCreateDocument.SelectedIndex != previousSelectedTransition)
            {
                if (transitionCreateDocument.SelectedIndex > previousSelectedTransition)
                    isTransitionMovementForward = true;
                else
                    isTransitionMovementForward = false;
            }
            else
                return;

            // ====== محاسبه Progress برای 3 مرحله ======
            float step = 100f / 3f;
            Progress = (int)(step * (transitionCreateDocument.SelectedIndex + 1));

            // ====== وقتی به اسلاید 5 می‌رویم (وابستگی علمی) ======
            if (transitionCreateDocument.SelectedIndex == 1)
            {
                try
                {
                    string titleFa = createDocumentSlide3.FieldOfStudyFa ?? "";
                    string titleEn = createDocumentSlide3.FieldOfStudyEn ?? "";
                    var authorNames = createDocumentSlide3.AuthorNames ?? new List<string>();
                    var authorNamesEn = createDocumentSlide3.AuthorNamesEn ?? new List<string>();

                    createDocumentSlide5.initializeVariables(
                        authorNames,
                        authorNamesEn,
                        titleEn,
                        titleFa
                    );
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show($"خطا در انتقال اطلاعات به اسلاید وابستگی علمی: {ex.Message}",
                        "خطا", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                }
            }

            previousSelectedTransition = transitionCreateDocument.SelectedIndex;
        }
    }
}
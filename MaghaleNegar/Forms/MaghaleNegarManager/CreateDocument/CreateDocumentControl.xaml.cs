using MaghaleNegar.Forms.MaghaleNegarManager.DocumentManager.View;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
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

            // ====== سه مرحله (مرحله سوم فقط برای نمایش در نوار بالا) ======
            Steps.Add("مشخصات مقاله");    
            Steps.Add("وابستگی علمی");    
            Steps.Add("ساخت مقاله");           

            Progress = 1;
            DataContext = this;

            btnExitCreateDocument.Click += BtnExitCreateDocument_Click;
            toggleButtonAlwaysOnTop.Click += ToggleButtonAlwaysOnTop_Click;
            btnMinimize.Click += (b1, e) => { MinimizeStateFormRequest?.Invoke(); };
            btnMaximize.Click += (b1, e) => { MaximizeStateFormRequest?.Invoke(); };

            transitionCreateDocument.SelectionChanged += TransitionCreateDocument_SelectionChanged;

            // ====== رفتن به اسلاید اول ======
            transitionCreateDocument.SelectedIndex = 0;

            // ====== رویداد دکمه بعدی در اسلاید 3 (رفتن به اسلاید 6) ======
            createDocumentSlide3.TransitionMoveNextCommand += () =>
            {
                transitionCreateDocument.SelectedIndex = 1;
            };

            // ====== رویدادهای اسلاید 6 ======
            createDocumentSlide6.TransitionDocumentManagerRequest += () =>
            {
                createDocumentSlide3.resetControls();
                createDocumentSlide6.resetControls();
                Progress = 1;

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
            if (transitionCreateDocument.SelectedIndex == 1)
            {
                createDocumentSlide6.close();
                return;
            }

            createDocumentSlide3.resetControls();
            createDocumentSlide6.resetControls();
            Progress = 1;

            transitionCreateDocument.SelectedIndex = 0;
            TransitionDocumentManagerRequest?.Invoke();
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

            float step = 100f / (float)transitionCreateDocument.Items.Count;
            Progress = (int)step * (transitionCreateDocument.SelectedIndex + 1);

            // ====== وقتی به اسلاید 6 می‌رویم (وابستگی علمی) ======
            if (transitionCreateDocument.SelectedIndex == 1)
            {
                try
                {
                    string titleFa = createDocumentSlide3.FieldOfStudyFa ?? "";
                    string titleEn = createDocumentSlide3.FieldOfStudyEn ?? "";
                    var authorNames = createDocumentSlide3.AuthorNames ?? new System.Collections.Generic.List<string>();
                    var authorNamesEn = createDocumentSlide3.AuthorNamesEn ?? new List<string>();


                    createDocumentSlide6.initializeVariables(
                        authorNames,
                        authorNamesEn,
                        titleEn,
                        titleFa
                    );
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show($"خطا در انتقال اطلاعات به اسلاید نهایی: {ex.Message}",
                        "خطا", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                }
            }

            previousSelectedTransition = transitionCreateDocument.SelectedIndex;
        }
    }
}
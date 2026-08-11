using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Office.Interop.Word;

namespace MaghaleNegar.Forms.JournalSettings
{
    public partial class JournalSettingsControl : UserControl, Interfaces.IStatusFormRequest
    {
        Document doc;
        public Action CloseFormRequest { get; set; }
        public Action NormalStateFormRequest { get; set; }
        public Action MinimizeStateFormRequest { get; set; }
        public Action MaximizeStateFormRequest { get; set; }
        public Action AlwaysOnTopEnableRequest { get; set; }
        public Action AlwaysOnTopDisableRequest { get; set; }

        public JournalSettingsControl(Document doc)
        {
            InitializeComponent();

            this.doc = doc;

            btnCloseApp.Click += BtnCloseApp_Click;
            btnConfirm.Click += BtnConfirm_Click;
        }

        private void BtnConfirm_Click(object sender, RoutedEventArgs e)
        {
            CloseFormRequest?.Invoke();
        }

        private void BtnCloseApp_Click(object sender, RoutedEventArgs e)
        {
            CloseFormRequest?.Invoke();
        }

        public static T FindChild<T>(DependencyObject parent, string childName)
           where T : DependencyObject
        {
            if (parent == null) return null;

            T foundChild = null;
            int childrenCount = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < childrenCount; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                T childType = child as T;
                if (childType == null)
                {
                    foundChild = FindChild<T>(child, childName);
                    if (foundChild != null) break;
                }
                else if (!string.IsNullOrEmpty(childName))
                {
                    var frameworkElement = child as FrameworkElement;
                    if (frameworkElement != null && frameworkElement.Name == childName)
                    {
                        foundChild = (T)child;
                        break;
                    }
                }
                else
                {
                    foundChild = (T)child;
                    break;
                }
            }
            return foundChild;
        }
    }
}
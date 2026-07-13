using System;
using MaghaleNegar.Forms.MaghaleNegarManager.DocumentManager.Utilities;

namespace MaghaleNegar.Forms.MaghaleNegarManager.DocumentManager.ViewModel
{
    public class MainVM : ViewModelBase, IGoToDocumentManager
    {
        public Action GoToMain { get; set; }
        public Action GoToDocuments { get; set; }
        public Action GoToLogin { get; set; }

        public MainVM()
        {
        }
    }
}

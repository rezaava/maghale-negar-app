using Microsoft.Office.Interop.Word;

namespace MaghaleNegar.Forms.JournalSettings
{
    partial class JournalSettingsForm
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent(Document doc)
        {
            this.components = new System.ComponentModel.Container();
            this.timerOpenAnimation = new System.Windows.Forms.Timer(this.components);
            this.elementHost1 = new System.Windows.Forms.Integration.ElementHost();
            this.journalSettingsControl = new MaghaleNegar.Forms.JournalSettings.JournalSettingsControl(doc);
            this.SuspendLayout();

            this.timerOpenAnimation.Interval = 1;
            this.timerOpenAnimation.Tick += new System.EventHandler(this.timerOpenAnimation_Tick);

            this.elementHost1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.elementHost1.Location = new System.Drawing.Point(0, 0);
            this.elementHost1.Name = "elementHost1";
            this.elementHost1.Size = new System.Drawing.Size(800, 600);
            this.elementHost1.TabIndex = 0;
            this.elementHost1.Text = "elementHost1";
            this.elementHost1.Child = this.journalSettingsControl;

            this.ClientSize = new System.Drawing.Size(800, 600);
            this.ControlBox = false;
            this.Controls.Add(this.elementHost1);
            this.DoubleBuffered = true;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "JournalSettingsForm";
            this.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.ShowIcon = false;
            this.Text = "JournalSettingsForm";
            this.ResumeLayout(false);
        }

        private System.Windows.Forms.Timer timerOpenAnimation;
        private System.Windows.Forms.Integration.ElementHost elementHost1;
        public JournalSettingsControl journalSettingsControl;
    }
}
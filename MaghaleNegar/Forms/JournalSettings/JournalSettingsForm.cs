using MaghaleNegar.Constants;
using System;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace MaghaleNegar.Forms.JournalSettings
{
    public partial class JournalSettingsForm : Form
    {
        [Flags]
        enum AnimateWindowFlags
        {
            AW_HOR_POSITIVE = 0x0000000,
            AW_HOR_NEGATIVE = 0x00000002,
            AW_VER_POSITIVE = 0x00000004,
            AW_VER_NEGATIVE = 0x00000008,
            AW_CENTER = 0x00000010,
            AW_HIDE = 0x00010000,
            AW_ACTIVATE = 0x00020000,
            AW_SLIDE = 0x00040000,
            AW_BLEND = 0x00080000
        }

        [DllImport("user32.dll")]
        static extern bool AnimateWindow(IntPtr hWnd, int time, AnimateWindowFlags flags);

        private int formBorderRadius = 10;
        private int formBorderSize = 5;
        private Color formBorderColor = ColorsApp.PrimaryColor;

        public JournalSettingsForm(Microsoft.Office.Interop.Word.Document doc)
        {
            InitializeComponent(doc);

            DedicatedFunctions.setManualScale(this, new Size(370, 600), new Size(300, 485));
            this.CenterToScreen();

            this.FormClosing += (e, a) =>
            {
                AnimateWindow(this.Handle, 100, AnimateWindowFlags.AW_BLEND | AnimateWindowFlags.AW_HIDE);
                Globals.ThisAddIn.DocumentManagerFormVisible = false;
            };

            #region Form inital settings
            this.Padding = new Padding(formBorderSize);
            this.BackColor = formBorderColor;
            this.Opacity = 0;
            setDoubleBuffer(elementHost1, true);
            this.SetStyle(ControlStyles.ResizeRedraw, true);
            this.Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(1, 1, Width, Height, formBorderRadius * 2, formBorderRadius * 2));
            timerOpenAnimation.Enabled = true;
            #endregion

            journalSettingsControl.CloseFormRequest += () =>
            {
                closeForm();
            };
            journalSettingsControl.MaximizeStateFormRequest += () =>
            {
                if (this.WindowState == FormWindowState.Normal)
                    this.WindowState = FormWindowState.Maximized;
                else if (this.WindowState == FormWindowState.Maximized)
                    this.WindowState = FormWindowState.Normal;
            };
            journalSettingsControl.MinimizeStateFormRequest += () =>
            {
                this.WindowState = FormWindowState.Minimized;
            };

            journalSettingsControl.MouseDown += JournalSettingsControl_MouseDown;
        }

        private void closeForm()
        {
            GC.Collect();
            this.Close();
        }

        private void JournalSettingsControl_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            releaseCapture();
            sendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        #region (UI) WindowsForm

        static void setDoubleBuffer(Control ctrl, bool doubleBuffered)
        {
            try
            {
                typeof(Control).InvokeMember("DoubleBuffered", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.SetProperty, null, ctrl, new object[] { doubleBuffered });
            }
            catch (Exception)
            {
            }
        }

        #region Resizable
        int Thickness = 10;
        int Area = 8;
        private bool Above, right, Under, left, Right_above, Right_under, Left_under, Left_above;
        public string getMosuePosition(Point mouse, Form form)
        {
            bool above_underArea = mouse.X > Area && mouse.X < form.ClientRectangle.Width - Area;
            bool right_left_Area = mouse.Y > Area && mouse.Y < form.ClientRectangle.Height - Area;

            bool _Above = mouse.Y <= Thickness;
            bool _Right = mouse.X >= form.ClientRectangle.Width - Thickness;
            bool _Under = mouse.Y >= form.ClientRectangle.Height - Thickness;
            bool _Left = mouse.X <= Thickness;

            Above = _Above && (above_underArea); if (Above) return "a";
            right = _Right && (right_left_Area); if (right) return "r";
            Under = _Under && (above_underArea); if (Under) return "u";
            left = _Left && (right_left_Area); if (left) return "l";

            Right_above = (_Right && (!right_left_Area)) && (_Above && (!above_underArea)); if (Right_above) return "ra";
            Right_under = ((_Right) && (!right_left_Area)) && (_Under && (!above_underArea)); if (Right_under) return "ru";
            Left_under = ((_Left) && (!right_left_Area)) && (_Under && (!above_underArea)); if (Left_under) return "lu";
            Left_above = ((_Left) && (!right_left_Area)) && (_Above && (!above_underArea)); if (Left_above) return "la";

            return "";
        }

        protected override void WndProc(ref Message m)
        {
            int x = (int)(m.LParam.ToInt64() & 0xFFFF);
            int y = (int)((m.LParam.ToInt64() & 0xFFFF0000) >> 16);
            Point pt = PointToClient(new Point(x, y));

            if (m.Msg == 0x84)
            {
                switch (getMosuePosition(pt, this))
                {
                    case "l": m.Result = (IntPtr)10; return;
                    case "r": m.Result = (IntPtr)11; return;
                    case "a": m.Result = (IntPtr)12; return;
                    case "la": m.Result = (IntPtr)13; return;
                    case "ra": m.Result = (IntPtr)14; return;
                    case "u": m.Result = (IntPtr)15; return;
                    case "lu": m.Result = (IntPtr)16; return;
                    case "ru": m.Result = (IntPtr)17; return;
                    case "": m.Result = pt.Y < 32 ? (IntPtr)2 : (IntPtr)1; return;
                }
            }
            base.WndProc(ref m);
        }
        #endregion

        #region Draggable 
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void releaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void sendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);
        #endregion

        #region Round Corner Windows Form
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            if (this.WindowState != FormWindowState.Minimized)
                this.Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(1, 1, Width, Height, formBorderRadius * 2, formBorderRadius * 2));
        }

        [DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]
        private static extern IntPtr CreateRoundRectRgn(int nLeftRect, int nTopRect, int nRightRect, int nBottomRect, int nWidthEllipse, int nHeightEllipse);
        #endregion

        #endregion

        private void timerOpenAnimation_Tick(object sender, EventArgs e)
        {
            Opacity += 0.1;
            if (Opacity == 1)
            {
                timerOpenAnimation.Enabled = false;
            }
        }
    }
}
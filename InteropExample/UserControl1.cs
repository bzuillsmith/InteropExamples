using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.Windows.Forms;

namespace InteropExample
{
    #region Interop Interfaces
    /// <summary>
    /// This interface is used as the COM Source interface for the UserControl1 class.
    /// </summary>
    /// <remarks>
    /// <para>The COM source interface allows for use of the COM connection points protocol.</para>
    /// <para>All events that this control needs to expose should be defined in this
    /// interface as VB6 only supports its WithEvents syntax for a single interface.</para>
    /// <para>Each method defined in this interface must match up to an event in the user
    /// control having the same name; the method signatures here must match the signature of
    /// the corresponding event's delegate.</para>
    /// <para>Interface is defined as a dispinterface (IDispatch) because VB6 requires it
    /// for source interfaces.</para>
    /// <para>Each method must be marked with a unique DispId value greater than 0. Without proper
    /// DispIds, raising an event may cause a COMException to be thrown if the VB6 client does not
    /// handle all defined events.</para>
    /// </remarks>
    [Guid(UserControl1.EventsId), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface UserControl1Events
    {
        [DispId(1)]
        void Click();
        [DispId(2)]
        void DblClick();
        [DispId(3)]
        void ToolStripItemClicked(string itemName);
    }

    /// <summary>
    /// This is the default interface implemented by the user control, and should
    /// contain all the methods and properties that will be exposed to COM.
    /// </summary>
    [Guid(UserControl1.InterfaceId)]
    public interface IUserControl1
    {
        /// <summary>
        /// Gets or sets a value indicating whether the user control is visible.
        /// </summary>
        [DispId(1)]
        bool Visible { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the user control is enabled.
        /// </summary>
        [DispId(2)]
        bool Enabled { get; set; }

        /// <summary>
        /// Gets or sets the foreground color of the user control.
        /// </summary>
        [DispId(3)]
        int ForegroundColor { get; set; }

        /// <summary>
        /// Gets or sets the background color of the user control.
        /// </summary>
        [DispId(4)]
        int BackgroundColor { get; set; }

        /// <summary>
        /// Gets or sets the background image of the user control.
        /// </summary>
        [DispId(5)]
        Image BackgroundImage { get; set; }

        /// <summary>
        /// Forces the control to invalidate its client area and immediately redraw 
        /// itself and any child controls.
        /// </summary>
        [DispId(6)]
        void Refresh();
    }
#endregion

    [Guid(ClassId), ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(UserControl1Events))]
    public partial class UserControl1: UserControl, IUserControl1
    {
        public UserControl1()
        {
            // This call is required by the Windows Form Designer.
            InitializeComponent();

            // Add any initialization after the InitializeComponent() call.
            this.DoubleClick += (this.UserControl1_DblClick);
            base.Click += (this.UserControl1_Click);
            this.LostFocus += (UserControl1_LostFocus);
            this.ControlAdded += (UserControl1_ControlAdded);
            //this.Load += (UserControl1_Load);

            // Raise custom Load event
            this.OnCreateControl();
        }


        #region VB6 Interop Code

        #region "COM Registration"

        // These  GUIDs provide the COM identity for this class 
        // and its COM interfaces. If you change them, existing 
        // clients will no longer be able to access the class.

        public const string ClassId = "534BE188-9376-4611-AECA-887958F8E1B6";
        public const string InterfaceId = "34595C50-B461-4600-A69D-3C62709CD833";
        public const string EventsId = "BFD16739-96E4-4E51-9598-F7D5220B3E98";

        // These routines perform the additional COM registration needed by ActiveX controls
        [EditorBrowsable(EditorBrowsableState.Never)]
        [ComRegisterFunction]
        private static void Register(System.Type t)
        {
            ComRegistration.RegisterControl(t);
        }

        [EditorBrowsable(EditorBrowsableState.Never)]
        [ComUnregisterFunction]
        private static void Unregister(System.Type t)
        {
            ComRegistration.UnregisterControl(t);
        }


        #endregion

        #region "VB6 Events"

        // This section shows some examples of exposing a UserControl's events to VB6.  Typically, you just
        // 1) Declare the event as you want it to be shown in VB6
        // 2) Raise the event in the appropriate UserControl event.
        public delegate void ClickEventHandler();
        public delegate void DblClickEventHandler();
        public delegate void ToolStripItemClickedEventHandler(string itemName);
        public new event ClickEventHandler Click; // Event must be marked as new since .NET UserControls have the same name.
        public event DblClickEventHandler DblClick;
        public event ToolStripItemClickedEventHandler ToolStripItemClicked;

        private void UserControl1_Click(object sender, System.EventArgs e)
        {
            if (null != Click)
                Click();
        }

        private void UserControl1_DblClick(object sender, System.EventArgs e)
        {
            if (null != DblClick)
                DblClick();
        }


        #endregion

        #region "VB6 Properties"

        // The following are examples of how to expose typical form properties to VB6.  
        // You can also use these as examples on how to add additional properties.

        // Must declare this property as new as it exists in Windows.Forms and is not overridable
        public new bool Visible
        {
            get { return base.Visible; }
            set { base.Visible = value; }
        }

        public new bool Enabled
        {
            get { return base.Enabled; }
            set { base.Enabled = value; }
        }

        public int ForegroundColor
        {
            get
            {
                return ActiveXControlHelpers.GetOleColorFromColor(base.ForeColor);
            }
            set
            {
                base.ForeColor = ActiveXControlHelpers.GetColorFromOleColor(value);
            }
        }

        public int BackgroundColor
        {
            get
            {
                return ActiveXControlHelpers.GetOleColorFromColor(base.BackColor);
            }
            set
            {
                base.BackColor = ActiveXControlHelpers.GetColorFromOleColor(value);
            }
        }

        public override Image BackgroundImage
        {
            get { return null; }
            set
            {
                if (null != value)
                {
                    MessageBox.Show("Setting the background image of an Interop UserControl is not supported, please use a PictureBox instead.", "Information");
                }
                base.BackgroundImage = null;
            }
        }

        #endregion

        #region "VB6 Methods"

        // Ensures that tabbing across VB6 and .NET controls works as expected
        private void UserControl1_LostFocus(object sender, System.EventArgs e)
        {
            Debug.Print("USerControl1_LostFocus");
            ActiveXControlHelpers.HandleFocus(this);
        }


        [SecurityPermission(SecurityAction.LinkDemand, Flags = SecurityPermissionFlag.UnmanagedCode)]
        protected override void WndProc(ref Message m)
        {
            const int WM_SETFOCUS = 0x7;
            const int WM_PARENTNOTIFY = 0x210;
            const int WM_DESTROY = 0x2;
            const int WM_LBUTTONDOWN = 0x201;
            const int WM_RBUTTONDOWN = 0x204;

            if (m.Msg == WM_SETFOCUS)
            {
                // Raise Enter event
                this.OnEnter(System.EventArgs.Empty);
            }
            else if (m.Msg == WM_PARENTNOTIFY && (m.WParam.ToInt32() == WM_LBUTTONDOWN || m.WParam.ToInt32() == WM_RBUTTONDOWN))
            {

                if (!this.ContainsFocus)
                {
                    // Raise Enter event
                    this.OnEnter(System.EventArgs.Empty);
                }
            }
            else if (m.Msg == WM_DESTROY && !this.IsDisposed && !this.Disposing)
            {
                // Used to ensure that VB6 will cleanup control properly
                this.Dispose();
            }

            base.WndProc(ref m);
        }



        // This event will hook up the necessary handlers
        private void UserControl1_ControlAdded(object sender, ControlEventArgs e)
        {
            ActiveXControlHelpers.WireUpHandlers(e.Control, ValidationHandler);
        }

        // Ensures that the Validating and Validated events fire appropriately
        internal void ValidationHandler(object sender, System.EventArgs e)
        {
            if (this.ContainsFocus) return;

            // Raise Leave event
            this.OnLeave(e);

            if (this.CausesValidation)
            {
                CancelEventArgs validationArgs = new CancelEventArgs();
                this.OnValidating(validationArgs);

                if (validationArgs.Cancel && this.ActiveControl != null)
                    this.ActiveControl.Focus();
                else
                {
                    // Raise Validated event
                    this.OnValidated(e);
                }
            }

        }

        #endregion

        #endregion


    }
}

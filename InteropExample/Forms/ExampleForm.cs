using Microsoft.InteropFormTools;
using System.Windows.Forms;

namespace InteropExample.Forms
{
    [InteropForm]
    public partial class ExampleForm : Form
    {
        public ExampleForm()
        {
            InitializeComponent();
        }

        private string _content;
        [InteropFormProperty]
        public string Content
        {
            get
            {
                return _content;
            }
            set
            {
                _content = value;
                this.textBox1.Text = _content;
            }
        }
    }
}

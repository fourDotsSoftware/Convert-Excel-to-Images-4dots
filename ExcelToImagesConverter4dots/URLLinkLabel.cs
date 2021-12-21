using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelToImagesConverter4dots
{
    public class URLLinkLabel : System.Windows.Forms.LinkLabel 
    {
        public URLLinkLabel()
            : base()
        {

        }

        protected override void OnClick(EventArgs e)
        {
            System.Diagnostics.Process.Start(this.Text);
        }
    }
}

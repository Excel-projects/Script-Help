using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SqlHelp.TaskPane
{
    public partial class DocumentViewer : UserControl
    {
        public DocumentViewer()
        {
            InitializeComponent();
            //string clickOnceLocation = Scripts.AssemblyInfo.GetClickOnceLocation();
            //this.ppcDocumentViewer.Document = Path.Combine(clickOnceLocation, @"Documentation\\As Built.pdf");
        }

    }
}

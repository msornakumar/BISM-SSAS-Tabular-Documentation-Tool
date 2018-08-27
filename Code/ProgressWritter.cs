using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
namespace BISMDocumenterLibrary
{
    class ProgressWritter
    {

      
       private string _InvokedAppType;

        public string InvokedAppType
        {
            get { return this._InvokedAppType; }
            set
            {
                this._InvokedAppType = value;
                
            }
        }
      
        public void WriteProgress(String ProgressText, object WriteProgressTo  )
        {
           if (this._InvokedAppType == "Windows")
            {
                System.Windows.Forms.TextBox ProgressTextBox = (System.Windows.Forms.TextBox) WriteProgressTo;
                ProgressText = ProgressText + Environment.NewLine;
                ProgressTextBox.AppendText(ProgressText);
           }
        }

    }
}

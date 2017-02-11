using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;

namespace BayantechAddIn
{
    public partial class BayantechRibbon
    {
        BidirectionalText Bidi;
        Word.Application app;
        //Word.Document doc;
        //Word.Range range;

        FixBidiForm frm_fix_bidi;
        private void BayantechRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            //Initialization
            app = Globals.ThisAddIn.Application;
            //range = doc.Range();
            Bidi = new BidirectionalText();
            //Track active document used
            app.DocumentChange += Application_DocumentChange;

            frm_fix_bidi = new FixBidiForm(Bidi);
        }

        void Application_DocumentChange()
        {
            if (app.Documents.Count > 0)
            {
                //Add all document objects that needs to be updated with active document
                Bidi.doc = app.ActiveDocument;
            }
        }

        private void btn_fix_bidi_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisAddIn.EXPIRED)
            {
                frm_fix_bidi.ShowDialog();
            }
            else
                Globals.ThisAddIn.showExpired();

            if (frm_fix_bidi.applied)
            {
                Globals.ThisAddIn.incrementLicense();
            }
        }

    }
}

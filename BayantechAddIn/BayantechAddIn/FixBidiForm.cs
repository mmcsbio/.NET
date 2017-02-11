using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Diagnostics;

namespace BayantechAddIn
{
    public partial class FixBidiForm : Form
    {
        public BidirectionalText Bidi;
        public bool applied = false;
        private Word.Application app = Globals.ThisAddIn.Application;
        private BayantechAddIn.Region region;

        public FixBidiForm(BidirectionalText Bidi)
        {
            InitializeComponent();

            region = BayantechAddIn.Region.Body;
            string[] regions = getRegions();
            cmb_region.Items.AddRange(regions);
            cmb_region.SelectedIndex = 0;

            this.Bidi = Bidi;
            applied = false;
        }

        private string[] getRegions()
        {
            return Enum.GetNames(typeof(BayantechAddIn.Region));
        }

        /// <summary>
        /// Get the region enum from selected string
        /// </summary>
        /// <param name="type">string to compare with enum types</param>
        /// <returns>Region Enum of item selected</returns>
        private BayantechAddIn.Region getRegionEnum(string type)
        {
            string item;
            try
            {
                item = type;
                region = (BayantechAddIn.Region)Enum.Parse(typeof(BayantechAddIn.Region), item);
            }
            catch
            {
                region = BayantechAddIn.Region.All;
            }
            return region;
        }

        private void btn_cancel_Click(object sender, EventArgs e)
        {
            Close();
            applied = false;
        }

        private void btn_apply_Click(object sender, EventArgs e)
        {

            btn_apply.Enabled = !btn_apply.Enabled;
            string regionItem = (string)cmb_region.SelectedItem;
            region = getRegionEnum(regionItem);

            Bidi.region = region;
            Bidi.fixBidiIssues();
            applied = true;

            Close();
            btn_apply.Enabled = !btn_apply.Enabled;
        }

    }
}

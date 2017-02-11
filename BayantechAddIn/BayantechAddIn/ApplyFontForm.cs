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
    public partial class ApplyFontForm : Form
    {
        public BidirectionalText Bidi;
        //Boolean to check if the user clicked "Apply" Button or not
        public bool applied = false;
        public bool applyOnStyles;
        private Word.Application app = Globals.ThisAddIn.Application;
        private string fontName;
        private Language language;
        private BayantechAddIn.Region region;

        public ApplyFontForm(BidirectionalText Bidi)
        {
            InitializeComponent();
            fontName = null;
            language = Language.All;
            
            //Font Name Combobox
            string[] fontNames = getFontNames();
            cmb_font_name.Items.AddRange(fontNames);
            cmb_font_name.SelectedIndex = 0;

            //Languages Combobox
            string[] languages = getLanguages();
            cmb_language.Items.AddRange(languages);
            cmb_language.SelectedIndex = 0;

            //Regions Combobox
            region = BayantechAddIn.Region.Body;
            string[] regions = getRegions();
            cmb_region.Items.AddRange(regions);
            cmb_region.SelectedIndex = 0;

            this.Bidi = Bidi;
            applyOnStyles = true;
            applied = false;
        }

        private string[] getFontNames()
        {
            string[] fontNames = new string[app.FontNames.Count - 1];
            for (int i = 1; i < app.FontNames.Count; i++)
            {
                fontNames[i - 1] = app.FontNames[i];
            }
            return fontNames;
        }

        private string[] getRegions()
        {
            return Enum.GetNames(typeof(BayantechAddIn.Region));
        }

        private string[] getLanguages()
        {
            return Enum.GetNames(typeof(Language));
        }

        /// <summary>
        /// Get the language enum from selected string
        /// </summary>
        /// <param name="type">string to compare with enum types</param>
        /// <returns>Language Enum item selected</returns>
        private Language getLanguageEnum(string type)
        {
            Language language;
            string item;
            try
            {
                item = type;
                language = (Language)Enum.Parse(typeof(Language), item);
            }
            catch
            {
                language = Language.All;
            }
            return language;
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
            this.Close();
            applied = false;
        }

        private void btn_apply_Click(object sender, EventArgs e)
        {
            //record time taken in processing
            Stopwatch watch = new Stopwatch();
            watch.Start();

            btn_apply.Enabled = !btn_apply.Enabled;
            string langItem = (string)cmb_language.SelectedItem;
            string fontItem = (string)cmb_font_name.SelectedItem;
            language = getLanguageEnum(langItem);
            fontName = fontItem;
            string regionItem = (string)cmb_region.SelectedItem;
            region = getRegionEnum(regionItem);

            Bidi.region = region;
            Bidi.applyFont(fontName, language, applyOnStyles);
            applied = true;

            watch.Stop();
            TimeSpan elapsed = watch.Elapsed;
            MessageBox.Show("Done Successfully!\n" + "Time Taken [" + elapsed.ToString() + "]", "Processing Completed");
            Close();
            btn_apply.Enabled = !btn_apply.Enabled;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace BayantechAddIn
{
    //part of the document to work on (ex: work on header only or shapes only in the document)
    public enum Region { Body, All, Selected, Shape, Header, Footer, Endnote, Footnote };
    public enum Language { All, Arabic, English };
    public class BidirectionalText
    {
        #region Characters Classification variables
        //LTR language character set
        private string ltrChars;
        //RTL language character set
        private string rtlChars;
        //Character set common in between RTL & LTR languages and have actual character unicode that is different in both languages 
        //(ex: number one in english unicode = U+0031, number one in arabic unicode = U+0661)
        private string commonChars;
        //Character set that is available in LTR language only(have unicode in LTR only), but is virtually available in RTL keyboard(no actual unicode in RTL)
        private string unmappedRtlChars;
        //Character set that is available in RTL language only(have unicode in RTL only), but is virtually available in LTR keyboard(no actual unicode in LTR)
        private string unmappedLtrChars;
        //RTL character set that are forced to be LTR characters
        private string ltrForcedChars;
        //LTR character set that are forced to be RTL characters
        private string rtlForcedChars;
        //Extra Unicode Blocks
        private string latinSupplement;
        private string latinExtA;
        private string latinExtB;
        private string currencySymbols;
        private string letterLikeSymbols;
        //Control Characters forced between text, recommended to be removed(disrubt finding process)
        private string generalPunctuation;
        #endregion

        #region Characters Matching Pattern Variables
        //The following are the patterns used to match the RTL & LTR chars variables
        public string ltrRegex;
        public string rtlRegex;
        public string commonRegex;
        public string unmappedLtrRegex;
        public string unmappedRtlRegex;
        public string rtlForcedRegex;
        public string ltrForcedRegex;
        #endregion

        public Region region;

        //Word Application and document to operate on
        public Word.Document doc;
        private Word.Application app;

        public BidirectionalText()
        {
            //Initialization
            //Old Regex used in old version(v1.0)
            //ltrRegex = "[a-zA-Z0-9 +\\-™©.§‏‎]{1,}";
            //rtlRegex = "[!a-zA-Z0-9 +\\-™©.§‏‎]{1,}";
            app = Globals.ThisAddIn.Application;
            doc = app.ActiveDocument;
            region = Region.All;

            #region Characters Classification Initialization
            ltrChars = "a-zA-Z`;'?";
            rtlChars = "لأأضصثقفغعهخحجدشسيبلاتنمكطئءؤرلاىةوزظذلآآلإإًًٌٍَُِّْ÷×؛ـ،’‘؟";
            commonChars = "1234567890";
            //the caret char(^) in word is searched by another caret(like this: ^^), thats why there is double carets in the characters below
            unmappedRtlChars = "\\!\\@#$%^^&\\*.+\\-=_~\\\\/:|, ";
            unmappedLtrChars = "";//no unmapped LTR Characters in English(LTR) language
            ltrForcedChars = "";//no forced LTR Characters in English(LTR) language
            rtlForcedChars = "}{][<>)(\"";
            latinSupplement = "-ÿ";// \u0080-\u00ff (Hidden chars exist)
            latinExtA = "Ā-ſ";// \u0100-\u017f
            latinExtB = "ƀ-ɏ";// \u0180-\u024f
            currencySymbols = "₠-⃏";// \u20a0-\u20cf
            letterLikeSymbols = "℀-⅏";// \u2100-\u214f
            generalPunctuation = " -⁯";// \u2000-\u206f (include Hidden chars left and right marks)
            #endregion

            #region Characters Matching Regex Initialization
            ltrRegex = addRegexEnds(ltrChars + commonChars + unmappedRtlChars + latinSupplement
                + latinExtA
                + latinExtB
                + currencySymbols
                + letterLikeSymbols
                + generalPunctuation);
            rtlRegex = addRegexEnds(rtlChars + rtlForcedChars);
            commonRegex = addRegexEnds(commonChars);
            unmappedLtrRegex = addRegexEnds(unmappedLtrChars);
            unmappedRtlRegex = addRegexEnds(unmappedRtlChars);
            rtlForcedRegex = addRegexEnds(rtlForcedChars);
            ltrForcedRegex = addRegexEnds(ltrForcedChars);
            #endregion

        }

        public BidirectionalText(string ltrRegex, string rtlRegex, Word.Document doc, Region region)
        {
            //Initialization
            app = Globals.ThisAddIn.Application;
            this.doc = doc;
            this.ltrRegex = ltrRegex;
            //currently left empty until needed
            this.rtlRegex = rtlRegex;
            this.region = region;
        }

        /// <summary>
        /// Surround Characters with regex pattern(find one or more occurrence)
        /// </summary>
        /// <param name="chars"></param>
        /// <returns></returns>
        public string addRegexEnds(string chars)
        {
            if(chars.Length > 0)
                return "[" + chars + "]{1,}";
            return "";
        }

        /// <summary>
        /// Get list of ranges of specified region property
        /// </summary>
        /// <returns></returns>
        private List<Word.Range> getRanges()
        {
            //Regions applied: All, selected, header, footer, endNote, footNote, body, shape
            List<Word.Range> ranges = new List<Word.Range>();

            if (region == Region.All || region == Region.Endnote)
            {
                Word.Range range;
                if (doc.Endnotes.Count > 0)
                {
                    ranges.Add(doc.Endnotes.Separator.Duplicate);
                    range = doc.Endnotes[1].Range;
                    int end = doc.Endnotes[doc.Endnotes.Count].Range.End;
                    range.SetRange(0, end);
                    ranges.Add(range.Duplicate);
                }
            }
            if (region == Region.All || region == Region.Footnote)
            {
                Word.Range range;
                if (doc.Footnotes.Count > 0)
                {
                    ranges.Add(doc.Footnotes.Separator.Duplicate);
                    range = doc.Footnotes[1].Range;
                    int end = doc.Footnotes[doc.Footnotes.Count].Range.End;
                    range.SetRange(0, end);
                    ranges.Add(range.Duplicate);
                }
            }
            if (region == Region.All || region == Region.Header)
            {
                foreach (Word.Section section in doc.Sections)
                {
                    foreach (Word.HeaderFooter item in section.Headers)
                    {
                        if (item.Exists)
                            ranges.Add(item.Range.Duplicate);
                    }
                }
            }
            if (region == Region.All || region == Region.Footer)
            {
                foreach (Word.Section section in doc.Sections)
                {
                    foreach (Word.HeaderFooter item in section.Footers)
                    {
                        if (item.Exists)
                            ranges.Add(item.Range.Duplicate);
                    }
                }
            }
            if (region == Region.All || region == Region.Shape)
                foreach (Word.Shape item in doc.Shapes)
                    if (item.TextFrame.HasText != 0)
                        ranges.Add(item.TextFrame.TextRange.Duplicate);

            if (region == Region.All || region == Region.Body)
                ranges.Add(doc.Content.Duplicate);

            if (region == Region.Selected)
                ranges.Add(app.Selection.Range.Duplicate);

            return ranges;
        }

        /// <summary>
        /// Deselect fullstop(sentence period) from specified range
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        private void excludeFullstop(Word.Range range)
        {
            bool empty = range.Start == range.End;
            if (!empty)
            {
                Word.Characters rngChars = range.Characters;
                int length = rngChars.Count;
                if (length > 0)
                {
                    //Note: Word DOM arrays are 1 based index
                    //keep track of the original end of the range so when it is changed can return to it later
                    int rngEnd = range.End;
                    string trimSymbol = ".";
                    //while last char value is a "dot"(sentence period), relocate the range end to be the end of the before last character and remove the last character
                    while (length > 1 && rngChars[length].Text.Equals(trimSymbol))
                    {
                        //[Substituted with below] range.End = range.End - 1;
                        //change indices by actual characters indices
                        range.End = rngChars[length - 1].End;
                        length--;
                    }
                    //start and end the range at the original end (rngEnd) of the range (so the find function can select a new range)
                    if (length == 1 && rngChars[length].Text.Equals(trimSymbol))
                        range.SetRange(rngEnd, rngEnd);
                }
            }
        }

        /// <summary>
        /// Deselect traling spaces(extra spaces at both ends) from specified range
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        private void excludeSpace(Word.Range range)
        {
            bool empty = range.Start == range.End;
            if (!empty)
            {
                Word.Characters rngChars = range.Characters;
                int length = rngChars.Count;
                if (length > 0)
                {
                    //Note: Word DOM arrays are 1 based index
                    //keep track of the original end of the range so when it is changed can return to it later
                    int rngEnd = range.End;
                    string trimSymbol = " ";

                    //while last char value is a "space", relocate the range end to be the end of the before last character and remove the last character
                    while (length > 1 && rngChars[length].Text.Equals(trimSymbol))
                    {
                        //[Substituted with below] range.End = range.End - 1;
                        range.End = rngChars[length - 1].End;
                        length--;
                    }
                    //while first char value is a "space", relocate the range start to be the start of the second character and remove the first character
                    while (length > 1 && rngChars[1].Text.Equals(trimSymbol))
                    {
                        //[Substituted with below] range.Start = range.Start + 1;
                        range.Start = rngChars[2].Start;
                        length--;
                    }
                    //start and end the range at the original end (rngEnd) of the range (so the find function can select a new range)
                    if (length == 1 && rngChars[length].Text.Equals(trimSymbol))
                        range.SetRange(rngEnd, rngEnd);
                }
            }
        }

        /// <summary>
        /// Set Paragraphs Reading order from Right to Left on specified range
        /// </summary>
        /// <param name="range">Document Range</param>
        private void paraRtl(Word.Range range)
        {
            bool empty = range.Start == range.End;
            if (!empty)
            {
                range.Select();
                app.Selection.RtlPara();
                //range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            }
        }

        /// <summary>
        /// Set Paragraphs Reading order from Left to Right on specified range
        /// </summary>
        /// <param name="range">Document Range</param>
        private void paraLtr(Word.Range range)
        {
            bool empty = range.Start == range.End;
            if (!empty)
            {
                range.Select();
                app.Selection.LtrPara();
                //range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            }
        }

        ///<summary>
        /// Set characters Reading order from Right to Left on selected range
        /// </summary>
        /// <param name="range">Document range</param>
        private void runRtl(Word.Range range)
        {
            bool empty = range.Start == range.End;
            if (!empty)
            {
                range.Select();
                app.Selection.RtlRun();
            }
        }

        ///<summary>
        /// Set Reading order from Left to Right on selected range
        /// </summary>
        /// <param name="range">Document range</param>
        private void runLtr(Word.Range range)
        {
            bool empty = range.Start == range.End;
            if (!empty)
            {
                range.Select();
                app.Selection.LtrRun();
            }
        }

        /// <summary>
        /// Apply "Run Ltr" command on english regex(ltrRegex) in specified range
        /// </summary>
        /// <param name="range"></param>
        /// <returns>True if successful and False if error happens</returns>
        private void runLtrRegex(Word.Range range, Word.WdColorIndex highlightColor = Word.WdColorIndex.wdNoHighlight)
        {
            bool empty = range.Start == range.End;
            if (!empty)
            {
                List<Word.Range> ranges = findRegex(ltrRegex, range);
                foreach (Word.Range item in ranges)
                {
                    //do not change the following two functions order, exclude space first then fullstop(sentence period)
                    excludeSpace(item);
                    excludeFullstop(item);

                    empty = item.Start == item.End;
                    if (!empty)
                    {
                        item.HighlightColorIndex = highlightColor;
                        item.Select();
                        app.Selection.LtrRun();
                    }
                }
            }
        }

        /// <summary>
        /// Reserved for the new approach to fix bidi used in fixBidi function
        /// </summary>
        /// <param name="range"></param>
        /// <param name="highlightColor"></param>
        private void runLtrRegex_new(Word.Range range, Word.WdColorIndex highlightColor = Word.WdColorIndex.wdNoHighlight)
        {
            bool empty = range.Start == range.End;
            if (!empty)
            {
                List<Word.Range> ranges = findRegex(ltrRegex, range);
                foreach (Word.Range item in ranges)
                {
                    //do not change the following two functions order, exclude space first then fullstop(sentence period)
                    //excludeSpace(item);
                    //excludeFullstop(item);

                    empty = item.Start == item.End;
                    if (!empty)
                    {
                        item.HighlightColorIndex = highlightColor;
                        item.Select();
                        app.Selection.LtrRun();
                    }
                }
            }
        }

        /// <summary>
        /// Find and Replace all text on specified region property
        /// </summary>
        /// <param name="find">Text to find</param>
        /// <param name="replace">Text to replace with</param>
        /// <param name="range">Document Range</param>
        public void findReplaceText(string find, string replace, Word.Range range)
        {
            //Initialization
            Word.Range item = range;
            object replaceType = Word.WdReplace.wdReplaceAll;
            object missing = Type.Missing;
            item.Find.ClearFormatting();
            item.Find.ClearHitHighlight();
            item.Find.Replacement.ClearFormatting();

            item.Find.Text = find;
            item.Find.Replacement.Text = replace;

            item.Find.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref replaceType, ref missing, ref missing, ref missing, ref missing);

        }

        /// <summary>
        /// Find text using Wildcards(Regex) on specified region and Replace all with text or pattern
        /// </summary>
        /// <param name="regex">Regular Expression to find</param>
        /// <param name="replace">Text or pattern to replace with</param>
        /// <param name="range">Document Range</param>
        public void findReplaceRegex(string regex, string replace, Word.Range range)
        {
            //Initialization
            Word.Range item = range;
            object replaceType = Word.WdReplace.wdReplaceAll;
            object missing = Type.Missing;
            item.Find.ClearFormatting();
            item.Find.ClearHitHighlight();
            item.Find.Replacement.ClearFormatting();

            item.Find.MatchWildcards = true;
            item.Find.Text = regex;
            item.Find.Replacement.Text = replace;

            item.Find.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref replaceType, ref missing, ref missing, ref missing, ref missing);
        }

        public void findReplaceLanguage(string regex, Word.WdLanguageID replace, Word.Range range)
        {
            //Initialization
            Word.Range item = range;
            object replaceType = Word.WdReplace.wdReplaceAll;
            object missing = Type.Missing;
            item.Find.ClearFormatting();
            item.Find.ClearHitHighlight();
            item.Find.Replacement.ClearFormatting();

            item.Find.MatchWildcards = true;
            item.Find.Text = regex;
            item.Find.Replacement.LanguageID = replace;

            item.Find.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref replaceType, ref missing, ref missing, ref missing, ref missing);
        }

        public void clearHighlights()
        {
            foreach (Word.Range range in getRanges())
            {
                range.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;
            }
        }
        /// <summary>
        /// Fix Issues that rise between RTL and LTR languages in the same document
        /// </summary>
        public bool fixBidiIssues()
        {
            //[testing] undo the whol tool steps in one step or "Ctrl+z"
            Word.UndoRecord undo = app.UndoRecord;
            undo.StartCustomRecord("Fix Bidi Issues");
            
            clearHighlights();
            //(new)[Step 1]:
            //doc.Paragraphs.ReadingOrder = Word.WdReadingOrder.wdReadingOrderRtl;

            foreach (Word.Range range in getRanges())
            {
                //[Step 1]: Appy "Para Rtl" command on the specified range
                paraRtl(range);

                //[Step 2]: Apply "Run Rtl" command on the specified range
                runRtl(range);

                //[Step 3]: Find all weird spaces and replace it another normal space
                //findReplaceText(" ", " ", range);

                //[Step 4] (removed), directional chars will not be replaced, will be matched in the regex

                //[Step 5]: Apply "Run Ltr" command on all english characters
                //runLtrRegex(range);
                findReplaceLanguage(ltrRegex, Word.WdLanguageID.wdEnglishUS, range);

                //[Step 6(Repeating Step 3)]: Replace Space(normal and non-break) along with(\/:) with Space character to change the language of the space based on context
                findReplaceRegex("([  \\\\/:]{1,})", "\\1", range);

                //Below functions is in Testing phase
                unifyParagraphRTLFont(range);
                fixSpaceIssues(range);

                //[testing] closing undo record
                undo.EndCustomRecord();
            }

            app.ActiveWindow.View.Type = Word.WdViewType.wdPrintView;
            //doc.Range().SetRange(0, 0);
            doc.Range(0, 0).Select();
            

            //////////////////New Approach to fix issues/////////////////////////
            //runRtl(doc.Content);
            //findReplaceRegex("[‎‏]", "", doc.Content);
            //findReplaceText(" ", " ", doc.Content);
            //runLtrRegex_new(doc.Content);
            //findReplaceRegex("([a-zA-Z.][ ]{1,}[a-zA-Z])", "\\1", doc.Content);
            //findReplaceText(" ", " ", doc.Content);

            return true;
        }

        /// <summary>
        /// Find Regular Expression in a given range
        /// </summary>
        /// <param name="regex"></param>
        /// <param name="range"></param>
        /// <returns></returns>
        private List<Word.Range> findRegex(string regex, Word.Range range, Word.WdColorIndex highlightColor = Word.WdColorIndex.wdNoHighlight)
        {

            Word.Range rngCopy = range.Duplicate;
            rngCopy.TextRetrievalMode.IncludeHiddenText = false;
            rngCopy.TextRetrievalMode.IncludeFieldCodes = false;

            //original range indeces
            int rngStart = rngCopy.Start;
            int rngEnd = rngCopy.End;
            int curStart = 0, curEnd = 0;
            rngCopy.SetRange(rngStart, rngStart);

            List<Word.Range> ranges = new List<Word.Range>();

            bool empty = (rngStart == rngEnd);
            if (!empty)
            {

                //Initialization
                rngCopy.Find.ClearFormatting();
                rngCopy.Find.ClearHitHighlight();
                rngCopy.Find.MatchWildcards = true;
                rngCopy.Find.Text = regex;
                rngCopy.Find.Forward = true;
                //rngCopy.Find.Wrap = Word.WdFindWrap.wdFindStop;

                rngCopy.Find.Execute();

                //while there is found text, add its range to the found list
                while (rngCopy.Find.Found)
                {
                    //If the new range is equal to the current range or reached end of find criteria             
                    if (curStart == rngCopy.Start && curEnd == rngCopy.End)
                        break;
                    //check last range if it exceeds rngEnd, make extra find and replace to see if there is extra range found or not
                    if (rngCopy.End > rngEnd)
                    {
                        if (rngCopy.Start < rngEnd)
                        {
                            rngCopy.SetRange(rngCopy.Start, rngEnd);
                            rngCopy.HighlightColorIndex = highlightColor;
                            ranges.Add(rngCopy.Duplicate);
                        }
                        break;
                    }

                    //highlight found range
                    rngCopy.HighlightColorIndex = highlightColor;
                    //add the range to the ranges list
                    ranges.Add(rngCopy.Duplicate);

                    //added new
                    curStart = rngCopy.Start;
                    curEnd = rngCopy.End;
                    rngCopy.SetRange(curEnd, curEnd);

                    rngCopy.Find.Execute();
                }
            }
            return ranges;
        }

        /// <summary>
        /// Removes Extra spaces(normal and non-break) including:\n
        /// - Spaces before period or any dot
        /// - Spaces around backslashes (\/)
        /// Separate between sentences (add missing space after period)
        /// </summary>
        /// <param name="range">specified range</param>
        /// <exception cref="[Not Implemented] Add Space after each Period(Fullstop only)"></exception>
        public void fixSpaceIssues(Word.Range range)
        {
            //remove extra normal spaces
            findReplaceRegex("[ ]{1,}", " ", range);
            //remove extra Non Break spaces
            findReplaceRegex("[ ]{1,}", " ", range);
            //find any spaces(normal or non-break) before a dot or period and remove it
            findReplaceRegex("[  ]{1,}.", ".", range);
            //find any backslash with spaces(on both ends, on right only, on left only) and replace it with backslash only
            findReplaceRegex("[  ]{1,}([\\\\/])[  ]{1,}", "\\1", range);
            findReplaceRegex("([\\\\/])[  ]{1,}", "\\1", range);
            findReplaceRegex("[  ]{1,}([\\\\/])", "\\1", range);
            findReplaceRegex("(.)" + "([" + rtlChars + "])", "\\1 \\2", range);
        }

        /// <summary>
        /// Apply unified font name and size on all paragraph RTL text based on first RTL text word
        /// </summary>
        /// <param name="range">specified range</param>
        /// <remarks>Used first word font of each paragraph to determine the paragraph unified font(assuming that first word in each paragraph will be always RTL language)</remarks>
        /// 
        public void unifyParagraphRTLFont(Word.Range range)
        {
            string unifiedFontNameBi = "";
            double unifiedFontSizeBi = 0.0;
            foreach (Word.Paragraph item in range.Paragraphs)
            {
                //get first word font of RTL Language and apply it to the whole paragraph RTL Language Font
                unifiedFontNameBi = item.Range.Words[1].Font.NameBi;
                unifiedFontSizeBi = item.Range.Words[1].Font.SizeBi;

                item.Range.Font.NameBi = unifiedFontNameBi;
                item.Range.Font.SizeBi = (float)unifiedFontSizeBi;
            }
        }

        /// <summary>
        /// Apply unified font size on all paragraph text(RTL and LTR) based on RTL text font size
        /// </summary>
        /// <param name="range"></param>
        public void unifyParagraphFontSizeAsRTL(Word.Range range)
        {
            double unifiedFontSizeBi = 0.0;
            foreach (Word.Paragraph item in range.Paragraphs)
            {
                //get first word font of RTL Language and apply it to the whole paragraph RTL Language Font
                unifiedFontSizeBi = item.Range.Words[1].Font.SizeBi;

                //Change RTL & LTR font size to be the same as RTL font size
                item.Range.Font.SizeBi = (float)unifiedFontSizeBi;
                item.Range.Font.Size = (float)unifiedFontSizeBi;
            }

        }

        /// <summary>
        /// Apply unified font size on all paragraph text(RTL and LTR) based on LTR text font size
        /// </summary>
        /// <param name="range"></param>
        public void unifyParagraphFontSizeAsLTR(Word.Range range)
        {
            double unifiedFontSize = 0.0;
            foreach (Word.Paragraph item in range.Paragraphs)
            {
                //get first word font of RTL Language and apply it to the whole paragraph RTL Language Font
                unifiedFontSize = item.Range.Words[1].Font.Size;

                //Change RTL & LTR font size to be the same as RTL font size
                item.Range.Font.SizeBi = (float)unifiedFontSize;
                item.Range.Font.Size = (float)unifiedFontSize;
            }

        }

        /// <summary>
        /// Apply Font on a specified text language in a doument
        /// </summary>
        /// <param name="fontName"></param>
        /// <param name="lang"></param>
        /// <param name="ApplyOnStyles">Apply font changes on Styles when possible if "all regions" is specified</param>
        /// <returns></returns>
        public void applyFont(string fontName, Language lang, bool applyOnStyles)
        {

            if (lang == Language.Arabic)
            {
                //If "All regions", Change all the document's styles font name of complex(Bidirection(arabic)) text
                if (region == Region.All && applyOnStyles)
                {
                    foreach (Word.Style s in doc.Styles)
                    {
                        if (s.InUse)
                        {
                            s.Font.NameBi = fontName;
                            //s.Font.Size = s.Font.SizeBi;
                        }
                    }
                }

                //get ranges of selected document region
                foreach (Word.Range range in getRanges())
                {
                    ////find arabic text ranges in each region range(document region)
                    //foreach (Word.Range found in findRegex(rtlRegex, range))
                    //{
                    //    found.Font.NameBi = fontName;
                    //}

                    //added new
                    range.Font.NameBi = fontName;
                    //range.Font.Size = range.Font.SizeBi;
                }
            }
            else if (lang == Language.English)
            {
                //If "All regions", Change all the document's styles font name of english text
                if (region == Region.All && applyOnStyles)
                {
                    foreach (Word.Style s in doc.Styles)
                    {
                        if (s.InUse)
                        {
                            s.Font.NameAscii = fontName;
                            //s.Font.Size = s.Font.SizeBi;
                        }
                    }
                }

                //get ranges of selected document region
                foreach (Word.Range range in getRanges())
                {
                    ////find english text ranges in each region range(document region)
                    //foreach (Word.Range found in findRegex(ltrRegex, range))
                    //{
                    //    found.Font.NameAscii = fontName;
                    //}

                    //added new
                    range.Font.NameAscii = fontName;
                    //range.Font.Size = range.Font.SizeBi;
                }
            }
            else if (lang == Language.All)
            {
                //If "All regions", Change all the document's styles font name of english text
                if (region == Region.All && applyOnStyles)
                {
                    foreach (Word.Style s in doc.Styles)
                    {
                        if (s.InUse)
                        {
                            s.Font.Name = fontName;
                            s.Font.NameAscii = fontName;
                            s.Font.NameBi = fontName;
                            s.Font.NameFarEast = fontName;
                            s.Font.NameOther = fontName;
                            //s.Font.Size = s.Font.SizeBi;
                        }
                    }
                }
                //get ranges of selected document region
                foreach (Word.Range range in getRanges())
                {
                    //change font of each region at once
                    range.Font.Name = fontName;
                    range.Font.NameAscii = fontName;
                    range.Font.NameBi = fontName;
                    range.Font.NameFarEast = fontName;
                    range.Font.NameOther = fontName;
                    //range.Font.Size = range.Font.SizeBi;
                }
            }
        }
    }
}

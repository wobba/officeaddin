using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;
using Shapes = Microsoft.Office.Interop.PowerPoint.Shapes;

namespace PowerPointAddIn1
{
    public partial class LanguageSelector
    {
        private void LanguageSelector_Load(object sender, RibbonUIEventArgs e)
        {
            Globals.ThisAddIn.Application.AfterNewPresentation += InitRibbon;
            Globals.ThisAddIn.Application.AfterPresentationOpen += InitRibbon;
            try
            {
                InitRibbon(null);
            }
            catch (Exception)
            {
            }
        }

        private void InitRibbon(Presentation dummy)
        {
            HashSet<MsoLanguageID> usedLanguages = FindUsedLanguages();
            List<RibbonDropDownItem> languages = CreateLanguageItems(usedLanguages);
            InsertUsedLanguagesAtTop(usedLanguages, languages);

            dropDownLanguage.Items.Clear();
            foreach (RibbonDropDownItem ribbonDropDownItem in languages)
            {
                dropDownLanguage.Items.Add(ribbonDropDownItem);
            }
        }
        
        private void InsertUsedLanguagesAtTop(IEnumerable<MsoLanguageID> usedLanguages, IList<RibbonDropDownItem> languages)
        {
            foreach (MsoLanguageID language in usedLanguages)
            {
                var item = CreateItem(language);
                languages.Insert(0, item);
            }
        }
        
        private string FormatLanguageString(MsoLanguageID language)
        {
            string formattedLanguage = language.ToString().Replace("msoLanguageID", "");
            StringBuilder sb = new StringBuilder(formattedLanguage.Length);
            sb.Append(formattedLanguage[0]);
            bool previousWasUpper = true;
            bool moreUppers = false;
            for (int i = 1; i < formattedLanguage.Length; i++)
            {
                if (Char.IsUpper(formattedLanguage[i]))
                {
                    if (!previousWasUpper) sb.Append(' ');
                    if (!moreUppers)
                    {
                        sb.Append('(');
                        moreUppers = true;
                    }
                    previousWasUpper = true;
                }
                else
                {
                    previousWasUpper = false;
                }
                sb.Append(formattedLanguage[i]);
            }
            if (moreUppers) sb.Append(')');
            return sb.ToString();
        }

        
        private List<RibbonDropDownItem> CreateLanguageItems(ICollection<MsoLanguageID> usedLanguages)
        {
            var languages = new List<RibbonDropDownItem>(100);
            languages.AddRange(from MsoLanguageID language in Enum.GetValues(typeof (MsoLanguageID))
                               where !usedLanguages.Contains(language)
                               let ddi = Factory.CreateRibbonDropDownItem()
                               select CreateItem(language)
                               );
            languages.Sort((l, r) => l.Label.CompareTo(r.Label));
            return languages;
        }

        RibbonDropDownItem CreateItem( MsoLanguageID languageId )
        {
            var ddi = Factory.CreateRibbonDropDownItem();
            ddi.Label = FormatLanguageString(languageId);
            ddi.Tag = languageId;
            return ddi;
        }

        private HashSet<MsoLanguageID> FindUsedLanguages()
        {
            Slides slides = Globals.ThisAddIn.Application.ActivePresentation.Slides;
            int slideCount = slides.Count;
            HashSet<MsoLanguageID> usedLanguages = new HashSet<MsoLanguageID>();
            for (int i = 1; i <= slideCount; i++)
            {
                GetUsedShapeLanguage(slides[i].Shapes, usedLanguages);
                GetUsedShapeLanguage(slides[i].NotesPage.Shapes, usedLanguages);
            }
            return usedLanguages;
        }

        private void ButtonSetLanguageClick(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.ActivePresentation.DefaultLanguageID = (MsoLanguageID)dropDownLanguage.SelectedItem.Tag;
            Slides slides = Globals.ThisAddIn.Application.ActivePresentation.Slides;
            int slideCount = slides.Count;
            for (int i = 1; i <= slideCount; i++)
            {
                SetShapeLanguage(slides[i].Shapes);
                SetShapeLanguage(slides[i].NotesPage.Shapes);
            }
        }

        private void SetShapeLanguage(Shapes slideShapes)
        {
            for (int j = 1; j <= slideShapes.Count; j++)
            {
                if (slideShapes[j].HasTextFrame == MsoTriState.msoTrue)
                {
                    slideShapes[j].TextFrame.TextRange.LanguageID = (MsoLanguageID)dropDownLanguage.SelectedItem.Tag;
                }

                if (slideShapes[j].HasTable == MsoTriState.msoTrue)
                {
                    for (int row = 1; row <= slideShapes[j].Table.Rows.Count; row++)
                    {
                        CellRange cells = slideShapes[j].Table.Rows[row].Cells;
                        for (int cell = 1; cell <= slideShapes[j].Table.Rows[row].Cells.Count; cell++)
                        {
                            var thisCell = cells[cell];

                            if (thisCell.Shape.HasTextFrame == MsoTriState.msoTrue)
                            {
                                thisCell.Shape.TextFrame.TextRange.LanguageID = (MsoLanguageID)dropDownLanguage.SelectedItem.Tag;
                            }
                        }

                    }
                }

            }
        }

        private void GetUsedShapeLanguage(Shapes slideShapes, HashSet<MsoLanguageID> usedLanguages)
        {
            for (int j = 1; j <= slideShapes.Count; j++)
            {
                if (slideShapes[j].HasTextFrame == MsoTriState.msoTrue)
                {
                    usedLanguages.Add(slideShapes[j].TextFrame.TextRange.LanguageID);
                }
            }
        }

        private void buttonSetLanguage_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.ActivePresentation.DefaultLanguageID = (MsoLanguageID)dropDownLanguage.SelectedItem.Tag;
            Slides slides = Globals.ThisAddIn.Application.ActivePresentation.Slides;
            int slideCount = slides.Count;
            for (int i = 1; i <= slideCount; i++)
            {
                SetShapeLanguage(slides[i].Shapes);
                SetShapeLanguage(slides[i].NotesPage.Shapes);
            }
        }
    }
}

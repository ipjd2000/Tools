using LLDOCX.CustomUserControl;
using LLDOCX.Options;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace LLDOCX.Helpers
{
    class WordDocumentHelper
    {

        public static List<DocumentHeadingData> ValidateDocument(SettingOption option)
        {
            var missingHeadings = new List<DocumentHeadingData>();
            Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;
            if (option.HeadingValidation)
            {
                missingHeadings = validateHeading(document);
            }
            if (option.CheckPageSize)
            {
                ValidatePageSize(document);
            }
            if (option.CheckPageOrientation)
            {
                ValidatePageOrientation(document);
            }

            if (option.CheckMarkup)
            {
                ValidateMarkup(document, option.KeepTrackingOn);
            }
            if (option.CheckFont)
            {
                ValidateFont(document, option);
            }
            if (option.CheckHyperlink)
            {
                ValidateHyperLink(document);
            }
            if (option.CheckComment)
            {
                ValidateComment(document);
            }
            if (option.CheckBookmarker)
            {
                ValidateBookMarks(document);
            }
            if (option.RemoveDocumentProperty)
            {
                ValidateDocProperties(document);
            }
            return missingHeadings;
        }
        private static List<DocumentHeadingData> BuildDocumentMetaData()
        {
            var documentheadings = new List<DocumentHeadingData>();
            var documentMetaData = new DocumentHeadingData()
            {
                Heading = "Related and Priority",
                SubHeadings = new List<string>()
                {
                    "REFERENCE TO RELATED APPLICATIONS", "RELATED APPLICATIONS", "CROSS-REFERENCES", "CROSS REFERENCES", "REFERENCE TO RELATED PATENTS", "CROSS REFERENCE TO RELATED APPLICATIONS", "CROSS-REFERENCES TO RELATED APPLICATIONS"
                }
                ,
                Index = 1
            };
            documentheadings.Add(documentMetaData);

            documentMetaData = new DocumentHeadingData()
            {
                Heading = "FIELD",
                IsRequired = true,
                Index = 2,
                SubHeadings = new List<string>()
                {
                     "FIELD AND BACKGROUND OF THE INVENTION", "FIELD OF THE INVENTION", "TECHNICAL FIELD", "TECHNICAL FIELD OF THE INVENTION"
                }
            };
            documentheadings.Add(documentMetaData);

            documentMetaData = new DocumentHeadingData()
            {
                Heading = "Background/Prior Art Headings",
                IsRequired = true,
                Index = 3,
                SubHeadings = new List<string>()
                {
                     "BACKGROUND OF THE INVENTION", "FIELD AND BACKGROUND OF THE INVENTION", "BRIEF STATEMENT OF THE INVENTION", "DESCRIPTION OF RELATED ART", "DESCRIPTION OF THE RELATED ART", "PRIOR ART", "DESCRIPTION OF THE PRIOR ART", "BRIEF STATEMENT OF THE PRIOR ART", "BACKGROUND ART"
                }
            };
            documentheadings.Add(documentMetaData);

            documentMetaData = new DocumentHeadingData()
            {
                Heading = "Summary Headings",
                IsRequired = true,
                Index = 4,
                SubHeadings = new List<string>()
                {
                    "SUMMARY OF THE INVENTION"
                }
            };
            documentheadings.Add(documentMetaData);

            documentMetaData = new DocumentHeadingData()
            {
                Heading = "Figure Headings",
                IsRequired = true,
                Index = 5,
                SubHeadings = new List<string>()
                {
                    "BRIEF DESCRIPTION OF DRAWINGS", "BRIEF DESCRIPTION OF THE DRAWINGS", "BRIEF DESCRIPTION OF FIGURES", "BRIEF DESCRIPTION OF THE FIGURES", "DETAILED DESCRIPTION OF THE DRAWINGS", "DRAWING", "DRAWINGS", "FIGURE", "FIGURES"
                }
            };
            documentheadings.Add(documentMetaData);

            documentMetaData = new DocumentHeadingData()
            {
                Heading = "Specification Headings",
                IsRequired = true,
                Index = 6,
                SubHeadings = new List<string>()
                {
                    "DESCRIPTION", "DETAILED DESCRIPTION", "DESCRIPTION OF THE INVENTION", "DETAILED DESCRIPTION OF THE INVENTION"
                }
            };
            documentheadings.Add(documentMetaData);

            documentMetaData = new DocumentHeadingData()
            {
                Heading = "Abstract Section",
                IsRequired = true,
                Index = 7,
                SubHeadings = new List<string>()
                {
                   "ABSTRACT", "ABSTRACT OF THE INVENTION", "INVENTION ABSTRACT", "DISCLOSURE ABSTRACT", "ABSTRACT OF THE DISCLOSURE", "AMENDMENT TO THE ABSTRACT", "SUBSTITUTE ABSTRACT"
                }
            };
            documentheadings.Add(documentMetaData);

            return documentheadings;
        }
        private static List<DocumentHeadingData> validateHeading(Word.Document document)
        {
            var listofMissingHeader = new List<DocumentHeadingData>();
            var listofHeading = BuildDocumentMetaData();
            foreach (var heading in listofHeading)
            {
                bool matchHeading = false;
                foreach (var subHeading in heading.SubHeadings)
                {
                    matchHeading = document.Content.Find.Execute(subHeading, true, true);
                    if (matchHeading)
                        break;
                }
                if (!matchHeading)
                {
                    listofMissingHeader.Add(heading);
                }
            }
            return listofMissingHeader;
        }

        private static void ValidatePageSize(Word.Document document)
        {
            var pageSetup = document.PageSetup;
            if (pageSetup.PaperSize != Word.WdPaperSize.wdPaperA4)
            {
                var dialog = MessageBox.Show("The page size of this document exceeds Letter or A4. Please correct the page size before proceeding. Do you want to set Page Size to A4?", LLDocxConstant.AppName, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dialog == DialogResult.Yes)
                {
                    pageSetup.PaperSize = Word.WdPaperSize.wdPaperA4;
                }
            }
        }
        private static void ValidatePageOrientation(Word.Document document)
        {
            var pageSetup = document.PageSetup;
            if (pageSetup.Orientation != Word.WdOrientation.wdOrientPortrait)
            {
                var dialog = MessageBox.Show("The orientation of the document is not portrait which may not be rendered correctly by the USPTO. Please correct by changing the orientation to portrait.. Do you want to set Page Orientation to Portrait?", LLDocxConstant.AppName, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dialog == DialogResult.Yes)
                {
                    pageSetup.Orientation = Word.WdOrientation.wdOrientPortrait;
                }
            }
        }
        private static void ValidateMarkup(Word.Document document, bool keepTracking)
        {
            var revisions = document.Revisions;
            if (document.TrackRevisions)
            {
                var option = ProcessData.GetData();
                option.KeepTrackingOn = false;
                ProcessData.SaveData(option);
            }
            if (revisions.Count > 0)
            {
                MessageBoxManager.Yes = "Accept All";
                MessageBoxManager.No = "Reject All";
                MessageBoxManager.Register();
                var dialog = MessageBox.Show("Do you want to  Accept All / Reject All or CANCEL revison", LLDocxConstant.AppName, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (dialog == DialogResult.Yes)
                {
                    document.AcceptAllRevisions();
                }
                else if (dialog == DialogResult.No)
                {
                    document.RejectAllRevisions();
                }
                MessageBoxManager.Unregister();
            }
            document.ShowRevisions = keepTracking;
        }
        private static void ValidateFont(Word.Document document, SettingOption option)
        {
            Word.Font font = document.Content.Font;
            font.Size = 12;
            font.Color = Word.WdColor.wdColorBlack;
            font.Name = option.DefaultFontName;
            font.Color = Word.WdColor.wdColorBlack;
            document.Content.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;
            ClearBackgroundColor(document);
        }
        private static void ClearBackgroundColor(Word.Document document)
        {
            Color color = Color.FromArgb(255, 255, 255);
            document.Background.Fill.ForeColor.RGB = ColorTranslator.ToOle(color);
            document.Background.Fill.Visible = Office.MsoTriState.msoTrue;
            document.Background.Fill.Solid();
        }
        private static void ValidateHyperLink(Word.Document document)
        {
            var hyperlinks = document.Hyperlinks;
            foreach (Word.Hyperlink hyperlink in hyperlinks)
            {
                hyperlink.Delete();
            }
            var fields = document.Fields;
            foreach (Word.Field field in fields)
            {
                field.Unlink();
            }

            object numberType = Word.WdNumberType.wdNumberAllNumbers;
            document.ConvertNumbersToText(numberType);

        }
        private static void ValidateComment(Word.Document document)
        {
            var comments = document.Comments;
            foreach (Word.Comment comment in comments)
            {
                comment.DeleteRecursively();
            }

        }
        private static void ValidateBookMarks(Word.Document document)
        {
            var bookmarks = document.Bookmarks;
            if (bookmarks.Count > 0)
            {
                foreach (Word.Bookmark bookmark in bookmarks)
                {
                    bookmark.Delete();
                }
            }
        }
        private static void ValidateDocProperties(Word.Document document)
        {
            document.BuiltInDocumentProperties[Word.WdBuiltInProperty.wdPropertyAuthor].Value = string.Empty;
            document.BuiltInDocumentProperties[Word.WdBuiltInProperty.wdPropertyKeywords].Value = string.Empty;
            document.BuiltInDocumentProperties[Word.WdBuiltInProperty.wdPropertyComments].Value = string.Empty;
            document.BuiltInDocumentProperties[Word.WdBuiltInProperty.wdPropertyTitle].Value = string.Empty;
            document.RemovePersonalInformation = true;
        }
    }
}

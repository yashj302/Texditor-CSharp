using System;
using System.Collections.Generic;
using System.Text;
using ESCommon.Rtf;
using ESCommon;
using System.Drawing;

namespace RtfConstructor
{
    public class Report
    {
        private RtfDocument rtf;
        
        public Report()
        {
            rtf = new RtfDocument();
            rtf.FontTable.Add(new RtfFont("Calibri"));
            rtf.FontTable.Add(new RtfFont("Constantia"));
            rtf.ColorTable.AddRange(new RtfColor[] {
                new RtfColor(Color.Red),
                new RtfColor(0, 0, 255)
            });

            RtfParagraphFormatting LeftAligned12 = new RtfParagraphFormatting(12, RtfTextAlign.Left);
            RtfParagraphFormatting Centered10 = new RtfParagraphFormatting(10, RtfTextAlign.Center);

            RtfFormattedParagraph header = new RtfFormattedParagraph(new RtfParagraphFormatting(16, RtfTextAlign.Center));
            RtfFormattedParagraph p1 = new RtfFormattedParagraph(new RtfParagraphFormatting(12, RtfTextAlign.Left));

            RtfTable t = new RtfTable(RtfTableAlign.Center, 2, 3);

            header.Formatting.SpaceAfter = TwipConverter.ToTwip(12F, MetricUnit.Point);
            header.AppendText("Calibri ");
            header.AppendText(new RtfFormattedText("Bold", RtfCharacterFormatting.Bold));

            t.Width = TwipConverter.ToTwip(5, MetricUnit.Centimeter);
            t.Columns[1].Width = TwipConverter.ToTwip(2, MetricUnit.Centimeter);

            foreach (RtfTableRow row in t.Rows)
            {
                row.Height = TwipConverter.ToTwip(2, MetricUnit.Centimeter);
            }

            t.MergeCellsVertically(1, 0, 2);

            t.DefaultCellStyle = new RtfTableCellStyle(RtfBorderSetting.None, Centered10);

            t[0, 0].Definition.Style = new RtfTableCellStyle(RtfBorderSetting.None, LeftAligned12, RtfTableCellVerticalAlign.Bottom);
            t[0, 0].AppendText("Bottom");

            t[1, 0].Definition.Style = new RtfTableCellStyle(RtfBorderSetting.Left, Centered10, RtfTableCellVerticalAlign.Center, RtfTableCellTextFlow.BottomToTopLeftToRight);
            t[1, 1].Definition.Style = t[1, 0].Definition.Style;
            t[1, 0].AppendText("Vertical");

            t[0, 1].Formatting = new RtfParagraphFormatting(10, RtfTextAlign.Center);
            t[0, 1].Formatting.TextColorIndex = 1;
            t[0, 1].AppendText(new RtfFormattedText("Black ", 0));
            t[0, 1].AppendText("Red ");
            t[0, 1].AppendText(new RtfFormattedText("Blue", 2));

            t[0, 2].AppendText("Normal");
            t[1, 2].AppendText(new RtfFormattedText("Italic", RtfCharacterFormatting.Caps | RtfCharacterFormatting.Italic));
            t[1, 2].AppendParagraph("+");
            t[1, 2].AppendParagraph(new RtfFormattedText("Caps", RtfCharacterFormatting.Caps | RtfCharacterFormatting.Italic));

            p1.Formatting.FontIndex = 1;
            p1.Formatting.IndentLeft = TwipConverter.ToTwip(6.05F, MetricUnit.Centimeter);
            p1.Formatting.SpaceBefore = TwipConverter.ToTwip(6F, MetricUnit.Point);
            p1.AppendText("Constantia ");
            p1.AppendText(new RtfFormattedText("Superscript", RtfCharacterFormatting.Superscript));
            p1.AppendParagraph(new RtfFormattedText("Inline", -1, 8));
            p1.AppendText(new RtfFormattedText(" font size ", -1, 14));
            p1.AppendText(new RtfFormattedText("change", -1, 8));
            
            RtfImage picture = new RtfImage(Properties.Resources.lemon, RtfImageFormat.Wmf);
            picture.ScaleX = 50;
            picture.ScaleY = 50;

            p1.AppendParagraph(picture);

            RtfFormattedText linkText = new RtfFormattedText("View article", RtfCharacterFormatting.Underline, 2);
            linkText.BackgroundColorIndex = 1;
            p1.AppendParagraph(new RtfHyperlink("http://www.codeproject.com/KB/cs/RtfConstructor.aspx", linkText));

            RtfFormattedParagraph p2 = new RtfFormattedParagraph();
            p2.Formatting = new RtfParagraphFormatting(10);
            
            p2.Tabs.Add(new RtfTab(TwipConverter.ToTwip(2.5F, MetricUnit.Centimeter), RtfTabKind.Decimal));
            p2.Tabs.Add(new RtfTab(TwipConverter.ToTwip(5F, MetricUnit.Centimeter), RtfTabKind.FlushRight, RtfTabLead.Dots));
            p2.Tabs.Add(new RtfTab(TwipConverter.ToTwip(7.5F, MetricUnit.Centimeter)));
            p2.Tabs.Add(new RtfTab(TwipConverter.ToTwip(10F, MetricUnit.Centimeter), RtfTabKind.Centered));
            p2.Tabs.Add(new RtfTab(TwipConverter.ToTwip(15F, MetricUnit.Centimeter), RtfTabLead.Hyphens));

            p2.Tabs[2].Bar = true;

            p2.AppendText("One");
            p2.AppendText(new RtfTabCharacter());
            p2.AppendText("Two");
            p2.AppendText(new RtfTabCharacter());
            p2.AppendText("Three");
            p2.AppendText(new RtfTabCharacter());
            p2.AppendText("Five");
            p2.AppendText(new RtfTabCharacter());
            p2.AppendText("Six");

            rtf.Contents.AddRange(new RtfDocumentContentBase[] {
                header,
                t,
                p1,
                p2,
            });
        }

        public RtfDocument GetRtf()
        {
            return rtf;
        }
    }
}
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace TaskManagementSystem.Helpers
{
    public class PdfFooterEventHelper : PdfPageEventHelper
    {
        private readonly string _printedBy;
        private PdfContentByte _cb;
        private PdfTemplate _template;
        private BaseFont _bf;

        public PdfFooterEventHelper(string printedBy)
        {
            _printedBy = printedBy;
        }

        public override void OnOpenDocument(PdfWriter writer, Document document)
        {
            _cb = writer.DirectContent;
            _template = _cb.CreateTemplate(50, 50);
            _bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        }

        public override void OnEndPage(PdfWriter writer, Document document)
        {
            int pageN = writer.PageNumber;
            string textLeft = "Developed by Md Firoz Ali";
            string textCenter = $"Printed by: {_printedBy}";
            string textRight = "Page " + pageN + " of ";

            float lenRight = _bf.GetWidthPoint(textRight, 9);
            Rectangle page = document.PageSize;

            // Left
            _cb.BeginText();
            _cb.SetFontAndSize(_bf, 9);
            _cb.SetTextMatrix(document.LeftMargin, page.GetBottom(30));
            _cb.ShowText(textLeft);
            _cb.EndText();

            // Center
            float centerPos = (page.Left + page.Right) / 2 - (_bf.GetWidthPoint(textCenter, 9) / 2);
            _cb.BeginText();
            _cb.SetFontAndSize(_bf, 9);
            _cb.SetTextMatrix(centerPos, page.GetBottom(30));
            _cb.ShowText(textCenter);
            _cb.EndText();

            // Right
            float rightPos = page.Right - document.RightMargin - lenRight;
            _cb.BeginText();
            _cb.SetFontAndSize(_bf, 9);
            _cb.SetTextMatrix(rightPos, page.GetBottom(30));
            _cb.ShowText(textRight);
            _cb.EndText();

            // Page total template
            _cb.AddTemplate(_template, rightPos + lenRight, page.GetBottom(30));
        }

        public override void OnCloseDocument(PdfWriter writer, Document document)
        {
            _template.BeginText();
            _template.SetFontAndSize(_bf, 9);
            _template.SetTextMatrix(0, 0);
            _template.ShowText("" + (writer.PageNumber - 1));
            _template.EndText();
        }
    }
}

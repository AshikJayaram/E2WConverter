using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentProcessor.Model;
using Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Runtime.InteropServices;
namespace DocumentProcessor.WordProcessor
{
    public class DocumentWriter : IDocumentWriter
    {
        #region private variables

        object oMissing = Missing.Value;

        object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

        #endregion

        #region Public methods
        public void WriteToDocument(List<MemberDetails> memberList)
        {
            var app = this.GetApplication();

            app.Visible = false;

            Document doc = app.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            this.WriteTableWithContent(app, doc, memberList, oMissing, oEndOfDoc);

            doc.Save();

            app.Quit();

            Marshal.FinalReleaseComObject(app);
        }

        #endregion

        #region private methods

        private void WriteTableWithContent(Application app, Document doc,List<MemberDetails> memberList, object oMissing, object endOfDoc)
        {
            var rowCount = (memberList.Count + 1) / 2;

            Range wordRange = doc.Bookmarks.get_Item(ref endOfDoc).Range;

            Table table = doc.Tables.Add(wordRange, rowCount, 2, ref oMissing, ref oMissing);

            table = this.GetTableFormat(table, app);

            string strText;

            var count = 1;

            for (int r = 1; r <= rowCount; r++)
                for (int c = 1; c <= 2; c++)
                {
                    if (count < memberList.Count)
                    {
                        strText = String.Format("\r\a" +memberList[count].MemberNumber + "\r\a" + 
                                            memberList[count].Title + "." + 
                                            memberList[count].Name + "\r\a" +
                                            memberList[count].AddressOne + "\r\a" +
                                            memberList[count].AddressTwo + "\r\a" +
                                            memberList[count].Place + " - " + memberList[count].PinCode  + "\r\a" +
                                            memberList[count].MobileNumber + "\r\a"
                                            );
                        table.Cell(r, c).Range.Text = strText.ToUpper();
                    }
                    count++;
                }
        }

        //get all the formattings from app config :ToDo
        private Table GetTableFormat(Table table, Application app)
        {
            table.Range.ParagraphFormat.SpaceAfter = 1;
            table.Range.Font.Name = "Calibri";
            table.Range.Font.Size = 11;
            table.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            table.Rows.Height = app.InchesToPoints(2);
            table.Borders.Enable = 1;
            table.Range.Rows.DistanceTop = 1;
            table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);
            table.AllowAutoFit = true;

            return table;
        }
        private Application GetApplication()
        {
            return new Application();
        }

        #endregion
    }
}

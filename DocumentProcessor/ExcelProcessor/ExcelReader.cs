using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using DocumentProcessor.Model;
using System.IO;
using System.Runtime.InteropServices;

namespace DocumentProcessor.ExcelProcessor
{
    public class ExcelReader : IExcelReader
    {
        #region public methods
        public List<MemberDetails> GetMembersList(string fileName)
        {
            var app = this.GetApplication();

            app.Visible = false;

            var workSheets = this.GetWorkSheets(app, fileName);

            var valueArray = this.GetObjectListFromSheet(workSheets[0]);

            List<MemberDetails> memberList =  this.GetMemberList(valueArray);

            app.Quit();

            Marshal.FinalReleaseComObject(app);

            return memberList;
        }

        public List<string> GetWorkSheetNames(string fileName)
        {
            var app = this.GetApplication();

            app.Visible = false;

            var workSheets = this.GetWorkSheets(app, fileName);

            return workSheets.Select(i => i.Name).ToList();
        }

        #endregion


        #region private methods

        private Application GetApplication()
        {
            return new Application();
        }

        private List<Worksheet> GetWorkSheets(Application app, string fileName)
        {
            var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ("Uploads\\"+fileName));

            var workBook = app.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            Worksheet workSheet = null;

            var sheets = new List<Worksheet>();

            foreach (var sheet in workBook.Sheets)
            {
                workSheet = (Worksheet)sheet;

                sheets.Add(workSheet);
            }

            return sheets;
        }

        private List<MemberDetails> GetMemberList(Object[,] valueArray)
        {
            var boundUpper = valueArray.GetUpperBound(0);
            var boundLower = valueArray.GetUpperBound(1);

            var memberList = new List<MemberDetails>();

            for (long i = 1; i <= boundUpper; i++)
            {
                var member = new MemberDetails
                {
                    MemberNumber = Enumerable.Range(0, boundUpper).Select(j => valueArray[i, 1]).FirstOrDefault() != null ? Enumerable.Range(0, boundUpper).Select(j => valueArray[i, 1]).FirstOrDefault().ToString() : "",
                    Title = Enumerable.Range(0, boundUpper).Select(j => valueArray[i, 2]).FirstOrDefault() != null ? Enumerable.Range(0, boundUpper).Select(j => valueArray[i, 2]).FirstOrDefault().ToString() : "",
                    Name = Enumerable.Range(0, boundUpper).Select(j => valueArray[i, 3]).FirstOrDefault() != null ? Enumerable.Range(0, boundUpper).Select(j => valueArray[i, 3]).FirstOrDefault().ToString() : "",
                    AddressOne = Enumerable.Range(0, boundUpper).Select(j => valueArray[i, 4]).FirstOrDefault() != null ? Enumerable.Range(0, boundUpper).Select(j => valueArray[i, 4]).FirstOrDefault().ToString() : "",
                    AddressTwo = Enumerable.Range(0, boundUpper).Select(j => valueArray[i, 5]).FirstOrDefault() != null ? Enumerable.Range(0, boundUpper).Select(j => valueArray[i, 5]).FirstOrDefault().ToString() : "",
                    Place = Enumerable.Range(0, boundUpper).Select(j => valueArray[i, 6]).FirstOrDefault() != null ? Enumerable.Range(0, boundUpper).Select(j => valueArray[i, 6]).FirstOrDefault().ToString() : "",
                    PinCode = Enumerable.Range(0, boundUpper).Select(j => valueArray[i, 7]).FirstOrDefault() != null ? Enumerable.Range(0, boundUpper).Select(j => valueArray[i, 7]).FirstOrDefault().ToString() : "",
                    MobileNumber = Enumerable.Range(0, boundUpper).Select(j => valueArray[i, 8]).FirstOrDefault() != null ? Enumerable.Range(0, boundUpper).Select(j => valueArray[i, 8]).FirstOrDefault().ToString() : ""
                };
                memberList.Add(member);
            }
            return memberList.ToList(); 
        }

        private Object[,] GetObjectListFromSheet(Worksheet sheet)
        {
            Range range = sheet.UsedRange;

            Object[,] valueArray = (object[,])range.get_Value(XlRangeValueDataType.xlRangeValueDefault);

            return valueArray;
        }

        #endregion
    }
}
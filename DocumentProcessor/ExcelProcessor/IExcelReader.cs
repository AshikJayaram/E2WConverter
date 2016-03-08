using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentProcessor.Model;

namespace DocumentProcessor.ExcelProcessor
{
    public interface IExcelReader
    {
        List<MemberDetails> GetMembersList(string fileName);

        List<string> GetWorkSheetNames(string fileName);
    }
}

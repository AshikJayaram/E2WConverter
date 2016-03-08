using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentProcessor.Model;

namespace DocumentProcessor.WordProcessor
{
    public interface IDocumentWriter
    {
        void WriteToDocument(List<MemberDetails> memberList);
    }
}

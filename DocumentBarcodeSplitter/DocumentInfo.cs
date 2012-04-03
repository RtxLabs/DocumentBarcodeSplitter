using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DocumentBarcodeSplitter
{
    class DocumentInfo
    {
        public DocumentInfo(int startPage)
        {
            this.startPage = startPage;
            this.endPage = startPage;
        }

        public String barcode;
        public int startPage;
        public int endPage;
    }
}

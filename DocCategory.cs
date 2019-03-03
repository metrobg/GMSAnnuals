using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace fyiReporting
{
    class DocCategory
    {
        private decimal _docseq;
        private string _category;
         

        public DocCategory(decimal sequence,string category)
        {
            _docseq = sequence;
            _category = category;
        }

        public decimal getDocSeq()
        {
            return _docseq;
        }

        public string getDocCategory()
        {

            return _category;
        }
    
    }
}

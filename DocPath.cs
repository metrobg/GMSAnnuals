using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace fyiReporting
{
    class DocPath
    {
        private decimal _pathnum;
        private string _storagepath;
         

        public DocPath(decimal number,string path)
        {
            _pathnum = number;
            _storagepath = path;
        }

        public decimal getPathNum()
        {
            return _pathnum;
        }

        public string getStoragePath()
        {

            return _storagepath;
        }
    
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace fyiReporting.RdlCmd
{
    public class WardDocument
    {
        private decimal _basepath;
        private string _docpath;
        private decimal _docnum;
        private string _doctype;
        private decimal _wardnum;
        private string _storagepath;
        private string _wardname;

        public bool isInitialized;

        public WardDocument()
        {
            isInitialized = true;
        }


        public WardDocument(decimal basepath, string docpath, decimal docnum, string doctype, decimal ward, string storagepath, string wardname)
        {
            _basepath = basepath;
            _docpath = docpath;
            _docnum = docnum;
            _doctype = doctype;
            _wardnum = ward;
            _storagepath = storagepath;
            _wardname = wardname;
        }

        public decimal getDocnum()
        {
            return _docnum;
        }

        public string getDoctype()
        {
            return _doctype;
        }

        public decimal getWard()
        {
            return _wardnum;
        }
        public decimal getBasePath()
        {
            return _basepath;
        }

        public string getDocPath()
        {
            return _docpath;
        }

        public string getStoragePath()
        {
            return _storagepath;
        }

        public string getWardName()
        {
            return _wardname;
        }

    }
}

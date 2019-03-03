using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace fyiReporting
{
    class DocType
    {
        private decimal _id;
        private string _short_name;
        private string _description;

        public DocType(decimal id, string short_Name, string description)
        {
            _id = id;
            _short_name = short_Name;
            _description = description;
        }

        public decimal getID()
        {
            return _id;
        }
        public string getShortName()
        {
            return _short_name;
        }

        public string getDescription()
        {
            return _description;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace fyiReporting.RdlCmd
{
    public class Ward
    {
        private decimal _ward_number;
        private string _ward_name;
        private decimal _responsible_employee;
        private string _employee_name;
        private string _file_number;
        private string _status;



        public Ward(decimal ward_number, string ward_name, decimal responsible_employee, string employee_name,string file_number,string status)
        {


            _ward_number = ward_number;
            _ward_name = ward_name;
            _responsible_employee = responsible_employee;
            _employee_name = employee_name;
            _file_number = file_number;
            _status = status;
        }



        public decimal getWardNumber()
        {
            return _ward_number;
        }
        public decimal getResponsibleEmployee()
        {
            return _responsible_employee;
        }

        public string getEmployeeName()
        {
            return _employee_name;
        }
        public string getStatus()
        {
            return _status;
        }
        public string getWardName()
        {
           // return _ward_name;
            return _ward_name.Replace("/", "_").Replace("/","_");  // handle ward_name_A/K/A so multiple folders are not created
        }
        public string getFileNumber()
        {
            return _file_number;
        }

    }
}

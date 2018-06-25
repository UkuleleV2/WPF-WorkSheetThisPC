using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WorksheetThisPC
{
    class Computer
    {
        public string hostname { set; get; }
        public string username { set; get; }
        public string system { set; get; }
        public string system_key { set; get; }
        public string processor { set; get; }
        public string memory { set; get; }
        public string disk { set; get; }
        public string ethernet_mac { set; get; }
        public string wireless_mac { set; get; }
        public string office { set; get; }
        public string office_key { set; get; }

        public Computer()
        {
            hostname = "";
            username = "";
            system = "";
            system_key = "";
            processor = "";
            memory = "";
            disk = "";
            ethernet_mac = "";
            wireless_mac = "";
            office = "";
            office_key = "";
        }
    }
}

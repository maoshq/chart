using System;
using System.Collections.Generic;
using System.Text;

namespace UITest.Model
{
    public class Driver
    {
        public string DriverName{ get; set; }

        public bool IsSelected { get; set; }
        public Driver(string driverName)
        {
            DriverName = driverName ?? throw new ArgumentNullException(nameof(driverName));
            IsSelected = false;
        }

        public Driver()
        {
        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace UITest.Model
{
    public class Settings 
    {
        public bool isNet { get; set; }

        public Settings(bool isNet)
        {
            this.isNet = isNet;
        }

        public Settings()
        {
        }
    }
}

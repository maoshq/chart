using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace UITest.Model
{
    public class Settings 
    {
        public bool isNet { get; set; }

        public List<string> historyDriver { get; set; }

        public bool flag { get; set; }

        public string CurrentDriver { get; set; }

        public Settings(bool isNet, List<string> historyDriver, bool flag, string currentDriver)
        {
            this.isNet = isNet;
            this.historyDriver = historyDriver ?? throw new ArgumentNullException(nameof(historyDriver));
            this.flag = flag;
            CurrentDriver = currentDriver ?? throw new ArgumentNullException(nameof(currentDriver));
        }

        public Settings()
        {
        }
    }
}

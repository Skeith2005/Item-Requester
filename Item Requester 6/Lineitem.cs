using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RequesterMethod
{
    public class LineItem
    {
        public string Description { get; set; }
        public string MFG { get; set; }
        public string Department { get; set; }
        public string Class { get; set; }
        public string Fineline { get; set; }
        public string Vendor { get; set; }
        public string RPLCost { get; set; }
        public string Retail { get; set; }
        public string DesiredGP { get; set; }
        public string List { get; set; }
        public string Purch { get; set; }
        public string Stock { get; set; }
        public string Pack { get; set; }
        public string UPC { get; set; }
    }

    public class Location
    {
        public string Name { get; set; }
    }

    public class MeasureUnits
    {
        public string Unit { get; set; }
    }
}

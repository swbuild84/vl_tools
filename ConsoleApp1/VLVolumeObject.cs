using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace ConsoleApp1
{
    [Serializable]
    public class VLVolumeObject
    {
        public int Id;
        public string Code;
        public string Name;
        public double Price;
        public string Unit;
        public string Formula;
        public VLVolumeObject() { }
        public VLVolumeObject(int id, string code, string name, double price, string unit, string formula)
        {
            Id = id;
            Code = code;
            Name = name;
            Price = price;
            Unit = unit;
            Formula = formula;
        }
    }

    [Serializable]
    public class VLVolumeCollection
    {             
        public List<VLVolumeObject> volumes = new List<VLVolumeObject>();
    }
}

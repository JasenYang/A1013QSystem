using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace A1013QSystem.Model
{
    public  class ChipModel
    {      
        public string chipSelect { set; get; }
        public string pathSelect { set;get;}

        public string baudRate { set; get; }
        public string parityCheck { set; get; }
        public string stopBit { set; get; }
        public string byteLength { set; get; }
        public string FIFOSelect { set; get; }
        public string DMAPattern { set; get; }
        public string receiveFIFO { set; get; }
        public string sendTarget { set; get; }
        public string receiveInterrupt { set; get; }
        public string sendInterrupt{ set; get; }
        public string receiveCache { set; get; }
    }
}

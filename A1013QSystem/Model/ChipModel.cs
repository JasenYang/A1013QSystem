using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace A1013QSystem.Model
{
    public  class ChipModel
    {      
        public int chipSelect { set; get; }
        public int pathSelect { set;get;}

        public string baudRate { set; get; }
        public string parityCheck { set; get; }
        public string stopBit { set; get; }
        public string byteLength { set; get; }
        public int FIFOSelect { set; get; }
        public string DMAPattern { set; get; }
        public string receiveFIFO { set; get; }
        public string sendTarget { set; get; }
        public int receiveInterrupt { set; get; }
        public int sendInterrupt{ set; get; }
        public int receiveCache { set; get; }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SetupNew.Models
{
    public class clsSLMP
    {
        public bool CommandOn { get; set; }
        public int CVExtPktNo { get; set; }
        public int RetryCount { get; set; } = 5;
        public int sec { get; set; } = 20;
        public int StdReadStartAddress { get; set; } = 100;
        public int StdReadCount { get; set; } = 100;
        public int StdWriteStartAddress { get; set; } = 200;
        public int StdWriteCount { get; set; } = 100;
        public bool ExtendedRequired { get; set; }=false;
        public int ExtendedReadStartAddress { get; set; } = 1000;
        public int ExtendedReadCount { get; set; } = 700;
        public int NoOfExtendedPackets { get; set; } = 6;
        public int WriteDelayCount { get; set; } = 1;
        public int CVRead { get; set; } = 0;
        public int CommandType { get; set; } = 1;
        public bool PLC_Communication_Error { get; set; } = true;
        public string IPAddress { get; set; }
        public int PortNo { get; set; }


    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SetupNew.Models
{
    internal class ModelAccess
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Password { get; set; }
        public string Email { get; set; }
        public string loginCode { get; set; }
        public enumAccessType AccessType { get; set; }
    }

    internal enum enumAccessType
    {
        Operator = 0,
        Supervisor = 1,
        Admin = 2
    }
}

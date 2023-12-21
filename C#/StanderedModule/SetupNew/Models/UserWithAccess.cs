using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SetupNew.Models
{
    internal class UserWithAccess
    {
        public string Username { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public bool IsAdmin { get; set; }
        public bool IsOperator { get; set; }
        public bool IsSuperWiser { get; set; }
        public bool HasSettingAccess { get; set; }

    }
}

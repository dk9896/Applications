using SetupNew.Interface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SetupNew.Models
{
    internal class Login : ILogin
    {
        public int checkLogin(string username, string password)
        {
            try
            {
                if (username == null || password == null) 
                {
                    return -1;
                }
                if (username == "admin" && password == "admin")
                {
                    return 1;
                }
                else
                {
                    return -1;
                }
            }
            catch
            {

                return -1;
            }

        }
    }
}

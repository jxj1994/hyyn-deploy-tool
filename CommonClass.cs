using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.RightsManagement;
using System.Text;
using System.Threading.Tasks;

namespace hyyn_deploy_tool
{
    internal class CommonClass
    {
    }

    [Serializable]
    class ConnectInfo
    {
        public string dbName { get; set; }
        public string user { get; set; }

        public ConnectInfo()
        {
            this.dbName = "";
            this.user = "";
        }

        public ConnectInfo(string dbName, string user)
        {
            this.dbName = dbName;
            this.user = user;
        }
    }
}

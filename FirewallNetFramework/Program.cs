using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using NetFwTypeLib;

namespace FirewallNetFramework
{
    class Program
    {
        static void Main(string[] args)
        {
            BlockIp("98.139.180.150");
        }

        private static void BlockIp(string ip)
        {
            INetFwPolicy2 firewallPolicy = (INetFwPolicy2)Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FwPolicy2"));
            INetFwRule firewallRule = firewallPolicy.Rules.OfType<INetFwRule>().Where(x => x.Name == "my blacklist").FirstOrDefault();

            if (firewallRule == null)
            {
                firewallRule = (INetFwRule)Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FWRule"));
                firewallRule.Name = "my blacklist";
                firewallPolicy.Rules.Add(firewallRule);
                firewallRule.Description = "Block inbound traffic";
                firewallRule.Profiles = (int)NET_FW_PROFILE_TYPE2_.NET_FW_PROFILE2_ALL;
                firewallRule.Protocol = (int)NET_FW_IP_PROTOCOL_.NET_FW_IP_PROTOCOL_TCP;
                firewallRule.Direction = NET_FW_RULE_DIRECTION_.NET_FW_RULE_DIR_IN;
                firewallRule.Action = NET_FW_ACTION_.NET_FW_ACTION_BLOCK;
                firewallRule.Enabled = true;
                firewallRule.RemoteAddresses = ip;
            }
            else
            {
                var remoteAddresses = firewallRule.RemoteAddresses;
                firewallRule.RemoteAddresses = remoteAddresses + "," + ip;
            }

            Marshal.ReleaseComObject(firewallPolicy);
        }
    }
}

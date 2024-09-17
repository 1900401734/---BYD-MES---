using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HslCommunication;
using HslCommunication.ModBus;
using HslCommunication.Profinet.Omron;
using MesDatasCore;

namespace PlcCommunication
{
    public class cFinsTcp : PlcRW
    {

        OmronFinsNet omronFinsNet = new OmronFinsNet();

        public override bool Connect()
        {
            omronFinsNet.IpAddress = IP;
            omronFinsNet.Port = Port;
            OperateResult result = omronFinsNet.ConnectServer();
            if (result.IsSuccess)
            {
                return true;
            }
            else
            {

                return false;
            }

        }
        public override bool Close()
        {

            OperateResult result = omronFinsNet.ConnectClose();
            if (result == OperateResult.CreateSuccessResult())
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        public override string ReadString(string address, ushort length)
        {
            // OperateResult<string> str = omronFinsNet.ReadString(address, length);
            OperateResult<int[]> str = omronFinsNet.ReadInt32(address, length);

            if (str != null)
            {
                if (str.Content !=null && str.Content.Length > 0)
                {
                    return str.Content[0].ToString();
                }
                return null;
            }
            else
            {
                return null;
            }

        }

        public override void WriteString(string addr, string value)
        {
            if (value != null && value.Length > 0)
            {
                // omronFinsNet.Write(addr, value);
                omronFinsNet.Write(addr, int.Parse(value));
            }

        }
    }

    public class cModbus : PlcRW
    {

        ModbusTcpNet MdbsTcp = new ModbusTcpNet();


        public override bool Connect()
        {
            MdbsTcp.IpAddress = IP;
            MdbsTcp.Port = Port;
            OperateResult result = MdbsTcp.ConnectServer();
            if (result == OperateResult.CreateSuccessResult())
            {
                return true;
            }
            else
            {
                return false;
            }

        }
        public override bool Close()
        {

            OperateResult result = MdbsTcp.ConnectClose();
            if (result == OperateResult.CreateSuccessResult())
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        public override string ReadString(string address, ushort length)
        {
            OperateResult<string> str = MdbsTcp.ReadString(address, length);

            if (str != null)
            {
                return str.Content;
            }
            else
            {
                return null;
            }

        }

        public override void WriteString(string addr, string value)
        {
            MdbsTcp.Write(addr, value);
        }

    }

    public class cPlcTest : PlcRW
    {
        public override bool Connect()
        {
            return true;

        }
        public override bool Close()
        {

            return true;

        }

        public override string ReadString(string address, ushort length)
        {
            return address + length.ToJsonString();

        }

        public override void WriteString(string addr, string value)
        {

        }
    }

    public enum PlcComType
    {
        Modebus,
        FinsTcp
    }
}

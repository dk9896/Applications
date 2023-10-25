using System;
using System.Threading;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.IO;
using System.ComponentModel.Design;
using System.Data;
using SetupNew.Models;
// Other required using statements

public class ClsPlcSLMP
{
    private bool AssignPLCData_Check1 = false;

    public bool AssignPLCData_Check
    {
        get { return AssignPLCData_Check1; }
        set { AssignPLCData_Check1 = value; }
    }
    private int RejCnt;
    private int count;
    public clsSLMP SLMPModel;

    private Thread arival;
    private bool FIRST_CONNECTION;
    
    private TcpClient SOCKET;
    private NetworkStream network_stream;

    private byte[] Writestream;
    private AutoResetEvent autoevent1 = new AutoResetEvent(false);
    private AutoResetEvent autoevent2 = new AutoResetEvent(false);
    private Timer Timer1;
    private Timer Timer2;
    private TimerCallback timer1Delegate;
    private TimerCallback timer2Delegate;
    private byte[] Readstream1 = new byte[21];
    private AsyncCallback evtDataArrival1;
    private AsyncCallback evtDataSent1;
    private byte[] databuffer1 = new byte[3000];

public ClsPlcSLMP()
{
    timer1Delegate = new TimerCallback(timer1_tick);
    timer2Delegate = new TimerCallback(timer2_tick);
    evtDataArrival1 = new AsyncCallback(WinShock_DataArival1);
    evtDataSent1 = new AsyncCallback(WinShock_Datasent1);
}

private void Initialise1()
{
    
}

public void Connect(string ipaddress, string ipport)
{
    //Initialise1();
    Timer1 = new Timer(timer1Delegate, autoevent1, 1000, 1000);
    Timer2 = new Timer(timer2Delegate, autoevent2, Timeout.Infinite, Timeout.Infinite);
    SOCKET = new TcpClient();
    arival = new Thread(START);
    arival.Start();
}

public void Stop_plc()
{
    try
    {
        if (SOCKET.Connected)
        {
            network_stream.Dispose();
            Timer1.Dispose();
            Timer2.Dispose();
            SOCKET.Close();
        }
    }
    catch { }
}

private void START()
{
    Connect_PLC1();
}

public void Connect_PLC1()
{
    try
    {
        Timer1.Change(Timeout.Infinite, Timeout.Infinite);
        SOCKET.Close();
        Thread.Sleep(100);
        SOCKET = new TcpClient();
        SOCKET.Connect(SLMPModel.IPAddress, SLMPModel.PortNo);
        network_stream = SOCKET.GetStream();
        network_stream.Flush();
        Thread.Sleep(100);
        FIRST_CONNECTION = true;
        Timer1.Change(100, 100);
    }
    catch (Exception ex)
    {
        Console.WriteLine("Connect plc Error");
        SLMPModel.PLC_Communication_Error = true;
        Connect_PLC1();
    }
}

public void GetReadArray1(int ReadStartAddress, int NoOfReadRegisters)
{
    byte[] Readstream1 = new byte[21];
    Readstream1[0] = 0x50;
    Readstream1[1] = 0x00;
    Readstream1[2] = 0x00;
    Readstream1[3] = 0xFF;
    Readstream1[4] = 0xFF;
    Readstream1[5] = 0x03; // Lower
    Readstream1[6] = 0x00; // Higher
    Readstream1[7] = 0x0C; // Lower
    Readstream1[8] = 0x00; // High
    Readstream1[9] = 0x00;
    Readstream1[10] = 0x00;
    Readstream1[11] = 0x01; // Read Command
    Readstream1[12] = 0x04;
    Readstream1[13] = 0x00; // Sub Command
    Readstream1[14] = 0x00;
    Readstream1[15] = (byte)(ReadStartAddress % 256);
    Readstream1[16] = (byte)(ReadStartAddress / 256);
    Readstream1[17] = 0x00;
    Readstream1[18] = 0xA8; // D*
    Readstream1[19] = (byte)(NoOfReadRegisters % 256);
    Readstream1[20] = (byte)(NoOfReadRegisters / 256);
}

public void GetWriteArray1(int WriteStartAddress, int NoOfWriteRegisters)
{
    int k;
    int ArraySize;
    int DataToWrite;
    long data;
    long data1;

    ArraySize = (NoOfWriteRegisters * 2) + 21;
    Writestream = new byte[ArraySize];

    Writestream[0] = 0x50;
    Writestream[1] = 0x00;
    Writestream[2] = 0x00;
    Writestream[3] = 0xFF;
    Writestream[4] = 0xFF;
    Writestream[5] = 0x03; // Lower
    Writestream[6] = 0x00; // Higher

    DataToWrite = 12 + (NoOfWriteRegisters * 2);

    Writestream[7] = (byte)(DataToWrite % 256); // Lower
    Writestream[8] = (byte)(DataToWrite / 256); // High

    Writestream[9] = 0x00;
    Writestream[10] = 0x00;
    Writestream[11] = 0x01; // Read Command
    Writestream[12] = 0x14;
    Writestream[13] = 0x00; // Sub Command
    Writestream[14] = 0x00;

    Writestream[15] = (byte)(WriteStartAddress % 256);
    Writestream[16] = (byte)(WriteStartAddress / 256);
    Writestream[17] = 0x00;

    Writestream[18] = 0xA8; // D*

    Writestream[19] = (byte)(NoOfWriteRegisters % 256);
    Writestream[20] = (byte)(NoOfWriteRegisters / 256);

    int J = WriteStartAddress + NoOfWriteRegisters;

    k = 21;

    for (int i = WriteStartAddress; i < J; i++)
    {
        if (SLMPModel.PLCData[i] < 0)
            data = (65536 + SLMPModel.PLCData[i]);
        else
            data = SLMPModel.PLCData[i];

        Writestream[k] = (byte)(data % 256);
        k++;

        Writestream[k] = (byte)(data / 256);
        k++;
    }
}

private void timer1_tick(object sender)
{
    try
    {
        Timer1.Change(Timeout.Infinite, Timeout.Infinite);

        if (SOCKET.Connected && !SLMPModel.CommandOn)
        {
            if (FIRST_CONNECTION)
            {
                Thread.Sleep(500);
                FIRST_CONNECTION = false;
            }

            switch (SLMPModel.CommandType)
            {
                case 1:
                    GetReadArray1(SLMPModel.StdReadStartAddress, SLMPModel.StdReadCount);
                    network_stream.BeginWrite(Readstream1, 0, Readstream1.Length, evtDataSent1, network_stream);
                    SLMPModel.CVRead++;
                    SLMPModel.CommandOn = true;
                    Timer2.Change(3000, 3000);
                    break;
                case 2:
                    if (SLMPModel.PLCData[SLMPModel.StdWriteStartAddress + SLMPModel.StdWriteCount - 1] > 30000)
                            SLMPModel.PLCData[SLMPModel.StdWriteStartAddress + SLMPModel.StdWriteCount - 1] = 0;
                    else
                            SLMPModel.PLCData[SLMPModel.StdWriteStartAddress + SLMPModel.StdWriteCount - 1]++;
                    GetWriteArray1(SLMPModel.StdWriteStartAddress, SLMPModel.StdWriteCount);
                    network_stream.BeginWrite(Writestream, 0, Writestream.Length, evtDataSent1, network_stream);
                        SLMPModel.CommandOn = true;
                    Timer2.Change(3000, 3000);
                    break;
                case 3:
                    GetReadArray1((SLMPModel.ExtendedReadStartAddress + (SLMPModel.ExtendedReadCount * SLMPModel.CVExtPktNo)), SLMPModel.ExtendedReadCount);
                    network_stream.BeginWrite(Readstream1, 0, Readstream1.Length, evtDataSent1, network_stream);
                    SLMPModel.CommandOn = true;
                    Timer2.Change(3000, 3000);
                    break;
                default:
                    SLMPModel.CommandType = 1;
                    Timer1.Change(100, 100);
                    break;
            }

            return;
        }
        else
        {
            Timer1.Change(100, 100);
        }

        if (!SOCKET.Connected)
        {
            Timer1.Change(Timeout.Infinite, Timeout.Infinite);
            Connect_PLC1();
            return;
        }
        else
        {
            SLMPModel.CommandOn = false;
            Timer1.Change(100, 100);
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine("Timer1 Error");
        try
        {
            Timer1.Change(Timeout.Infinite, Timeout.Infinite);
            Connect_PLC1();
        }
        catch
        {
        }
    }
}

private void timer2_tick(object sender)
{
    try
    {
        Timer2.Change(Timeout.Infinite, Timeout.Infinite);
        SLMPModel.PLC_Communication_Error = true;
        SLMPModel.CommandOn = false;
        SLMPModel.CommandType = 1;
        Timer1.Change(1000, 1000);
    }
    catch
    {
    }
}

private void WinShock_Datasent1(IAsyncResult dr)
{
    try
    {
        network_stream.BeginRead(databuffer1, 0, 3000, evtDataArrival1, network_stream);
    }
    catch (Exception ex)
    {
        Console.WriteLine("Winshock Data Sent ERROR");
    }
}

private void WinShock_DataArival1(IAsyncResult dr)
{
    try
    {
        int numberofbytes;
        byte[] SocketData;
        string RegData;
        int i, J, k, L, M, n, ExpectedArraySize, ExtendReadFrom, ExtndedReadFrom, ExpectedLength;
        long Idata, Idata1;
        SLMPModel.PLC_Communication_Error = false;
        SLMPModel.CommandOn = false;

        numberofbytes = network_stream.EndRead(dr);
        SocketData = new byte[numberofbytes];

        if (numberofbytes > 0)
        {
            for (i = 0; i < numberofbytes; i++)
            {
                SocketData[i] = databuffer1[i];
            }

            switch (SLMPModel.CommandType)
            {
                case 1:
                    k = SLMPModel.StdReadCount * 2;
                    ExpectedArraySize = k + 10;

                    if (SocketData.Length == ExpectedArraySize)
                    {
                        if ((SocketData[0] == 0xD0) && (SocketData[3] == 0xFF) && (SocketData[4] == 0xFF) && (SocketData[5] == 0x03))
                        {
                            J = 11;

                            for (i = SLMPModel.StdReadStartAddress; i < (SLMPModel.StdReadStartAddress + SLMPModel.StdReadCount); i++)
                            {
                                M = (int)SocketData[J + 1];
                                n = (int)SocketData[J];
                                Idata = (M * 256) + n;

                                if (Idata > 32767)
                                {
                                    Idata1 = Idata - 65536;
                                }
                                else
                                {
                                    Idata1 = Idata;
                                }

                                SLMPModel.PLCData[i] = (int)Idata1;
                                J += 2;
                            }

                            if (SLMPModel.CVRead == 1)
                            {
                                SLMPModel.CommandType = 2;
                            }

                            if (((SLMPModel.CVRead >= SLMPModel.WriteDelayCount) && ((SLMPModel.PLCData[SLMPModel.StdReadStartAddress + SLMPModel.StdReadCount - 1] == 0) || (SLMPModel.ExtendedRequired == false))))
                            {
                                SLMPModel.CVRead = 0;
                            }

                            if ((SLMPModel.ExtendedRequired == true) && (SLMPModel.PLCData[SLMPModel.StdReadStartAddress + SLMPModel.StdReadCount - 1] > 0))
                            {
                                    SLMPModel.CommandType = 3;
                                    SLMPModel.CVExtPktNo = 0;
                            }

                            AssignPLCData_Check = true;
                        }
                        else
                        {
                            RejCnt++;
                        }
                    }
                    else
                    {
                        RejCnt++;
                    }
                    break;
                case 2:
                    if ((SocketData.Length == 10) && (SocketData[0] == 0xD0) && (SocketData[3] == 0xFF) && (SocketData[4] == 0xFF) && (SocketData[5] == 0x03))
                    {
                        SLMPModel.CommandType = 1;
                        count++;

                        if (count > 32000)
                        {
                            count = 0;
                        }
                    }
                    else
                    {
                        RejCnt++;
                    }
                    break;
                case 3:
                    k = SLMPModel.ExtendedReadCount * 2;
                    ExpectedArraySize = k + 10;

                    if (SocketData.Length == ExpectedArraySize)
                    {
                        if ((SocketData[0] == 0xD0) && (SocketData[3] == 0xFF) && (SocketData[4] == 0xFF) && (SocketData[5] == 0x03))
                        {
                            J = 11;
                            ExtendReadFrom = SLMPModel.ExtendedReadStartAddress + (SLMPModel.ExtendedReadCount * SLMPModel.CVExtPktNo);

                            for (i = ExtendReadFrom; i < (ExtendReadFrom + SLMPModel.ExtendedReadCount); i++)
                            {
                                M = (int)SocketData[J + 1];
                                n = (int)SocketData[J];
                                Idata = (M * 256) + n;

                                if (Idata > 32767)
                                {
                                    Idata1 = Idata - 65536;
                                }
                                else
                                {
                                    Idata1 = Idata;
                                }

                                SLMPModel.PLCData[i] = (int)Idata1;
                                J += 2;
                            }

                            SLMPModel.CVExtPktNo++;

                            if (SLMPModel.CVExtPktNo >= SLMPModel.NoOfExtendedPackets)
                            {
                                SLMPModel.CVExtPktNo = 0;

                                if (SLMPModel.CVRead == 1)
                                {
                                    SLMPModel.CommandType = 2;
                                }
                                else
                                {
                                    SLMPModel.CommandType = 1;
                                }

                                if (SLMPModel.CVRead >= SLMPModel.WriteDelayCount)
                                {
                                    SLMPModel.CVRead = 0;
                                }
                            }

                            AssignPLCData_Check = true;
                        }
                        else
                        {
                            RejCnt++;
                        }
                    }
                    else
                    {
                        RejCnt++;
                    }
                    break;
            }
        }

        network_stream.Flush();
        Timer1.Change(30, 30);
    }
    catch (Exception ex)
    {
        Console.WriteLine("Winshock Data Arrival Error");

        try
        {
            Timer1.Change(30, 30);
        }
        catch
        {
        }
    }
}
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Sockets;
using System.Net;
using System.Threading;
using System.IO;
using System.Runtime.InteropServices;
using System.Collections;
using System.Net.NetworkInformation;

namespace AutoAdjust
{
    class Program
    {
        #region[Windows窗口处理函数]
        [DllImport("user32.dll")]
        private static extern bool SetCursorPos(int X, int Y);//该函数用于设置鼠标的位置,其中X和Y是相对于屏幕左上角的绝对位置.
        [DllImport("user32.dll")]
        private static extern void mouse_event(MouseEventFlag flags, int dx, int dy, uint data, UIntPtr extraInfo);
        [DllImport("User32.dll", EntryPoint = "FindWindow")]
        private extern static IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("user32.dll")]
        static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string strClass, string strWindow);
        [DllImport("user32.dll")]
        static extern bool GetWindowRect(HandleRef hwnd, out NativeRECT rect);
        [DllImport("user32.dll")]
        static extern IntPtr GetTopWindow(IntPtr hWnd);
        [DllImport("user32.dll", SetLastError = true)]
        static extern int CloseWindow(IntPtr hWnd);[DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool DestroyWindow(IntPtr hwnd);
        [DllImport("user32.dll")]
        static extern uint GetMenuItemID(IntPtr hMenu, int nPos);
        [DllImport("user32.dll", EntryPoint = "GetMenu", CharSet = CharSet.Auto)]
        private static extern IntPtr GetMenu(IntPtr hwnd);
        [DllImport("user32.dll", EntryPoint = "GetMenuItemCount", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int GetMenuItemCount(IntPtr hwnd);
        [DllImport("user32.dll", EntryPoint = "GetSubMenu", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetSubMenu(IntPtr hMenu, int nPos);
        [DllImport("user32.dll", EntryPoint = "GetMenuItemRect", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool GetMenuItemRect(IntPtr hwnd, IntPtr hMenu, int posMenu, out NativeRECT rect);
        [DllImport("user32.dll", SetLastError = true)]
        static extern bool GetMenuItemInfo(IntPtr hMenu, uint uItem, bool fByPosition, ref MENUITEMINFO lpmii);
        [DllImport("user32.dll")]
        static extern IntPtr GetDlgItem(IntPtr hDlg, int nIDDlgItem);
        [DllImport("user32.dll")]
        static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
        [DllImport("User32.DLL")]
        public static extern int SendMessage(IntPtr hWnd, uint Msg, int wParam, string lParam);
        [DllImport("user32.dll", EntryPoint = "PostMessage", CharSet = CharSet.Auto)]
        static extern bool PostMessage(IntPtr hWnd, uint Msg, int wParam, int lParam);
        #endregion

        #region[鼠标键盘操作键值]
        private struct NativeRECT
        {
            public int left;
            public int top;
            public int right;
            public int bottom;
        } //将枚举作为位域处理 
        enum MouseEventFlag : uint //设置鼠标动作的键值 
        {
            Move = 0x0001, //发生移动 
            LeftDown = 0x0002, //鼠标按下左键 
            LeftUp = 0x0004, //鼠标松开左键
            RightDown = 0x0008, //鼠标按下右键 
            RightUp = 0x0010, //鼠标松开右键 
            MiddleDown = 0x0020, //鼠标按下中键 
            MiddleUp = 0x0040, //鼠标松开中键 
            XDown = 0x0080,
            XUp = 0x0100,
            Wheel = 0x0800, //鼠标轮被移动
            VirtualDesk = 0x4000, //虚拟桌面 
            Absolute = 0x8000
        }
        public const uint WM_SETTEXT = 0x000C;
        const int BM_CLICK = 0xF5;
        const int WM_CLOSE = 0x0010;
        #endregion
        
        private static byte[] result = new byte[1024];
        static Socket serverSocket;

        static void Main(string[] args)
        {
            //指定utf8格式的文件不含有bom头，因为科傻无法读取含bom头的文件
            var utf8WithoutBom = new System.Text.UTF8Encoding(false);

            // 0.打开科傻平差软件的.exe文件
            System.Diagnostics.Process ps;
            try
            {
                ps = new System.Diagnostics.Process();
                ps.StartInfo.FileName = "D:/科傻平差软件/Cosawin.exe";
                ps.StartInfo.UseShellExecute = false;
                ps.StartInfo.RedirectStandardInput = true;
                ps.StartInfo.RedirectStandardOutput = true;
                ps.StartInfo.Verb = "runas";
                ps.Start();
            }
            catch
            {
                ps = null;
                Console.WriteLine("无法启动cosa程序！");
            }
            Thread.Sleep(1000);

            int error_code = 0;

            while (true)
            {
                String content = "";

                // 1.创建一个服务器端Socket，即ServerSocket，指定绑定的端口，并监听此端口
                //String hostName = Dns.GetHostName();
                //IPHostEntry localhost = Dns.GetHostByName(hostName);//获取IPv4地址即可
                //IPAddress serve_ip = localhost.AddressList[0];
                ////IPAddress serve_ip = new IPAddress(new byte[] { 39,108,189,248 });
                //String ipAddress = serve_ip.ToString();
                //↑这是获得本机外网IP地址的方法，获取openvpn内网IP的方法在GetVPN_IPAddress()中

                IPAddress server_ip = GetVPN_IPAddress("以太网 4");
                if (isPing(server_ip))
                {
                    Console.WriteLine("本服务器IP地址为：{0}", server_ip.ToString());
                }
                else
                {
                    error_code = 1;
                    break;
                }

                serverSocket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);

                EndPoint point = new IPEndPoint(server_ip, 8884);
                serverSocket.Bind(point);  //绑定IP地址：端口  


                serverSocket.Listen(10);    //设定最多10个排队连接请求  
                Console.WriteLine("启动监听{0}接口成功", serverSocket.LocalEndPoint.ToString());

                // 2.调用accept()等待客户端连接
                Console.WriteLine("服务端已就绪，等待客户端接入，服务端ip地址:{0}。等待接收in2文件......\n", server_ip);
                Socket socket_connected = serverSocket.Accept();
                String client_ip = socket_connected.RemoteEndPoint.ToString();
                Console.WriteLine("客户端已找到！IP地址为{0}", client_ip);
                Console.WriteLine("连接已经建立......\n");

                // 3.连接后获取输入流，读取客户端信息
                try
                {
                    //通过Socket接收数据  
                    int receiveNumber = socket_connected.Receive(result);
                    //Console.WriteLine("接收客户端{0}消息{1}", null, Encoding.ASCII.GetString(result, 0, receiveNumber));
                    content = Encoding.ASCII.GetString(result, 0, receiveNumber);
                    Console.WriteLine("客户端：{0}", content);

                    serverSocket.Close();
                    socket_connected.Close();
                }
                catch (Exception ex)
                {
                    if (ex.Source != null)
                    {
                        Console.WriteLine("IOException source: {0}", ex.Source);
                    }
                    //Console.WriteLine(ex.Message);
                    socket_connected.Shutdown(SocketShutdown.Both);
                    socket_connected.Close();
                }

                // 4.将信息写入到文件中
                String[] strings_content = content.Split('\n');
                ArrayList list_content = new ArrayList(strings_content);

                String name = (String)list_content[0];
                list_content.Remove(name);//或者是 .RemoveAt(0)
                Console.WriteLine("读取客户端发送过来的{0}.in2成功\n", name);

                String path = "C:\\" + name;
                String filename_in2 = path + "\\" + name + ".in2";

                if (!Directory.Exists(path))//验证路径是否存在
                {
                    Directory.CreateDirectory(path);
                    //不存在则创建
                }

                FileStream fs;
                if (!File.Exists(filename_in2))
                {
                    fs = new FileStream(filename_in2, FileMode.Create, FileAccess.ReadWrite);
                }
                else
                {
                    fs = new FileStream(filename_in2, FileMode.Open, FileAccess.ReadWrite);
                }
                fs.Close();

                StreamWriter sw = new StreamWriter(filename_in2, false, utf8WithoutBom);
                foreach (String item in list_content)
                {
                    if (!item.Equals(""))
                    {
                        sw.WriteLine(item);
                    }
                }
                sw.Close();

                // 5.在科傻软件中进行平差
                String file_ou2 = Auto_adjust(name, filename_in2);
                if (File.Exists(file_ou2))
                {
                    Console.WriteLine("平差完毕！已生成.ou2文件!");
                }
                //反馈“平差成功”的讯息
                Byte[] check_result = Encoding.UTF8.GetBytes("checked");
                Socket socket_check = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                socket_check.Bind(new IPEndPoint(server_ip, 12345));
                socket_check.Listen(10);
                Socket socket_temp = socket_check.Accept();
                socket_temp.Send(check_result, check_result.Length, SocketFlags.None);
                Console.WriteLine("已发送反馈信息\r\n");
                socket_temp.Close();
                socket_check.Close();

                // 6.传回结果文件
                Byte[] outBuffer = new Byte[1024];
                String out_content = "";

                FileStream fs2 = new FileStream(file_ou2, FileMode.Open, FileAccess.Read);
                StreamReader sr2 = new StreamReader(fs2, Encoding.Default);
                String line = null;
                while ((line = sr2.ReadLine()) != null)
                {
                    out_content += line + "\n";
                }
                outBuffer = Encoding.UTF8.GetBytes(out_content);

                Socket serverSocket2 = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                serverSocket2.Bind(new IPEndPoint(server_ip, 54321));  //绑定IP地址：端口  
                serverSocket2.Listen(10);
                Socket socket2 = serverSocket2.Accept();
                socket2.Send(outBuffer, outBuffer.Length, SocketFlags.None);
                Console.WriteLine("服务器向客户端发送{0}.ou2成功\r\n\r\n", name);

                fs2.Close();
                sr2.Close();
                socket2.Close();
                serverSocket2.Close();

                //7.将COSA软件回调至初始状态
                if (Init_cosa(name))
                {
                    Console.WriteLine("已初始化，可以进行下一次平差！\r\n");
                }else
                {
                    Console.WriteLine("未初始化！\r\n");
                }

                continue;
            }

            if (error_code == 1)
            {
                Console.WriteLine("未连接至VPN服务器！请检查原因，然后重启程序");
                Console.ReadKey();
            }
        }

        public static bool InputStream { get; set; }

        public static String Auto_adjust(String name,String filename_in2)
        {
            String backName_ou2 = "";

            IntPtr ptr_cosa = FindWindow(null, "COSAWIN98－控制测量数据处理通用软件包(CODAPS)");
            if (ptr_cosa != IntPtr.Zero)
            {
                //Console.WriteLine("已找到cosa窗口！");
            }
            else
            {
                Console.WriteLine("未找到cosa窗口！");
            }
            Thread.Sleep(500);

            IntPtr ptrMenu = GetMenu(ptr_cosa);
            int numMenus = GetMenuItemCount(ptrMenu);
            if (ptrMenu == IntPtr.Zero || numMenus <= 0)
            {
                Console.WriteLine("该窗体不包含菜单");
            }

            //获取窗体大小
            NativeRECT wndRec, menuRec_adjust, menuFileSaveRec;
            Point wndPos = new Point();
            Point menPos = new Point();
            Point menuFileSavePos = new Point();
            IntPtr ptrMenuEdit = GetSubMenu(ptrMenu, 1);
            numMenus = GetMenuItemCount(ptrMenuEdit);//菜单包含的子条目
            bool isOK = GetMenuItemRect(ptr_cosa, ptrMenu, 1, out menuRec_adjust); //获取菜单"平差"选项（参数1）的位置
            if (!isOK)
            {
                throw new Win32Exception();
            }
            //Mouse_Click(menuRec_adjust);
            Thread.Sleep(500);

            uint id_edit_find = GetMenuItemID(ptrMenuEdit, 0);//找到"平差"→"平面网"的位置
            PostMessage(ptr_cosa, 0x0111, (int)id_edit_find, 0);//点击"平面网"
            Thread.Sleep(500);

            IntPtr ptrOpenWindow = FindWindow(null, "输入平面观测文件 COSAWIN98");
            IntPtr ptr_EditFileName = FindWindowEx(ptrOpenWindow, IntPtr.Zero, "Edit", null);//找到输入文件名的组件
            int i = SendMessage(ptr_EditFileName, WM_SETTEXT, 0, filename_in2);//输入接收到的文件位置
            Thread.Sleep(500);
            IntPtr ptrOpenButton = FindWindowEx(ptrOpenWindow, IntPtr.Zero, "Button", "打开(&O)");
            i = SendMessage(ptrOpenButton, BM_CLICK, 0, "");//点击“打开(O)按钮”
            Thread.Sleep(500);

            IntPtr ptrErrorWindow = FindWindow(null, "COSAWIN98");
            IntPtr ptrPositiveButton = FindWindowEx(ptrErrorWindow, IntPtr.Zero, "Button", "是(&Y)");

            while (true)//点击所有报错界面的是
            {
                if (ptrPositiveButton != IntPtr.Zero)
                {
                    i = SendMessage(ptrPositiveButton, BM_CLICK, 0, "");
                    //Console.WriteLine("Button:是(&Y)");
                    Thread.Sleep(1000);
                    break;
                }
                else
                {
                    break;
                }
            }
            Thread.Sleep(1000);

            IntPtr ptrFinishWindow = FindWindow("#32770", "平面网平差 COSAWIN98");
            IntPtr ptrFinishButton = FindWindowEx(ptrFinishWindow, IntPtr.Zero, "Button", "确定");
            i = SendMessage(ptrFinishButton, BM_CLICK, 0, "");
            //已生成.ou2文件

            backName_ou2 = "C:\\" + name + "\\" + name + ".ou2";

            return backName_ou2;
        }

        public static Boolean Init_cosa(String ou2_name)
        {
            Boolean isInit = false;
            try
            {
                IntPtr ptr_cosa = FindWindow(null, "COSAWIN98－控制测量数据处理通用软件包(CODAPS) - " + ou2_name + ".ou2");
                //COSAWIN98－控制测量数据处理通用软件包(CODAPS) - trytry.ou2
                IntPtr ptrMenu = GetMenu(ptr_cosa);
                //int numMenus = GetMenuItemCount(ptrMenu);

                //获取窗体大小
                NativeRECT wndRec, menuRec_file, menuFileSaveRec;
                Point wndPos = new Point();
                Point menuFileSavePos = new Point();
                //IntPtr ptr_cosaEdit = GetSubMenu(ptrMenu, 1);
                //numMenus = GetMenuItemCount(ptr_cosaEdit);//菜单包含的子条目
                bool isOK = GetMenuItemRect(ptr_cosa, ptrMenu, 0, out menuRec_file); //获取菜单"文件"选项（参数1）的位置
                if (!isOK)
                {
                    throw new Win32Exception();
                }
                Mouse_Click(menuRec_file);
                Thread.Sleep(500);

                //uint id_edit_find = GetMenuItemID(ptr_cosaEdit, 2);//找到"文件"→"关闭"的位置
                //PostMessage(ptr_cosa, 0x0111, (int)id_edit_find, 0);//点击"关闭"
                //警告：对于点击后不弹出窗口的选项，即后面不带“...”的选项，不能使用PostMessage
                NativeRECT menuRec_close;
                IntPtr ptrMenuFile = GetSubMenu(ptrMenu, 0);
                isOK = GetMenuItemRect(ptr_cosa, ptrMenuFile, 2, out menuRec_close);
                if (!isOK)
                {
                    throw new Win32Exception();
                }
                Mouse_Click(menuRec_close);
                Thread.Sleep(500);

                isInit = true;
            }catch(Exception ex)
            {

            }

            return isInit;
        }

        private static void Mouse_Click(NativeRECT menuRec)
        {
            Point menPos = new Point();
            menPos.X = (menuRec.right + menuRec.left) / 2;
            menPos.Y = (menuRec.top + menuRec.bottom) / 2;
            SetCursorPos(menPos.X, menPos.Y);
            mouse_event(MouseEventFlag.LeftDown | MouseEventFlag.LeftUp, 0, 0, 1, UIntPtr.Zero);//模拟单击
        }

        public static IPAddress GetVPN_IPAddress(String name)
        {
            IPAddress ipaddress = new IPAddress(new byte[] { 39, 108, 189, 248 });

            StringBuilder sb = new StringBuilder();

            // Get a list of all network interfaces (usually one per network card, dialup, and VPN connection)
            NetworkInterface[] networkInterfaces = NetworkInterface.GetAllNetworkInterfaces();

            foreach (NetworkInterface network in networkInterfaces)
            {
                if (network.Name == name)
                {

                    // Read the IP configuration for each network
                    IPInterfaceProperties properties = network.GetIPProperties();

                    // Each network interface may have multiple IP addresses
                    foreach (IPAddressInformation each_address in properties.UnicastAddresses)
                    {
                        // We're only interested in IPv4 addresses for now
                        if (each_address.Address.AddressFamily != AddressFamily.InterNetwork)
                            continue;

                        // Ignore loopback addresses (e.g., 127.0.0.1)
                        if (IPAddress.IsLoopback(each_address.Address))
                            continue;

                        ipaddress = each_address.Address;
                    }
                    break;
                }
            }

            return ipaddress;
        }

        public static Boolean isPing(IPAddress server_ip)
        {
            Boolean pinged = false;

            //Ping 实例对象;
            Ping pingSender = new Ping();
            //ping选项;
            PingOptions options = new PingOptions();
            options.DontFragment = true;
            string data = "ping test data";
            byte[] buf = Encoding.ASCII.GetBytes(data);
            //调用同步send方法发送数据，结果存入reply对象;
            PingReply reply = pingSender.Send(server_ip, 120, buf, options);

            if (reply.Status == IPStatus.Success)
            {
                Console.WriteLine("主机地址::" + reply.Address);
                Console.WriteLine("往返时间::" + reply.RoundtripTime);
                Console.WriteLine("生存时间TTL::" + reply.Options.Ttl);
                Console.WriteLine("缓冲区大小::" + reply.Buffer.Length);
                Console.WriteLine("数据包是否分段::" + reply.Options.DontFragment);

                pinged = true;
            }

            return pinged;
        }
    }
}

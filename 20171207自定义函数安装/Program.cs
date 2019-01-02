using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.IO;
namespace _20171207自定义函数安装
{
    class Program
    {
        private static string dstFileName = "ExcelUdf.xll";//用户电脑上的文件名

        [DllImport("kernel32.dll", SetLastError = true, CallingConvention = CallingConvention.Winapi)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool IsWow64Process([In] IntPtr process, [Out] out bool wow64Process);
        static void Main(string[] args)
        {
            InstallExcelUdfAddins();

        }

        private static void InstallExcelUdfAddins()
        {
            Console.WriteLine("正在安装中，请稍等。");
            dynamic excelComObj;
            bool isExitExcel = false;

            try
            {
                excelComObj = Marshal.GetActiveObject("Excel.Application");
                Console.WriteLine("检测到当前有已打开的Excel程序\r\n请注意关闭Excel程序并保存后内容后再运行，防止数据丢失造成损失！");
                Console.WriteLine("请退出Excel后按Y键继续，，按其他键退出程序。\r\n若未发现Excel程序已打开，同样按Y键程序强制继续，按其他键退出程序。");
                string input = Console.ReadLine();
                if (input.ToUpper() == "Y")
                {
                    isExitExcel = true;
                }
                else
                {
                    return;
                }
            }
            catch (Exception)
            {
                excelComObj = CreateExcelApp();
                isExitExcel = true;
            }



            if (isExitExcel == true)
            {
                bool is64BitWindows = Environment.Is64BitOperatingSystem;
                bool retVal;
                bool is64BitProcess = false;
                foreach (var item in Process.GetProcesses())
                {
                    if (item.ProcessName.ToLower() == "excel")
                    {
                        try
                        {
                            if (is64BitWindows)//仅当系统是64位时才需要判断，若为32位肯定是32位的OFFICE
                            {
                                is64BitProcess = !(IsWow64Process(item.Handle, out retVal) && retVal);
                            }
                        }
                        catch (Exception)
                        {
                        }
                        item.Kill();//强制性退出，已经通知过用户关闭Excel再运行，未关闭完的就强制退出。
                    }

                }
                CopyFileToAddinsDirectoryAndStartAddin(is64BitProcess);
                Console.WriteLine("安装完成，请按任意键退出程序！");
                Console.ReadKey();
            }
        }

        private static dynamic CreateExcelApp()
        {
            dynamic excelComObj;
            System.Type oType = System.Type.GetTypeFromProgID("Excel.Application");
            excelComObj = System.Activator.CreateInstance(oType);
            excelComObj.DisplayAlerts = false;
            return excelComObj;
        }

        private static void CopyFileToAddinsDirectoryAndStartAddin(bool is64BitProcess)
        {
            string dstDir = GetdstDir();
            string srcfilePath = Path.Combine(Path.GetTempPath(),dstFileName );
            byte[] resFilebytes;
            if (is64BitProcess)
            {
                resFilebytes = Properties.Resources.ExcelUDF_AddIn64_packed;

            }
            else
            {
                resFilebytes = Properties.Resources.ExcelUDF_AddIn_packed;
            }

            using (FileStream fileStream = new FileStream(srcfilePath, FileMode.Create))
            {
                fileStream.Write(resFilebytes, 0, resFilebytes.Length);
            }

            File.Copy(srcfilePath, Path.Combine(dstDir, dstFileName), true);

            dynamic excelApp = CreateExcelApp();
            foreach (var item in excelApp.AddIns)
            {
                string addinsName = item.Name;
                if (addinsName==dstFileName)
                {
                    item.Installed = true;
                }
            }
            excelApp.Quit();
            Marshal.ReleaseComObject(excelApp);
            excelApp = null;
            GC.Collect();
        }

        private static string GetdstDir()
        {
            string appDataDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string dstDir = Path.Combine(appDataDir, @"Microsoft\AddIns");
            if (Directory.Exists(dstDir))
            {
                Directory.CreateDirectory(dstDir);
            }

            return dstDir;
        }
    }
}

using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

using Autodesk.RevitAddIns;

namespace RevitAddInDeployer
{
    public class Program
    {
        /// <summary>
        /// 读取配置文件ini
        /// </summary>
        /// <param name="section"></param>
        /// <param name="key"></param>
        /// <param name="def"></param>
        /// <param name="retVal"></param>
        /// <param name="size"></param>
        /// <param name="filePath"></param>
        /// <returns></returns>
        [DllImport("kernel32")]
        public static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);

        [DllImport("kernel32")]
        public static extern int WritePrivateProfileString(string section, string key, string setVal, string filePath);

        public const string INI_FILE_NAME = "Setup.ini";

        public const string CONFIG_ADDIN_CONTENT = "AddInContent";
        public const string CONFIG_ADDIN_TYPE = "Type";
        public const string CONFIG_ADDIN_NAME = "Name";
        public const string CONFIG_ADDIN_ASSEMBLY_NAME = "AssemblyName";
        public const string CONFIG_ADDIN_FULLCLASS_NAME = "FullClassName";
        public const string CONFIG_ADDIN_VENDOR_ID = "VendorId";

        public const string CONFIG_ADDIN_FILE = "AddInFile";
        public const string CONFIG_ADDIN_FILE_NAME = "FileName";

        public const string CONFIG_PLATFORM = "PlatForm";
        public const string CONFIG_VERSION_COUNT = "VersionCount";
        public const string CONFIG_VERSION = "Version";
        public const string CONFIG_ARCH_X86 = "PathX86";
        public const string CONFIG_ARCH_X64 = "PathX64";

        /// <summary>
        /// 错误信息
        /// </summary>
        public static List<string> ErrorMsgSet = new List<string>();
        /// <summary>
        /// 当前程序路径
        /// </summary>
        public static string CurAppDir = "";

        /// <summary>
        /// 版本信息
        /// </summary>
        public struct VersionInfo
        {
            public string appVersion;
            public string pathX86;
            public string pathX64;
        };

        /// <summary>
        /// 配置文件信息
        /// </summary>
        public struct AddInInfo
        {
            /// <summary>
            /// 插件类型（外部应用/外部命令）APP/CMD
            /// </summary>
            public string addInType;
            /// <summary>
            /// 插件名称
            /// </summary>
            public string addInName;
            /// <summary>
            /// 插件dll
            /// </summary>
            public string addInAssemblyName;
            /// <summary>
            /// 程序入口类名
            /// </summary>
            public string addInAssemblyFullClassName;
            /// <summary>
            /// 开发商ID
            /// </summary>
            public string vendorId;
            /// <summary>
            /// AddIn文件名称
            /// </summary>
            public string manifestFileName;
            /// <summary>
            /// 版本数量
            /// </summary>
            public int versionCount;
            /// <summary>
            /// 版本列表
            /// </summary>
            public List<VersionInfo> versionInfo;

            public void InitVersionInfo()
            {
                versionCount = 0;
                versionInfo = new List<VersionInfo>();
            }
        };

        /// <summary>
        /// AddIn文件路径信息
        /// </summary>
        public struct DeployPath
        {
            /// <summary>
            /// Revit AddIn文件存储路径
            /// </summary>
            public string addInFilePath;
            /// <summary>
            /// 插件dll路径
            /// </summary>
            public string addInAssemblyPath;
        }

        /// <summary>
        /// AddIn文件信息
        /// </summary>
        public struct DeployInfo
        {
            public int deployCount;
            public List<DeployPath> deployItem;

            public DeployInfo(bool initTag)
            {
                deployCount = 0;
                deployItem = new List<DeployPath>();
            }

        };

        /// <summary>
        /// 错误信息输出到控制台
        /// </summary>
        public static void ShowErrorMsg()
        {
            Console.WriteLine("发生错误,本次插件安装失败！");

            foreach (string msg in ErrorMsgSet)
            {
                Console.WriteLine(msg);
                Console.Read();
            }

            ErrorMsgSet.Clear();
        }

        /// <summary>
        /// 读取配置文件中属性值（键值对格式）
        /// </summary>
        /// <param name="section">节名称</param>
        /// <param name="key">属性Key值</param>
        /// <param name="retVal">属性Value值</param>
        /// <param name="filePath">配置文件路径</param>
        /// <returns></returns>
        public static bool ReadParamINI(string section, string key, ref string retVal, string filePath)
        {
            int size = 260;

            StringBuilder ItemVal = new StringBuilder(size);

            GetPrivateProfileString(section, key, "", ItemVal, size, filePath);
            retVal = ItemVal.ToString();

            if (retVal == "")
            {
                string msg = "安装配置参数" + key + "不存在或为空!";
                ErrorMsgSet.Add(msg);

                return false;
            }

            return true;
        }

        /// <summary>
        /// 从配置文件读取相关信息
        /// </summary>
        /// <param name="iniFilePath"></param>
        /// <param name="addInInfo"></param>
        /// <returns></returns>
        public static bool GetAddInInfoFromINI(string iniFilePath, ref AddInInfo addInInfo)
        {
            if (!File.Exists(iniFilePath))
            {
                string msg = "安装配置文件" + iniFilePath + "不存在或为空!";
                ErrorMsgSet.Add(msg);

                return false;
            }

            try
            {
                bool readParamOK = true;

                readParamOK &= ReadParamINI(CONFIG_ADDIN_CONTENT, CONFIG_ADDIN_TYPE, ref addInInfo.addInType, iniFilePath);

                readParamOK &= ReadParamINI(CONFIG_ADDIN_CONTENT, CONFIG_ADDIN_NAME, ref addInInfo.addInName, iniFilePath);

                readParamOK &= ReadParamINI(CONFIG_ADDIN_CONTENT, CONFIG_ADDIN_ASSEMBLY_NAME, ref addInInfo.addInAssemblyName, iniFilePath);

                readParamOK &= ReadParamINI(CONFIG_ADDIN_CONTENT, CONFIG_ADDIN_FULLCLASS_NAME, ref addInInfo.addInAssemblyFullClassName, iniFilePath);

                readParamOK &= ReadParamINI(CONFIG_ADDIN_CONTENT, CONFIG_ADDIN_VENDOR_ID, ref addInInfo.vendorId, iniFilePath);

                readParamOK &= ReadParamINI(CONFIG_ADDIN_FILE, CONFIG_ADDIN_FILE_NAME, ref addInInfo.manifestFileName, iniFilePath);

                string versionCountStr = "";
                readParamOK &= ReadParamINI(CONFIG_PLATFORM, CONFIG_VERSION_COUNT, ref versionCountStr, iniFilePath);
                addInInfo.versionCount = Int32.Parse(versionCountStr);

                string curVersionKey = "";
                VersionInfo curVersionInfo = new VersionInfo();

                for (int i = 0; i < addInInfo.versionCount; ++i)
                {
                    curVersionKey = CONFIG_VERSION + "_" + i;
                    readParamOK &= ReadParamINI(CONFIG_PLATFORM, curVersionKey, ref curVersionInfo.appVersion, iniFilePath);

                    curVersionKey = CONFIG_ARCH_X86 + "_" + i;
                    readParamOK &= ReadParamINI(CONFIG_PLATFORM, curVersionKey, ref curVersionInfo.pathX86, iniFilePath);

                    curVersionKey = CONFIG_ARCH_X64 + "_" + i;
                    readParamOK &= ReadParamINI(CONFIG_PLATFORM, curVersionKey, ref curVersionInfo.pathX64, iniFilePath);

                    addInInfo.versionInfo.Add(curVersionInfo);
                }

                return readParamOK;
            }
            catch (Exception e)
            {
                return false;
            }
        }

        /// <summary>
        /// 组织AddIn文件信息
        /// </summary>
        /// <param name="addInInfo"></param>
        /// <param name="deployInfo"></param>
        /// <returns></returns>
        public static bool GetAddInInfoFromProduct(AddInInfo addInInfo, ref DeployInfo deployInfo)
        {
            //获取本机安装的所有Revit信息
            IList<RevitProduct> revitProductArray = RevitProductUtility.GetAllInstalledRevitProducts();
            if (0 == revitProductArray.Count)
            {
                string msg = "在本机上未检测到任何Revit系列软件的安装！";
                ErrorMsgSet.Add(msg);
                return false;
            }

            bool findProduct = false;

            string versionName = "";

            DeployPath curDeployPath = new DeployPath();

            for (int i = 0; i < addInInfo.versionCount; ++i)
            {
                VersionInfo versionInfo = addInInfo.versionInfo[i];

                foreach (RevitProduct product in revitProductArray)
                {
                    versionName = System.Enum.GetName(product.Version.GetType(), product.Version);

                    if (!string.IsNullOrEmpty(versionName) && versionName.Contains(versionInfo.appVersion))
                    {
                        //根据配置文件中版本信息拼接AddIn文件存储路径
                        curDeployPath.addInFilePath = Path.Combine(product.AllUsersAddInFolder, addInInfo.manifestFileName);
                        //拼接安装包解压的dll文件路径
                        if (AddInArchitecture.OS32bit == product.Architecture)
                        {
                            curDeployPath.addInAssemblyPath = Path.Combine(CurAppDir, versionInfo.pathX86, addInInfo.addInAssemblyName);
                        }
                        else if (AddInArchitecture.OS64bit == product.Architecture)
                        {
                            curDeployPath.addInAssemblyPath = Path.Combine(CurAppDir, versionInfo.pathX64, addInInfo.addInAssemblyName);
                        }

                        deployInfo.deployCount++;
                        deployInfo.deployItem.Add(curDeployPath);

                        findProduct = true;
                    }
                }
            }

            if (!findProduct)
            {
                string msg = "在本机上未检测到符合当前需要安装版本的Revit系列软件！";
                ErrorMsgSet.Add(msg);
            }

            return findProduct;
        }

        static int Main(string[] args)
        {
            //获取本机安装的所有Revit
            IList<RevitProduct> revitProductArray = RevitProductUtility.GetAllInstalledRevitProducts();
            if (args.Length == 0)
            {
                if (revitProductArray.Count > 0)
                {
                    RevitProduct product = revitProductArray[0];
                    string addInDirCur = product.AllUsersAddInFolder;
                    int lastIdx = addInDirCur.LastIndexOf("\\");

                    string addInDirAll = addInDirCur.Substring(0, lastIdx);

                    System.Diagnostics.Process.Start("explorer.exe", addInDirAll);
                }
            }
            else if (args.Length == 2)
            {
                CurAppDir = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase;

                string iniFilePath = args[1].ToLower();

                //配置文件路径是相对路径则需加上当前exe所在路径作为根路径
                if (!iniFilePath.Contains("\\") || !Path.IsPathRooted(iniFilePath))
                {
                    iniFilePath = Path.Combine(CurAppDir, iniFilePath);
                }

                AddInInfo addInInfo = new AddInInfo();
                addInInfo.InitVersionInfo();

                //读取配置文件信息
                if (!GetAddInInfoFromINI(iniFilePath, ref addInInfo))
                {
                    ShowErrorMsg();
                    return 1;
                }

                string addInName = addInInfo.addInName;
                string addInAssemblyName = addInInfo.addInAssemblyName;
                string addInAssemblyFullClassName = addInInfo.addInAssemblyFullClassName;
                string vendorId = addInInfo.vendorId;

                string manifestFileName = addInInfo.manifestFileName;

                DeployInfo addInDeployInfo = new DeployInfo(true);
                if (!GetAddInInfoFromProduct(addInInfo, ref addInDeployInfo))
                {
                    ShowErrorMsg();
                    return 1;
                }


                string doWhat = args[0].ToLower();

                //执行插件安装
                if (doWhat.Equals("setup"))
                {
                    for (int i = 0; i < addInDeployInfo.deployCount; ++i)
                    {
                        DeployPath deployPath = addInDeployInfo.deployItem[i];

                        string addInAssemblyPath = deployPath.addInAssemblyPath;
                        string manifestPath = deployPath.addInFilePath;

                        RevitAddInManifest manifest = new RevitAddInManifest();
                        //外部应用
                        if (addInInfo.addInType.ToLower().Equals("app"))
                        {
                            RevitAddInApplication addInApp = new RevitAddInApplication(addInName, addInAssemblyPath, Guid.NewGuid(), addInAssemblyFullClassName, vendorId);
                            manifest.AddInApplications.Add(addInApp);
                        }
                        //外部命令
                        else if (addInInfo.addInType.ToLower().Equals("cmd"))
                        {
                            RevitAddInCommand addInCmd = new RevitAddInCommand(addInAssemblyPath, Guid.NewGuid(), addInAssemblyFullClassName, vendorId);
                            addInCmd.Text = addInInfo.addInName;
                            manifest.AddInCommands.Add(addInCmd);
                        }

                        manifest.SaveAs(manifestPath);
                    }

                    #region 资源文件添加到Revit安装目录下(项目专用，其它项目请删除此段代码)
                    string ermsg = null;
                    string resPath = Path.Combine(CurAppDir, "ProjectResource.dll");
                    if (revitProductArray?.Count > 0 && addInInfo.versionInfo?.Count > 0 && File.Exists(resPath))
                    {
                        foreach (var version in addInInfo.versionInfo)
                        {
                            try
                            {
                                RevitProduct product = revitProductArray.ToList().Find(x => x.Version.ToString().Contains(version.appVersion));
                                if (product != null)
                                {
                                    string curResPath = Path.Combine(product.InstallLocation, "ProjectResource.dll");
                                    if (!File.Exists(curResPath))
                                        File.Copy(resPath, curResPath, true);
                                }
                            }
                            catch (Exception ex)
                            {
                                ermsg = "Revit" + version.appVersion + "版本资源文件添加错误";
                                ErrorMsgSet.Add(ermsg);
                            }
                        }
                    }
                    if (!string.IsNullOrEmpty(ermsg))
                        ShowErrorMsg();
                    #endregion
                }
                //执行插件卸载
                else if (doWhat.Equals("uninst"))
                {
                    for (int i = 0; i < addInDeployInfo.deployCount; ++i)
                    {
                        DeployPath deployPath = addInDeployInfo.deployItem[i];

                        string addInAssemblyPath = deployPath.addInAssemblyPath;
                        string manifestPath = deployPath.addInFilePath;

                        if (File.Exists(manifestPath))
                        {
                            File.Delete(manifestPath);
                        }
                    }

                    #region 删除Revit安装目录下的资源文件(项目专用，其它项目请删除此段代码)
                    string ermsg = null;
                    if (revitProductArray?.Count > 0 && addInInfo.versionInfo?.Count > 0)
                    {
                        foreach (var version in addInInfo.versionInfo)
                        {
                            try
                            {
                                RevitProduct product = revitProductArray.ToList().Find(x => x.Version.ToString().Contains(version.appVersion));
                                if (product != null)
                                {
                                    string curResPath = Path.Combine(product.InstallLocation, "ProjectResource.dll");
                                    if (File.Exists(curResPath))
                                    {
                                        File.Delete(curResPath);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                ermsg = "卸载时Revit" + version.appVersion + "目录下资源文件删除失败";
                                ErrorMsgSet.Add(ermsg);
                            }
                        }
                    }
                    if (!string.IsNullOrEmpty(ermsg))
                        ShowErrorMsg();
                    #endregion
                }
                else
                {
                    string msg = "无效参数：" + doWhat;
                    ErrorMsgSet.Add(msg);

                    ShowErrorMsg();
                }
            }

            return 0;
        }
    }
}

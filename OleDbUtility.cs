using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Runtime.InteropServices;

namespace XlsToAccessImporter
{
    internal static class OleDbUtility
    {
        private static readonly string[] AccessProviders =
        {
            "Microsoft.ACE.OLEDB.16.0",
            "Microsoft.ACE.OLEDB.12.0",
            "Microsoft.Jet.OLEDB.4.0"
        };

        private static readonly string[] ExcelProviders =
        {
            "Microsoft.ACE.OLEDB.16.0",
            "Microsoft.ACE.OLEDB.12.0"
        };

        public static string BuildAccessConnectionString(string accessPath)
        {
            return FindWorkingProvider(AccessProviders, function: delegate(string candidate)
            {
                string connString = string.Format("Provider={0};Data Source={1};Persist Security Info=False;", candidate, accessPath);
                using (OleDbConnection connection = new OleDbConnection(connString))
                {
                    connection.Open();
                    connection.Close();
                }

                return connString;
            });
        }

        public static string EnsureAccessDatabase(string accessPath)
        {
            if (string.IsNullOrWhiteSpace(accessPath))
            {
                throw new InvalidOperationException("Access 文件路径不能为空。");
            }

            if (File.Exists(accessPath))
            {
                return accessPath;
            }

            string extension = Path.GetExtension(accessPath) ?? string.Empty;
            if (!extension.Equals(".mdb", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("当前仅支持自动创建 .mdb 文件，请选择 .mdb 路径或提供已有 Access 文件。");
            }

            string directory = Path.GetDirectoryName(accessPath);
            if (string.IsNullOrWhiteSpace(directory) || !Directory.Exists(directory))
            {
                throw new InvalidOperationException("Access 文件所在目录不存在，请先创建目录：" + (directory ?? string.Empty));
            }

            CreateEmptyMdb(accessPath);
            return accessPath;
        }

        private static void CreateEmptyMdb(string accessPath)
        {
            List<string> failures = new List<string>();
            object catalog = null;
            Type catalogType = Type.GetTypeFromProgID("ADOX.Catalog");
            if (catalogType == null)
            {
                throw new ProviderNotFoundException("未找到 ADOX.Catalog 组件，无法自动创建 MDB 文件。请安装 Access Database Engine 或 MDAC/ADOX 组件。");
            }

            try
            {
                catalog = Activator.CreateInstance(catalogType);
                foreach (string provider in AccessProviders)
                {
                    try
                    {
                        string createConnectionString = string.Format(
                            "Provider={0};Data Source={1};Jet OLEDB:Engine Type=5;",
                            provider,
                            accessPath);
                        catalogType.InvokeMember(
                            "Create",
                            System.Reflection.BindingFlags.InvokeMethod,
                            null,
                            catalog,
                            new object[] { createConnectionString });

                        if (File.Exists(accessPath))
                        {
                            return;
                        }
                    }
                    catch (Exception ex)
                    {
                        failures.Add(provider + ": " + UnwrapMessage(ex));
                        SafeDelete(accessPath);
                    }
                }
            }
            finally
            {
                if (catalog != null && Marshal.IsComObject(catalog))
                {
                    Marshal.FinalReleaseComObject(catalog);
                }
            }

            throw new ProviderNotFoundException(
                "自动创建 MDB 文件失败，请确认本机已安装可用的 Access/Jet 引擎。" + Environment.NewLine + string.Join(Environment.NewLine, failures.ToArray()));
        }

        private static void SafeDelete(string accessPath)
        {
            try
            {
                if (File.Exists(accessPath))
                {
                    File.Delete(accessPath);
                }
            }
            catch
            {
            }
        }

        private static string UnwrapMessage(Exception ex)
        {
            Exception current = ex;
            while (current.InnerException != null)
            {
                current = current.InnerException;
            }

            return current.Message;
        }

        public static string BuildExcelConnectionString(string excelPath, bool hasHeaders)
        {
            string extension = Path.GetExtension(excelPath) ?? string.Empty;
            string extendedProperties;
            if (extension.Equals(".xls", StringComparison.OrdinalIgnoreCase))
            {
                extendedProperties = string.Format("Excel 8.0;HDR={0};IMEX=1", hasHeaders ? "YES" : "NO");
            }
            else
            {
                extendedProperties = string.Format("Excel 12.0 Xml;HDR={0};IMEX=1", hasHeaders ? "YES" : "NO");
            }

            return FindWorkingProvider(ExcelProviders, function: delegate(string candidate)
            {
                string connString = string.Format("Provider={0};Data Source={1};Extended Properties=\"{2}\";", candidate, excelPath, extendedProperties);
                using (OleDbConnection connection = new OleDbConnection(connString))
                {
                    connection.Open();
                    connection.Close();
                }

                return connString;
            });
        }

        private static string FindWorkingProvider(IEnumerable<string> providers, Func<string, string> function)
        {
            List<string> failures = new List<string>();
            foreach (string provider in providers)
            {
                try
                {
                    return function(provider);
                }
                catch (Exception ex)
                {
                    failures.Add(provider + ": " + ex.Message);
                }
            }

            throw new ProviderNotFoundException(
                "未找到可用的 OLE DB 驱动，请安装 Microsoft Access Database Engine。" + Environment.NewLine + string.Join(Environment.NewLine, failures.ToArray()));
        }
    }
}

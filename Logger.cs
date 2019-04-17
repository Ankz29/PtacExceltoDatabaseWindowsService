using NLog;
using NLog.Config;
using NLog.Targets;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace PtacDealerExcelToTableService
{
        class Logger
        {
            private static readonly NLog.Logger _logger = NLog.LogManager.GetCurrentClassLogger();

            private static bool _enabled = false;

            public static int ErrorCount { get; set; }

            public static string ErrorInformation { get; set; }
            public static void Info(string message)
            {
                if (_enabled)
                    _logger.Info(message);
            }

            public static void Warn(string message)
            {
                if (_enabled)
                    _logger.Warn(message);
            }

            public static void Debug(string message)
            {
                if (_enabled)
                    _logger.Debug(message);
            }

            public static void Error(string message)
            {
                if (_enabled)
                    _logger.Error(message);
            }

            public static void Error(Exception x)
            {
                if (_enabled)
                    _logger.Error(BuildExceptionMessage(x));
            }

            public static void Fatal(string message)
            {
                if (_enabled)
                    _logger.Fatal(message);
            }

            public static void Fatal(Exception x)
            {
                if (_enabled)
                    _logger.Fatal(BuildExceptionMessage(x));
            }

            public static string ConfigureLogger(string logPath)
            {
                string result = string.Empty;

                try
                {

                    LoggingConfiguration config = new LoggingConfiguration();

                    FileTarget fileTarget = new FileTarget();
                    config.AddTarget("file", fileTarget);

                    if (!Directory.Exists(logPath))
                        Directory.CreateDirectory(logPath);

                    fileTarget.FileName = Path.Combine(logPath, "Log.txt");
                    fileTarget.ArchiveFileName = Path.Combine(logPath, "Log.{#####}.txt");
                    fileTarget.ArchiveAboveSize = 104857600; //102,400KB = 100MB   //10240; // 10kb
                    fileTarget.ArchiveNumbering = ArchiveNumberingMode.Sequence;
                    fileTarget.ConcurrentWrites = true;
                    fileTarget.KeepFileOpen = false;

                    fileTarget.Layout = "${longdate} | ${level} | ${message}";

                    LoggingRule rule = new LoggingRule("*", LogLevel.Info, fileTarget);

                    config.LoggingRules.Add(rule);


                    LogManager.Configuration = config;
                    result = fileTarget.FileName.ToString();

                    _enabled = true;

                }
                catch (Exception ex)
                {
                    result = ex.Message;
                }
                return result;
            }
            public static string FileName { get; set; }
            public static string BuildExceptionMessage(Exception x)
            {
                Exception logException = x;
                /*
                ErrorInfo.FileName = FileName;
                ErrorInfo.Message = logException.Message;
                ErrorInfo.Source = logException.Source;
                ErrorInfo.StackTrace = logException.StackTrace;
                ErrorInfo.TargetSite = logException.StackTrace;
                */
                if (x.InnerException != null)
                    logException = x.InnerException;

                string strErrorMsg = Environment.NewLine + "Message :" + logException.Message;

                strErrorMsg += Environment.NewLine + "Source :" + logException.Source;

                strErrorMsg += Environment.NewLine + "Stack Trace :" + logException.StackTrace;

                strErrorMsg += Environment.NewLine + "TargetSite :" + logException.TargetSite;

                return strErrorMsg;
            }

        }
    }


using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.IO;
using System.Windows.Automation;    // Automation UI
using System.Windows.Forms;         // SendKeysのために必要

namespace H264toAviConvCLI
{
    class Program
    {
        private static readonly string FFMPEG_EXE_PATH = Directory.GetCurrentDirectory() + @"\Lib\ffmpeg\";
        private static readonly string CONV_EXE_PATH = Directory.GetCurrentDirectory() + @"\Lib\H264_To_AVI_Convertor\Converter.exe";
        private static readonly string WINDOW_TITLE = "H264 to AVI convertor v3.1.5";
        private static readonly int DEFAULT_WAIT_TIME = 300;
        private static readonly int KEY_WAIT_TIME = 20;

        private static readonly string CONVERT_BUTTON_ID = "1";
        private static readonly string SAVING_BUTTON_ID = "1002";
        private static readonly string SOURCE_BUTTON_ID = "1000";
        private static readonly string ALL_BUTTON_ID = "1008";
        private static readonly string OK_BUTTON_ID = "1";
        private static readonly string CONV_DIALOG_OK_BUTTON_ID = "2";
        private static readonly string DIRECTORY_LIST_ID = "14145";

        static void Main(string[] args)
        {
            AutomationElement mainForm;
            AutomationElement subForm;

            Process.Start(CONV_EXE_PATH);
            Process process = SearchProcessByWindowTitle(WINDOW_TITLE);

            // 起動します
            Thread.Sleep(DEFAULT_WAIT_TIME);
            mainForm = AutomationElement.FromHandle(process.MainWindowHandle);


            // Saveボタン押します
            ClickButtonById(mainForm, SAVING_BUTTON_ID);

            var btnClear5 = FindElementById(mainForm, DIRECTORY_LIST_ID);
            Keyin(true, btnClear5, "{Down}");
            Keyin(true, btnClear5, "{Down}");
            Keyin(true, btnClear5, "{Down}");
            Keyin(true, btnClear5, "{ENTER}"); // EnterでMainFormに戻る


            // Sourceボタン押します
            ClickButtonById(mainForm, SOURCE_BUTTON_ID);

            Thread.Sleep(DEFAULT_WAIT_TIME);
            Process process1 = SearchProcessByWindowTitle("フォルダーの参照");
            subForm = AutomationElement.FromHandle(process1.MainWindowHandle);

            var btnClear3 = FindElementById(subForm, DIRECTORY_LIST_ID);
            Keyin(true, btnClear3, "{Down}");
            Keyin(true, btnClear3, "{Down}");

            //Thread.Sleep(DEFAULT_WAIT_TIME);
            ClickButtonById(subForm, OK_BUTTON_ID);


            ClickButtonById(mainForm, ALL_BUTTON_ID);
            ClickButtonById(mainForm, CONVERT_BUTTON_ID);

            Thread.Sleep(DEFAULT_WAIT_TIME + 3000);

            ClickButtonById(mainForm, CONV_DIALOG_OK_BUTTON_ID);


            Console.WriteLine("すべて完了しました");
            Console.ReadLine();

            process.CloseMainWindow();
        }

        //指定したタイトルの文字列が含まれているプロセスを取得
        //一個目を戻すだけなので、複数対応はしていません。
        static public Process SearchProcessByWindowTitle(string title)
        {
            Process process = null;
            foreach (Process p in Process.GetProcesses())
            {
                if (p.MainWindowTitle.Contains(title))
                {
                    process = p;
                    break;
                }
            }
            if (process == null)
            {
                Console.WriteLine(title + "のプロセスが見つかりません。");
            }
            return process;
        }

        // 指定したID属性に一致するAutomationElementを返します
        private static AutomationElement FindElementById(AutomationElement rootElement, string automationId)
        {
            return rootElement.FindFirst(
                TreeScope.Element | TreeScope.Descendants,
                new PropertyCondition(AutomationElement.AutomationIdProperty, automationId));
        }

        private static IEnumerable<AutomationElement> FindAllElementById(AutomationElement rootElement, string automationId)
        {
            return rootElement.FindAll(
                TreeScope.Element | TreeScope.Descendants,
                new PropertyCondition(AutomationElement.AutomationIdProperty, automationId))
                .Cast<AutomationElement>();
        }

        // 指定したName属性に一致するAutomationElementをすべて返します
        private static IEnumerable<AutomationElement> FindAllElementsByName(AutomationElement rootElement, string name)
        {
            return rootElement.FindAll(
                TreeScope.Element | TreeScope.Descendants,
                new PropertyCondition(AutomationElement.NameProperty, name))
                .Cast<AutomationElement>();
        }

        // 指定したName属性に一致するボタン要素をすべて返します
        private static IEnumerable<AutomationElement> FindButtonsByName(AutomationElement rootElement, string name)
        {
            const string BUTTON_CLASS_NAME = "Button";
            return from x in FindAllElementsByName(rootElement, name)
                   where x.Current.ClassName == BUTTON_CLASS_NAME
                   select x;
        }

        public static void Keyin(bool focus, AutomationElement element, string text)
        {
            if (focus)
            {
                element.SetFocus();
            }

            Thread.Sleep(KEY_WAIT_TIME);
            SendKeys.SendWait(text);
            Thread.Sleep(KEY_WAIT_TIME);
            Console.WriteLine("Key " + text + " Input");
        }

        public static void ClickButtonById(AutomationElement rootElement, string automationId)
        {
            Console.WriteLine("Click "+automationId+" Button");
            var button = FindAllElementById(rootElement, automationId).First().GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
            button.Invoke();
        }

        // 指定したControlType属性に一致する要素をすべて返します
        private static AutomationElementCollection FindAllElementsByControlType(AutomationElement rootElement, ControlType controlType)
        {
            return rootElement.FindAll(
                TreeScope.Element | TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, controlType));

            /*
             var mainTable = FindAllElementsByControlType(mainForm, ControlType.Edit);
                ValuePattern txtboxName = (ValuePattern)mainTable[0].GetCurrentPattern(ValuePattern.Pattern);
                txtboxName.SetValue(@"F:\Downloads");
             */
        }

    }
}

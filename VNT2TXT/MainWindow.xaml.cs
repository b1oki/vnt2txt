using System;
using System.Text;
using System.Collections.Specialized;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
namespace VNT2TXT
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void MethodDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                StringCollection filenames = new StringCollection();
                foreach (string line in files)
                {
                    if (string.Compare(Path.GetExtension(line), ".vnt", true) == 0)
                    {
                        filenames.Add(line);
                    }
                }
                foreach (string line in filenames)
                {
                    string lines = File.ReadAllText(line);
                    string text = lines.Substring(lines.IndexOf("ENCODING=QUOTED-PRINTABLE:") + "ENCODING=QUOTED-PRINTABLE:".Length, lines.IndexOf("DCREATED") - lines.IndexOf("ENCODING=QUOTED-PRINTABLE:") - "ENCODING=QUOTED-PRINTABLE:".Length);
                    text = DecodeQuotedPrintable(text);
                    string info = lines.Substring(lines.IndexOf("DCREATED:"), lines.IndexOf("CLASS:") - lines.IndexOf("DCREATED:"));
                    string create = info.Substring(info.IndexOf("DCREATED:") + "DCREATED:".Length, info.IndexOf("LAST-MODIFIED") - "DCREATED:".Length);
                    string modify = info.Substring(info.IndexOf("LAST-MODIFIED:") + "LAST-MODIFIED:".Length);
                    DateTime created = DateTime.ParseExact(create, "yyyyMMdd'T'HHmmss'Z\r\n'", null);
                    DateTime modified = DateTime.ParseExact(modify, "yyyyMMdd'T'HHmmss'Z\r\n'", null);
                    string print = "Создано: " + created.ToString() + "\nИзменено: " + modified.ToString() + "\nТекст заметки:\n" + text;
                    ShowMsg(print);
                }
            }
        }

        void ShowMsg(string input)
        {
            myText.Text = input;
        }

        string DecodeQuotedPrintable(string input)
        {
            string output = Regex.Replace(input, "=\r\n", "", RegexOptions.Compiled);
            byte[] result;
            using (var _stream = new MemoryStream(output.Length))
            {
                int _current_pos = 0;
                while (true)
                {
                    if (_current_pos >= output.Length)
                        break;

                    var _char = output[_current_pos];
                    switch (_char)
                    {
                        case '=':
                            _stream.WriteByte(Convert.ToByte(output.Substring(_current_pos + 1, 2), 16));
                            _current_pos += 3;
                            break;
                        default:
                            _stream.WriteByte(Convert.ToByte(_char));
                            _current_pos += 1;
                            break;
                    }
                }
                result = _stream.ToArray();
            }
            output = Encoding.UTF8.GetString(result);
            return output;
        }
    }
}

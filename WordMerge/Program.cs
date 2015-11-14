using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace WordMerge
{
    class Program
    {
        /// <summary>
        /// プログラムのエントリポイント
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            new Program().Run();
        }

        /// <summary>
        /// 主要処理
        /// </summary>
        void Run()
        {
            var wordMachine = new WordMachine();

            var desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            var mergeInPath = desktop + Path.DirectorySeparatorChar + "MergeIn";

            // デスクトップに直接PDFを書き出すのが便利な気がする
            var tempFile = desktop + Path.DirectorySeparatorChar + "temp.docx";
            var mergeOutPath = desktop + Path.DirectorySeparatorChar + "統合ファイル.pdf";

            // ソートされたワードのファイル一覧取得
            var files = Directory.GetFiles(mergeInPath);
            var wordFiles = files.Where(IsWordWithIndex);
            var withIndex = wordFiles.Select(wf => new { filename = wf, index = GetIndex(wf) });
            var sortedFiles = withIndex.OrderBy(wi => wi.index);

            foreach(var file in sortedFiles)
            {
                Console.Out.WriteLine("[Info]file:" + file);
            }

            // 読み込む
            var docs = sortedFiles.Select(wf => wordMachine.OpenDocument(wf.filename));

            // 一時的にWordドキュメントを生成
            var tempDoc = wordMachine.NewDocument();
            tempDoc.Range(0, 0).Text = " ";

            // マージする
            foreach(var doc in docs)
            {
                tempDoc.Paste(doc);
            }

            // PDFに出力する
            tempDoc.SavePdf(mergeOutPath);

            // tempファイルを保存
            tempDoc.SaveAs2(tempFile);
        }

        /// <summary>
        /// 数字付きのワードのファイル名かどうかを検査
        /// </summary>
        /// <param name="fullpath"></param>
        /// <returns>true:ワードのファイル false:それ以外</returns>
        bool IsWordWithIndex(string fullpath)
        {
            var elements = fullpath.Split(Path.DirectorySeparatorChar);
            var filename = elements.Last();
            var extension = filename.Split('.').LastOrDefault() ?? "NONE";
            var body = filename.Split('.').FirstOrDefault() ?? "NONE";

            var indexStr = body.ToCharArray()
                               .Reverse()
                               .TakeWhile(Char.IsNumber)
                               .Reverse()
                               .CharsToString();

            if (extension.ToLower() != "docx")
            {
                return false;
            }

            try
            {
                int.Parse(indexStr);
            }
            catch(Exception)
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// 数字付きワードファイルのファイル名から数字を取得
        /// </summary>
        /// <param name="fullpath"></param>
        /// <returns></returns>
        int GetIndex(string fullpath)
        {
            var elements = fullpath.Split(Path.DirectorySeparatorChar);
            var filename = elements.Last();
            var body = filename.Split('.').FirstOrDefault() ?? "NONE";

            var indexStr = body.ToCharArray()
                               .Reverse()
                               .TakeWhile(Char.IsNumber)
                               .Reverse()
                               .CharsToString();

            var index = int.Parse(indexStr);

            return index;
        }
    }

    /// <summary>
    /// Wordの状態遷移を管理
    /// </summary>
    class WordMachine
    {
        Word.Documents documents = null;
        List<Object> objects = null;

        public WordMachine()
        {
            var application = new Word.Application();
            application.Visible = false;

            documents = application.Documents;

            objects = new List<Object>();
            objects.Add(application);
            objects.Add(documents);
        }

        ~WordMachine()
        {
            foreach (var obj in objects)
            {
                try
                {
                    if (obj == null)
                    {
                        continue;
                    }

                    if (Marshal.IsComObject(obj) == false)
                    {
                        continue;
                    }

                    Marshal.FinalReleaseComObject(obj);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }
            }
        }

        public Word.Document OpenDocument(string filename)
        {
            var missing = System.Reflection.Missing.Value as Object;
            var objTrue = true as Object;
            var doc = documents.Open(filename, ref missing, ref objTrue);

            objects.Add(doc);
            return doc;
        }

        public Word.Document NewDocument()
        {
            var doc = documents.Add();
            objects.Add(doc);
            return doc;
        }
    }

    /// <summary>
    /// Word.Document拡張
    /// </summary>
    static class WordDocumentExtension
    {
        /// <summary>
        /// WordドキュメントをPDFで保存
        /// </summary>
        /// <param name="document"></param>
        /// <param name="filename"></param>
        static public void SavePdf(this Word.Document document, string filename)
        {
            try
            {
                document.ExportAsFixedFormat(filename, Word.WdExportFormat.wdExportFormatPDF);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        /// <summary>
        /// itemの内容をtargetにペーストする
        /// </summary>
        /// <param name="target"></param>
        /// <param name="item"></param>
        static public void Paste(this Word.Document document, Word.Document item)
        {
            try
            {
                object start = item.Content.Start;
                object end = item.Content.End;
                item.Range(ref start, ref end).Copy();

                var rng = document.Range(document.Content.End - 1, document.Content.End - 1);
                rng.Paste();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        private static int getLastPosition(ref Word.Document document)
        {
            return document.Content.End - 1;
        }
    }

    /// <summary>
    /// IEnumerable<char>拡張
    /// </summary>
    static class IEnumerableCharExtension
    {
        /// <summary>
        /// IEnumerable<char>をstringに変換する
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        static public string CharsToString(this IEnumerable<char> source)
        {
            return new String(source.ToArray());
        }
    }
}

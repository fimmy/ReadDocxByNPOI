// See https://aka.ms/new-console-template for more information
using NPOI.XWPF.UserModel;
using System.Text.RegularExpressions;

var filePath = "xxx.docx";
using (FileStream fs = File.OpenRead(filePath))
{
    XWPFDocument docx = new XWPFDocument(fs);
    try
    {
        SearchDocx(docx);
    }
    catch (Exception ex)
    {
        // 错误处理
        Console.WriteLine(ex.StackTrace);
    }
    finally
    {
        // 关闭docx
        docx.Close();
    }
}

/// <summary>
/// 通过正则匹配文字
/// </summary>
/// <param name="text"></param>
/// <returns></returns>
static Match GetMatch(string text)
{
    var regex = new Regex(@"key\s(\S+)");
    var match = regex.Match(text);
    return match;
}
/// <summary>
/// 查询Docx文档
/// </summary>
/// <param name="document"></param>
static void SearchDocx(XWPFDocument document)
{
    foreach (var paragraph in document.Paragraphs)
    {
        SearchParagraph(paragraph);
    }
    foreach (var table in document.Tables)
    {
        SearchTable(table);
    }
}
/// <summary>
/// 查询段落
/// </summary>
/// <param name="paragraph"></param>
static void SearchParagraph(XWPFParagraph paragraph)
{
    var text = paragraph.Text;
    var match = GetMatch(text);
    if (match.Success)
    {
        // 查询成功，这里可以添加自已的操作
        // todo
    }
}
/// <summary>
/// 查询表格
/// </summary>
/// <param name="table"></param>
static void SearchTable(XWPFTable table)
{
    foreach (var row in table.Rows)
    {
        foreach (var cell in row.GetTableCells())
        {
            if (cell.Paragraphs.Any())
            {
                foreach (var p in cell.Paragraphs)
                {
                    SearchParagraph(p);
                }
            }
            if (cell.Tables.Any())
            {
                foreach (var t in cell.Tables)
                {
                    SearchTable(t);
                }
            }
            var text = table.Text;
            var match = GetMatch(text);
            if (match.Success)
            {
                // 查询成功，这里可以添加自已的操作
                // todo
            }
        }
    }
}
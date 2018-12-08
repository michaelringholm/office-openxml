using System.Collections.Generic;

namespace com.opusmagus.office.openxml
{
    public interface IOpenDocument
    {
        void ReplaceProperties(string sourceDocPath, string targetDocPath, Dictionary<string, string> bookmarks);
        void ReplaceBookmarks(string sourceDocPath, string targetDocPath, Dictionary<string, string> bookmarks);
    }
}
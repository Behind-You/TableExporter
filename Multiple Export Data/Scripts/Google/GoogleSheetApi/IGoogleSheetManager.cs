using System.Collections.Generic;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4.Data;

namespace MyGoogleServices
{
    public interface IGoogleSheetManager : IGoogleManager
    {
        bool ActivateService(UserCredential credential);
        IList<IList<object>> GetSheet(string spreadSheetID, string range);
        ValueRange GetSheetValueRange(string spreadSheetID, string range);
        ValueRange[] GetSheetValueRanges(string spreadSheetID, string sheetName, string[] ranges);
    }
}
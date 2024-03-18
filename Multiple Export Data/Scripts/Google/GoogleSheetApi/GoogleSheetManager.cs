using Google.Apis.Auth.OAuth2;
using Google.Apis.Requests;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;


namespace MyGoogleServices
{
    /// <summary>
    /// Google Sheet 매니저.
    /// </summary>
    public class GoogleSheetManager : IGoogleSheetManager
    {
        public ManagerTypes GetManagerTypes()
        {
            return ManagerTypes.Sheet;
        }

        public GoogleSheetManager()
        {
            _credential = null;
        }

        public GoogleSheetManager(UserCredential credential)
        {
            _credential = credential;
        }

        static UserCredential _credential;
        private SheetsService m_SheetService;
        static private bool SheetService_Enabled = false;
        public BaseClientService GetService()
        {
            if (!SheetService_Enabled)
                return null;

            return m_SheetService;
        }
        public bool ActivateService(UserCredential credential)
        {
            if (_credential == null)
                _credential = credential;
            return ActivateService();
        }
        public bool ActivateService()
        {
            //만약 크리덴셜 파일이 없을경우 실패 반환
            if (_credential == null)
            {
                m_SheetService = null;
                return false;
            }
            // Create Google Sheets API service.
            m_SheetService = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = _credential,
                ApplicationName = "Google Sheets API .NET Quickstart"
            });
            SheetService_Enabled = true;
            return SheetService_Enabled;
        }

        /// <summary>
        /// 시트 고유값과 받아올 시트 범위 입력.
        /// </summary>
        /// <param name="spreadSheetID"></param>
        /// <param name="range"></param>
        /// <returns></returns>
        public IList<IList<Object>> GetSheet(String spreadSheetID, String range)
        {
            if (!SheetService_Enabled)
                ActivateService();
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    m_SheetService.Spreadsheets.Values.Get(spreadSheetID, range);
            ValueRange response = request.Execute();
            return response.Values;
        }
        
        /// <summary>
         /// 시트 고유값과 받아올 시트 범위 입력.
         /// </summary>
         /// <param name="spreadSheetID"></param>
         /// <param name="range"></param>
         /// <returns></returns>
        public ValueRange GetSheetValueRange(String spreadSheetID, String range)
        {
            if (!SheetService_Enabled)
                ActivateService();
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    m_SheetService.Spreadsheets.Values.Get(spreadSheetID, range);
            return request.Execute();
        }
        /// <summary>
         /// 시트 고유값과 받아올 시트 범위 입력.
         /// </summary>
         /// <param name="spreadSheetID"></param>
         /// <param name="range"></param>
         /// <returns></returns>
        public ValueRange[] GetSheetValueRanges(String spreadSheetID,String sheetName, String[] ranges)
        {
            List<ValueRange> result = new List<ValueRange>();
            foreach(var range in ranges)
            {
                if (!SheetService_Enabled)
                    ActivateService();
                string rangeStr = sheetName + "!" + range;
                SpreadsheetsResource.ValuesResource.GetRequest request = m_SheetService.Spreadsheets.Values.Get(spreadSheetID, rangeStr);
                ValueRange valueRange = request.Execute();
                result.Add(valueRange);
            }
            return result.ToArray();
        }

        public void Update(ValueRange valueRange, String spreadSheetID, String range)
        {
            var update = m_SheetService.Spreadsheets.Values.Update(valueRange, spreadSheetID, range);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            update.Execute();
        }

    }
}

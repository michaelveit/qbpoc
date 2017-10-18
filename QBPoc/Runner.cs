using Interop.QBFC13;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QBPoc
{
    public class Runner
    {
        public void DoAction()
        {
            QBSessionManager sessionManager = new QBSessionManager();
            bool booSessionBegun = false;
            try
            {
                booSessionBegun = true;
                sessionManager.OpenConnection("", "QB POC");
                sessionManager.BeginSession("", ENOpenMode.omDontCare);// @"C:\Users\Public\Documents\Intuit\QuickBooks\Company Files\Joe's Business.qbw", ENOpenMode.omMultiUser); //@"C:\Users\Public\Documents\Intuit\QuickBooks\Company Files\Joe's Business.qbw", ENOpenMode.omSingleUser); // , ENOpenMode.omDontCare);
                
                IMsgSetRequest requestSet = getLatestMsgSetRequest(sessionManager);
                /*
                ICheckAdd checkAddRq= requestSet.AppendCheckAddRq();
                checkAddRq.AccountRef.FullName.SetValue("Test Bank");
                checkAddRq.PayeeEntityRef.FullName.SetValue("Gene Simmoms");
                checkAddRq.Memo.SetValue("Test Check");
                checkAddRq.IsToBePrinted.SetValue(true);


                IExpenseLineAdd expenseLineAdd = checkAddRq.ExpenseLineAddList.Append();
                expenseLineAdd.AccountRef.FullName.SetValue("Payroll Expenses");
                expenseLineAdd.Amount.SetValue(100.00);
                expenseLineAdd.Memo.SetValue("CRM");
                checkAddRq.IncludeRetElementList.Add("TxnID");

                IMsgSetResponse checkResponseMsgSet = sessionManager.DoRequests(requestSet);
                IResponse checkResponse = checkResponseMsgSet.ResponseList.GetAt(0);
                System.Diagnostics.Debug.WriteLine(checkResponseMsgSet.ToXMLString());

                requestSet = getLatestMsgSetRequest(sessionManager);
                ICheckQuery query = requestSet.AppendCheckQueryRq();
                IMsgSetResponse responseSet = sessionManager.DoRequests(requestSet);
                System.Diagnostics.Debug.WriteLine(responseSet.ToXMLString());
                */
                requestSet = getLatestMsgSetRequest(sessionManager);
                var rq = requestSet.AppendCustomDetailReportQueryRq();
                rq.CustomDetailReportType.SetValue(ENCustomDetailReportType.cdrtCustomTxnDetail);
                rq.IncludeColumnList.Add(ENIncludeColumn.icClearedStatus);
                rq.IncludeColumnList.Add(ENIncludeColumn.icTxnID);
                rq.ReportAccountFilter.ORReportAccountFilter.ListIDList.Add("8000002C-1508291427");
                rq.ORReportPeriod.ReportPeriod.FromReportDate.SetValue(DateTime.Now.AddHours(-1));
                rq.ORReportPeriod.ReportPeriod.ToReportDate.SetValue(DateTime.Now);
                rq.SummarizeRowsBy.SetValue(ENSummarizeRowsBy.srbItemDetail);

                IMsgSetResponse responseSet = sessionManager.DoRequests(requestSet);
               
                var rp = (IReportRet)responseSet.ResponseList.GetAt(0).Detail;
                System.Diagnostics.Debug.WriteLine("rows " + rp.NumRows.GetValue());
                System.Diagnostics.Debug.WriteLine("cols " + rp.NumColumns.GetValue());
                System.Diagnostics.Debug.WriteLine("list count " + rp.ReportData.ORReportDataList.Count);

                
                IORReportData data;
                for(int i = 0; i < rp.ReportData.ORReportDataList.Count - 1; ++i)
                {
                    data = rp.ReportData.ORReportDataList.GetAt(i);
                    if (data != null)
                    {
                        if (data.DataRow != null)
                        {
                            
                            if (data.DataRow.ColDataList.Count > 1)
                            {
                                if (data.DataRow.ColDataList.GetAt(1).value.GetValue() == "3C3-1508291494")
                                {
                                    System.Diagnostics.Debug.WriteLine(data.DataRow.ColDataList.GetAt(0).value.GetValue());
                                }
                            }
                        }
                    }

                }


            }
            catch (Exception ex)
            {
                Console.Out.WriteLine(ex.Message);
                
            }
            finally
            {
                if (booSessionBegun)
                {
                    sessionManager.EndSession();
                    sessionManager.CloseConnection();
                }

            }

        }

        private IMsgSetRequest getLatestMsgSetRequest(QBSessionManager sessionManager)
        {
            // Find and adapt to supported version of QuickBooks
            double supportedVersion = QBFCLatestVersion(sessionManager);

            short qbXMLMajorVer = 0;
            short qbXMLMinorVer = 0;

            if (supportedVersion >= 6.0)
            {
                qbXMLMajorVer = 6;
                qbXMLMinorVer = 0;
            }
            else if (supportedVersion >= 5.0)
            {
                qbXMLMajorVer = 5;
                qbXMLMinorVer = 0;
            }
            else if (supportedVersion >= 4.0)
            {
                qbXMLMajorVer = 4;
                qbXMLMinorVer = 0;
            }
            else if (supportedVersion >= 3.0)
            {
                qbXMLMajorVer = 3;
                qbXMLMinorVer = 0;
            }
            else if (supportedVersion >= 2.0)
            {
                qbXMLMajorVer = 2;
                qbXMLMinorVer = 0;
            }
            else if (supportedVersion >= 1.1)
            {
                qbXMLMajorVer = 1;
                qbXMLMinorVer = 1;
            }
            else
            {
                qbXMLMajorVer = 1;
                qbXMLMinorVer = 0;
            }

            // Create the message set request object
            IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", qbXMLMajorVer, qbXMLMinorVer);
            return requestMsgSet;

        }

        private double QBFCLatestVersion(QBSessionManager SessionManager)
        {
            // Use oldest version to ensure that this application work with any QuickBooks (US)
            IMsgSetRequest msgset = SessionManager.CreateMsgSetRequest("US", 1, 0);
            msgset.AppendHostQueryRq();
            IMsgSetResponse QueryResponse = SessionManager.DoRequests(msgset);
            //MessageBox.Show("Host query = " + msgset.ToXMLString());
            //SaveXML(msgset.ToXMLString());


            // The response list contains only one response,
            // which corresponds to our single HostQuery request
            IResponse response = QueryResponse.ResponseList.GetAt(0);

            // Please refer to QBFC Developers Guide for details on why 
            // "as" clause was used to link this derrived class to its base class
            IHostRet HostResponse = response.Detail as IHostRet;
            IBSTRList supportedVersions = HostResponse.SupportedQBXMLVersionList as IBSTRList;

            int i;
            double vers;
            double LastVers = 0;
            string svers = null;

            for (i = 0; i <= supportedVersions.Count - 1; i++)
            {
                svers = supportedVersions.GetAt(i);
                vers = Convert.ToDouble(svers);
                if (vers > LastVers)
                {
                    LastVers = vers;
                }
            }
            return LastVers;
        }
    }
}

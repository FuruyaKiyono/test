using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Collections.Specialized;
using System.Xml;
using System.IO;
using System.Text;

public partial class DataViews_HeadOffice_KiwiAPIClient : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        //日付入力フィールドではIMEをOFFに設定
        Form.DefaultButton = Search.UniqueID;
        StartYear.Style.Add("ime-mode", "disabled");
        StartMonth.Style.Add("ime-mode", "disabled");
        StartDay.Style.Add("ime-mode", "disabled");
        EndYear.Style.Add("ime-mode", "disabled");
        EndMonth.Style.Add("ime-mode", "disabled");
        EndDay.Style.Add("ime-mode", "disabled");
    }

    protected void Search_Click(object sender, EventArgs e)
    {
        try
        {
            ActiveDirectoryBoundary adBoundary = new ActiveDirectoryBoundary();
            StringCollection scGroupList = adBoundary.SearchGroup(txtUser.Text, txtPassword.Text);

            if (!CheckAccess(scGroupList))
            {
                MessageLabel.Text = "検索結果がありません。";
                return;
            }

            //ワークフローAPIのバウンダリ
            WorkflowBinding bWorkflow = new WorkflowBinding();

            //IDとパスワードの設定
            bWorkflow.strUserName = "apiuser";
            bWorkflow.strPassword = "resuipa";

            RequestManageFormType[] resFormType;
            WorkflowGetRequestsRequestType reqType = new WorkflowGetRequestsRequestType();
            reqType.manage_request_parameter = new WorkflowGetRequestType();

            //ワークフローの申請フォームIDの設定
            reqType.manage_request_parameter.request_form_id = ddfForm.SelectedItem.Value;
            //決済完了日の設定
            reqType.manage_request_parameter.start_approval_date = DateTime.Parse(StartYear.Text + "/" + StartMonth.Text + "/" + StartDay.Text);
            reqType.manage_request_parameter.start_approval_dateSpecified = true;
            reqType.manage_request_parameter.end_approval_date = DateTime.Parse(EndYear.Text + "/" + EndMonth.Text + "/" + EndDay.Text).AddDays(1);
            reqType.manage_request_parameter.end_approval_dateSpecified = true;
            //完了区分の設定
            reqType.manage_request_parameter.filter = WorkflowGetManageRequestFilter.Complete;
            reqType.manage_request_parameter.filterSpecified = true;

            //申請ID一覧の取得
            resFormType = bWorkflow.WorkflowGetRequests(reqType);

            List<string> lstRequests = new List<string>();

            foreach (XMLElement element in bWorkflow.arrayReturns)
            {
                if (element.Name.Equals("manage_item_detail") && element.NodeType == XmlNodeType.Element)
                {
                    XMLAttribute attrPid = (XMLAttribute)element.arrayAttributes[0];
                    lstRequests.Add(attrPid.Value);
                }
            }

            string[] strRequests = lstRequests.ToArray();

            //バウンダリの結果情報を消去
            bWorkflow.arrayReturns.Clear();

            if (strRequests.Length > 0)
            {
                //申請IDごとにデータ取得
                WorkflowApplicationType[] resApplicationType;
                resApplicationType = bWorkflow.WorkflowGetRequestById(strRequests);

                //結果を値引データCSVへ出力
                CreateCSV(bWorkflow);
            }
            else
            {
                MessageLabel.Text = "検索結果がありません。";
            }
        }
        catch (Exception exception)
        {
            MessageLabel.Text = exception.Message;
        }
    }
    private void CreateCSV(WorkflowBinding bWorkflow)
    {
        Response.ContentType = "text/pain";
        Response.AppendHeader("Content-Disposition", "attachment; filename=" + DateTime.Now.ToString("yyyyMMdd") + Server.UrlEncode(ddfForm.SelectedItem.Text + ".csv"));
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("shift-jis");

        KiwiDemandInformation dInfo = new KiwiDemandInformation();

        //対象区分
        string targetCategory = string.Empty;
        //対象コード
        string targetCode = string.Empty;
        //管理種別
        string Category = string.Empty;
        //処理区分
        string ProcessingCategory = string.Empty;
        //対象名称
        string targetName = string.Empty;
        //対象金額
        string targetPrice = string.Empty;
        //精算期間
        string Period = string.Empty;
        //精算区分
        string AdjustmentCode = string.Empty;
        //期初月
        string BeginningMonth = string.Empty;
        //精算条件
        string Adjustment = string.Empty;
        //精算方法
        string AdjustmentMethod = string.Empty;
        //開始日付
        string StartDay = string.Empty;
        //打切日付
        string EndDay = string.Empty;
        //精算条件コードを新規登録するかどうか
        string NewAdjustment = string.Empty;
        //支払予定日（月ずれ）
        string PaymentMonth = string.Empty;
        //支払予定日（日付）
        string PaymentDay = string.Empty;
        //変更日付
        string ChangeDay = string.Empty;
        //メーカーコード
        string Maker = string.Empty;
        //メーカー名
        string MakerName = string.Empty;


        //タイトル
        List<string> lstTitle = new List<string>();
        //値引・リベート申請の明細行
        List<string> lstLine = new List<string>();
        //通常申請のアイテム
        List<string> lstItem = new List<string>();
        //決済情報
        List<string> lstProcess = new List<string>();

        //１申請目だけタイトル行をつける
        int nDemand = 0;

        foreach (XMLElement element in bWorkflow.arrayReturns)
        {
            //申請情報
            if (element.Name.Equals("application") && element.NodeType == XmlNodeType.Element)
            {
                for (int i = 0; i < element.arrayAttributes.Count; i++)
                {
                    XMLAttribute attribute = (XMLAttribute)element.arrayAttributes[i];

                    if (attribute.Name.Equals("date"))
                    {
                        dInfo.DemandDate = attribute.Value;
                    }
                    else if (attribute.Name.Equals("number"))
                    {
                        dInfo.DemandNumber = attribute.Value;
                    }
                }
            }
            //申請者情報
            else if (element.Name.Equals("applicant"))
            {
                for (int i = 0; i < element.arrayAttributes.Count; i++)
                {
                    XMLAttribute attribute = (XMLAttribute)element.arrayAttributes[i];

                    if (attribute.Name.Equals("name"))
                    {
                        dInfo.Operator = attribute.Value;
                    }
                }
            }
            //申請内容
            else if (element.Name.Equals("item"))
            {
                switch (ddfForm.SelectedItem.Text)
                {
                    case "値引申請":
                        AnalyzeDiscountItem(element, ref dInfo, ref targetCategory, ref targetCode, ref Category, ref ProcessingCategory, lstLine);
                        break;
                    case "サロンリベート登録申請":
                        AnalyzeSalonRebateItem(element, ref dInfo, ref targetCategory, ref targetCode, ref targetName, ref targetPrice, ref Period, ref BeginningMonth, ref Adjustment, ref AdjustmentMethod, ref AdjustmentCode, ref StartDay, ref EndDay, ref ProcessingCategory, ref NewAdjustment, ref PaymentMonth, ref PaymentDay, lstLine);　//追加抽出項目を追加する
                        break;
                    //case "セルアウトリベート登録申請":
                        //AnalyzeSelloutRebateItemOld(element, ref dInfo, ref targetCategory, ref targetCode, ref targetName, ref targetPrice, ref Period, ref BeginningMonth, ref Adjustment, ref AdjustmentMethod, ref StartDay, ref EndDay, ref ProcessingCategory, lstLine);
                        //break;
                    case "セルアウトリベート登録申請":
                        AnalyzeSelloutRebateItem(element, ref dInfo, ref targetCategory, ref targetCode, ref targetName, ref targetPrice, ref Period, ref BeginningMonth, ref Adjustment, ref AdjustmentMethod, ref StartDay, ref EndDay, ref ProcessingCategory, ref Maker, ref MakerName, lstLine);
                        break;
                    case "担当変更申請":
                        AnalyzeChangeOfPersonItem(element, ref dInfo, ref ChangeDay, lstLine);
                        break;
                    default:
                        lstItem.Add("\"" + ((XMLAttribute)element.arrayAttributes[1]).Value + "\"");
                        if (nDemand == 0)
                        {
                            lstTitle.Add(((XMLAttribute)element.arrayAttributes[0]).Value);
                        }

                        break;
                }
            }
            //決済情報
            else if (element.Name.Equals("processor") && element.NodeType == XmlNodeType.Element)
            {
                string strDate = string.Empty;
                string strProcessor = string.Empty;
                string strResult = string.Empty;

                for (int i = 0; i < element.arrayAttributes.Count; i++)
                {
                    XMLAttribute attribute = (XMLAttribute)element.arrayAttributes[i];

                    switch (attribute.Name)
                    {
                        case "date":
                            strDate = attribute.Value;
                            break;
                        case "processor_name":
                            strProcessor = attribute.Value;
                            break;
                        case "result":
                            strResult = attribute.Value;
                            break;
                    }
                }

                if (nDemand == 0)
                {
                    lstTitle.Add("決済者,決済日付,結果");
                }
                lstProcess.Add(strProcessor + "," + strDate + "," + strResult);
            }
            //申請終了
            else if (element.Name.Equals("application") && element.NodeType == XmlNodeType.EndElement)
            {
                switch (ddfForm.SelectedItem.Text)
                {
                    case "値引申請":
                    case "サロンリベート登録申請":
                    case "セルアウトリベート登録申請":
                    //case "セルアウトリベート登録申請（NEW）":
                    case "担当変更申請":
                    foreach (string strLine in lstLine)
                        {
                            string line = strLine;

                            foreach (string strProcess in lstProcess)
                            {
                                line += "," + strProcess;
                            }
                            Response.Write(line);
                            Response.Write("\r\n");
                        }
                        break;
                    default:
                        string swLine = string.Empty;
                        string TitleLine = string.Empty;

                        foreach (string strItem in lstItem)
                        {
                            swLine += strItem + ",";
                        }

                        for (int i = 0; i < lstProcess.Count; i++)
                        {
                            swLine += lstProcess[i];

                            if (i < lstProcess.Count - 1)
                            {
                                swLine += ",";
                            }
                        }

                        if (nDemand == 0)
                        {
                            Response.Write("申請番号,申請日,申請者,所属");
                            foreach (string title in lstTitle)
                            {
                                Response.Write("," + title);
                            }
                            Response.Write(System.Environment.NewLine);
                        }

                        Response.Write(dInfo.DemandNumber + "," + dInfo.DemandDate + "," + dInfo.Operator + "," + dInfo.Division + "," + swLine);
                        Response.Write("\r\n");

                        nDemand++;
                        break;
                }

                lstLine.Clear();
                lstProcess.Clear();
                lstItem.Clear();
            }
        }

        Response.Flush();
        Response.End();

        string scriptStr = "<script type='text/javascript'>";
        scriptStr += "window.open('about:blank','_self').close();";
        scriptStr += "</script>";

        ClientScript.RegisterClientScriptBlock(GetType(), "closewindow", scriptStr);
    }

    //値引申請のItem解析
    private void AnalyzeDiscountItem(XMLElement element, ref KiwiDemandInformation dInfo, ref string targetCategory, ref string targetCode, ref string Category, ref string ProcessingCategory, List<string> lstLine)
    {
        XMLAttribute attribute = (XMLAttribute)element.arrayAttributes[0];

        switch (attribute.Value)
        {
            case "所属":
                dInfo.Division = ((XMLAttribute)element.arrayAttributes[1]).Value;
                break;

            case "処理区分":
                ProcessingCategory = ((XMLAttribute)element.arrayAttributes[1]).Value;
                break;

            case "対象区分":
                targetCategory = ((XMLAttribute)element.arrayAttributes[1]).Value;
                targetCategory = targetCategory.Replace("得意先", "1");
                targetCategory = targetCategory.Replace("請求先", "2");
                targetCategory = targetCategory.Replace("入金親", "3");
                break;

            case "対象コード":
                targetCode = ((XMLAttribute)element.arrayAttributes[1]).Value;
                break;

            case "管理種別":
                Category = ((XMLAttribute)element.arrayAttributes[1]).Value;
                Category = Category.Replace("商品", "1");
                Category = Category.Replace("ラインアップ", "2");
                Category = Category.Replace("ブランド", "3");
                Category = Category.Replace("メーカー", "4");
                Category = Category.Replace("売上種別", "5");
                break;

            case "条件1":
            case "条件2":
            case "条件3":
            case "条件4":
            case "条件5":
            case "条件6":
            case "条件7":
            case "条件8":
            case "条件9":
            case "条件10":
                lstLine.Add(dInfo.DemandNumber + "," + dInfo.DemandDate + "," + dInfo.Operator + "," + dInfo.Division + "," + ProcessingCategory + "," + targetCategory + "," + targetCode + "," + Category + "," + ((XMLAttribute)element.arrayAttributes[1]).Value);
                break;
        }
    }

    //サロンリベートのItem解析
    private void AnalyzeSalonRebateItem(XMLElement element, ref KiwiDemandInformation dInfo, ref string targetCategory, ref string targetCode, ref string targetName, ref string targetPrice, ref string Period, ref string BeginningMonth, ref string Adjustment, ref string AdjustmentMethod, ref string AdjustmentCode, ref string StartDay, ref string EndDay, ref string ProcessingCategory, ref string NewAdjustment, ref string PaymentMonth, ref string PaymentDay, List<string> lstLine)　//追加抽出項目を追加する
    {
        XMLAttribute attribute = (XMLAttribute)element.arrayAttributes[0];

        switch (attribute.Value)
        {
            case "所属":
                dInfo.Division = ((XMLAttribute)element.arrayAttributes[1]).Value;
                break;

            case "処理区分":
                ProcessingCategory = ((XMLAttribute)element.arrayAttributes[1]).Value;
                break;

            case "対象区分":
                targetCategory = ((XMLAttribute)element.arrayAttributes[1]).Value;
                targetCategory = targetCategory.Replace("得意先", "T");
                targetCategory = targetCategory.Replace("請求先", "G");
                targetCategory = targetCategory.Replace("入金親", "O");
                targetCategory = targetCategory.Replace("団体", "D");
                break;

            case "対象コード":
                targetCode = ((XMLAttribute)element.arrayAttributes[1]).Value;
                break;

            case "対象名称":
                targetName = ((XMLAttribute)element.arrayAttributes[1]).Value;
                break;

            case "対象金額":
                targetPrice = ((XMLAttribute)element.arrayAttributes[1]).Value;
                targetPrice = targetPrice.Replace("値引後", "1");
                targetPrice = targetPrice.Replace("値引前", "2");
                break;

            case "精算方法":
                AdjustmentMethod = ((XMLAttribute)element.arrayAttributes[1]).Value;
                AdjustmentMethod = AdjustmentMethod.Replace("現金", "0");
                AdjustmentMethod = AdjustmentMethod.Replace("商品", "1");
                break;

            case "精算区分":
                AdjustmentCode = ((XMLAttribute)element.arrayAttributes[1]).Value;
                AdjustmentCode = AdjustmentCode.Replace("月次締", "0");
                AdjustmentCode = AdjustmentCode.Replace("請求締", "1");
                break;

            case "精算期間":
                Period = ((XMLAttribute)element.arrayAttributes[1]).Value;
                Period = Period.Replace("毎月", "0");
                Period = Period.Replace("四半期", "1");
                Period = Period.Replace("半期", "2");
                Period = Period.Replace("通期", "3");
                break;

            case "計算期初月":
                BeginningMonth = ((XMLAttribute)element.arrayAttributes[1]).Value;
                BeginningMonth = BeginningMonth.Replace(" 月", string.Empty);
                break;

            //2015年2月　追加項目"お支払予定日(月ずれ)"
            case "お支払予定日(月ずれ)":
                PaymentMonth = ((XMLAttribute)element.arrayAttributes[1]).Value;
                PaymentMonth = PaymentMonth.Replace("翌月", "1");
                PaymentMonth = PaymentMonth.Replace("翌々月", "2");
                PaymentMonth = PaymentMonth.Replace("3か月後", "3"); //
                break;

            //2015年2月　追加項目"お支払予定日"
            case "お支払予定日":
                PaymentDay = ((XMLAttribute)element.arrayAttributes[1]).Value;
                PaymentDay = PaymentDay.Replace("10 ", "10");
                PaymentDay = PaymentDay.Replace("20 ", "20");
                PaymentDay = PaymentDay.Replace("末 ", "99");
                PaymentDay = PaymentDay.Replace("日払い", string.Empty);
                break;

            //2015年2月　追加項目"リベート精算条件"
            case "リベート精算条件":
                NewAdjustment = ((XMLAttribute)element.arrayAttributes[1]).Value;
                NewAdjustment = NewAdjustment.Replace("新規登録した精算条件番号を設定する", "0");
                NewAdjustment = NewAdjustment.Replace("既に登録されている精算条件番号を設定する", "1");
                break;

            case "リベート精算条件番号":
                Adjustment = ((XMLAttribute)element.arrayAttributes[1]).Value;
                break;

            case "開始日付":
                StartDay = ((XMLAttribute)element.arrayAttributes[1]).Value;
                StartDay = StartDay.Replace("年 ", "/");
                StartDay = StartDay.Replace("月 ", "/");
                StartDay = StartDay.Replace("日", string.Empty);
                break;

            case "打切日付":
                EndDay = ((XMLAttribute)element.arrayAttributes[1]).Value;
                EndDay = EndDay.Replace("年 ", "/");
                EndDay = EndDay.Replace("月 ", "/");
                EndDay = EndDay.Replace("日", string.Empty);
                break;


            case "条件1":
            case "条件2":
            case "条件3":
            case "条件4":
            case "条件5":
            case "条件6":
            case "条件7":
            case "条件8":
            case "条件9":
            case "条件10":
                string strCondition = ((XMLAttribute)element.arrayAttributes[1]).Value;
                strCondition = strCondition.Replace("売上種別", "0");
                strCondition = strCondition.Replace("メーカー", "1");
                strCondition = strCondition.Replace("ブランド", "2");
                strCondition = strCondition.Replace("ラインアップ", "3");
                strCondition = strCondition.Replace("商品", "4");
                strCondition = strCondition.Replace("総合計", "5");
                strCondition = strCondition.Replace("含む", "1");
                strCondition = strCondition.Replace("含まない", "0");
                strCondition = strCondition.Replace(" ,", ",");
                strCondition = strCondition.Replace(", ", ",");

                lstLine.Add(dInfo.DemandNumber + "," + dInfo.DemandDate + "," + dInfo.Division + "," + dInfo.Operator + "," + ProcessingCategory + "," + targetCategory + "," + targetCode + "," + targetName + "," + targetPrice + "," + Period + "," + BeginningMonth + "," + AdjustmentMethod + "," + AdjustmentCode + "," + Adjustment + "," + StartDay + "," + EndDay + ",1," + NewAdjustment + "," + PaymentMonth + "," + PaymentDay + "," + strCondition);　//追加抽出項目を追加する
                break;
        }
    }

    //セルアウトリベートのItem解析
    private void AnalyzeSelloutRebateItem(XMLElement element, ref KiwiDemandInformation dInfo, ref string targetCategory, ref string targetCode, ref string targetName, ref string targetPrice, ref string Period, ref string BeginningMonth, ref string Adjustment, ref string AdjustmentMethod, ref string StartDay, ref string EndDay, ref string ProcessingCategory, ref string Maker, ref string MakerName, List<string> lstLine)
    {
        XMLAttribute attribute = (XMLAttribute)element.arrayAttributes[0];

        switch (attribute.Value)
        {
            case "所属":
                dInfo.Division = ((XMLAttribute)element.arrayAttributes[1]).Value;
                break;

            case "処理区分":
                ProcessingCategory = ((XMLAttribute)element.arrayAttributes[1]).Value;
                break;

            case "対象区分":
                targetCategory = ((XMLAttribute)element.arrayAttributes[1]).Value;
                targetCategory = targetCategory.Replace("得意先", "T");
                targetCategory = targetCategory.Replace("請求先", "G");
                targetCategory = targetCategory.Replace("入金親", "O");
                targetCategory = targetCategory.Replace("団体", "D");
                break;

            case "対象コード":
                targetCode = ((XMLAttribute)element.arrayAttributes[1]).Value;
                break;

            case "対象名称":
                targetName = ((XMLAttribute)element.arrayAttributes[1]).Value;
                break;

            case "対象金額":
                targetPrice = ((XMLAttribute)element.arrayAttributes[1]).Value;
                targetPrice = targetPrice.Replace("仕入価", "1");
                targetPrice = targetPrice.Replace("サロン価", "2");
                break;

            case "精算方法":
                AdjustmentMethod = ((XMLAttribute)element.arrayAttributes[1]).Value;
                AdjustmentMethod = AdjustmentMethod.Replace("現金", "0");
                AdjustmentMethod = AdjustmentMethod.Replace("商品", "1");
                break;

            case "精算期間":
                Period = ((XMLAttribute)element.arrayAttributes[1]).Value;
                Period = Period.Replace("毎月", "0");
                Period = Period.Replace("四半期", "1");
                Period = Period.Replace("半期", "2");
                Period = Period.Replace("通期", "3");
                break;

            case "計算期初月":
                BeginningMonth = ((XMLAttribute)element.arrayAttributes[1]).Value;
                BeginningMonth = BeginningMonth.Replace(" 月", string.Empty);
                break;

            case "メーカー精算条件番号":
                Adjustment = ((XMLAttribute)element.arrayAttributes[1]).Value;
                break;

            case "開始日付":
                StartDay = ((XMLAttribute)element.arrayAttributes[1]).Value;
                StartDay = StartDay.Replace("年 ", "/");
                StartDay = StartDay.Replace("月 ", "/");
                StartDay = StartDay.Replace("日", string.Empty);
                break;

            case "打切日付":
                EndDay = ((XMLAttribute)element.arrayAttributes[1]).Value;
                EndDay = EndDay.Replace("年 ", "/");
                EndDay = EndDay.Replace("月 ", "/");
                EndDay = EndDay.Replace("日", string.Empty);
                break;

            case "メーカーコード":
                Maker = ((XMLAttribute)element.arrayAttributes[1]).Value;
                break;

            case "メーカー名":
                MakerName = ((XMLAttribute)element.arrayAttributes[1]).Value;
                break;

            case "条件1":
            case "条件2":
            case "条件3":
            case "条件4":
            case "条件5":
            case "条件6":
            case "条件7":
            case "条件8":
            case "条件9":
            case "条件10":
                string strCondition = ((XMLAttribute)element.arrayAttributes[1]).Value;
                strCondition = strCondition.Replace("メーカー", "メーカー,");
                strCondition = strCondition.Replace("ブランド", "ブランド,");
                strCondition = strCondition.Replace("ラインアップ", "ラインアップ,");
                strCondition = strCondition.Replace("商品", "商品,");
                strCondition = strCondition.Replace(" ,", ",");
                strCondition = strCondition.Replace(", ", ",");


                lstLine.Add(dInfo.DemandNumber + "," + dInfo.DemandDate + "," + dInfo.Division + "," + dInfo.Operator + "," + ProcessingCategory + "," + targetCategory + "," + targetCode + "," + targetName + "," + targetPrice + "," + Period + "," + BeginningMonth + "," + AdjustmentMethod + "," + Adjustment + "," + StartDay + "," + EndDay + "," + Maker + "," + MakerName + "," + strCondition);
                break;
        }
    }
    

    //担当変更のItem解析
    private void AnalyzeChangeOfPersonItem(XMLElement element, ref KiwiDemandInformation dInfo, ref string ChangeDay, List<string> lstLine)
    {
        XMLAttribute attribute = (XMLAttribute)element.arrayAttributes[0];

        switch (attribute.Value)
        {
            case "所属":
                dInfo.Division = ((XMLAttribute)element.arrayAttributes[1]).Value;
                break;

            case "変更日付":
                ChangeDay = ((XMLAttribute)element.arrayAttributes[1]).Value;
                break;

            case "1":
            case "2":
            case "3":
            case "4":
            case "5":
            case "6":
            case "7":
            case "8":
            case "9":
            case "10":
                lstLine.Add(dInfo.DemandNumber + "," + dInfo.DemandDate + "," + dInfo.Operator + "," + dInfo.Division + "," + ChangeDay + "," + ((XMLAttribute)element.arrayAttributes[1]).Value);
                break;
        }
    }

    //アクセス権のチェック
    private bool CheckAccess(StringCollection scGroupList)
    {
        switch (ddfForm.SelectedItem.Text)
        {

            case "値引申請":
            case "新規得意先申請":
            case "得意先変更申請":
            case "担当変更申請":
            case "情報機器外部利用申請":
            case "IT機器設置変更依頼書":
            case "IT機器紛失届":
            case "フォルダ設定依頼書":
            case "IT関連支出申請":
            case "一括値引申請":
            case "一括担当変更申請":
                foreach (string strGroup in scGroupList)
                {
                    if (strGroup.Contains("IT戦略部") || strGroup.Contains("業務推進部長") || strGroup.Contains("取締役（個人指定）") || strGroup.Contains("法務・庶務"))
                    {
                        return true;
                    }
                }
                break;

            case "販売促進申請":
            case "仕入先契約申請":
            case "債権管理申請":
            case "投資・資産取得申請":
            case "PR・イベント・教育費用申請":
            case "福利厚生申請":
            case "住所変更及び通勤手当変更申請":
            case "氏名変更届":
            case "家族異動届":
            case "人事（組織）情報申請":
            case "返品申請":
            case "公休出勤管理申請(振替・食事代)":
            case "公休出勤管理申請(時間外・振休)":
            case "社章購入申請":
            case "死亡報告書":
            case "源泉徴収票発行依頼":
                foreach (string strGroup in scGroupList)
                {
                    if (strGroup.Contains("人事課") || strGroup.Contains("取締役（個人指定）"))
                    {
                        return true;
                    }
                }
                break;

            case "入金値引申請":
            case "返金申請":
            case "請求書発行依頼（個別申請）":
            case "請求書発行依頼（一括申請）":
            case "納品書発行依頼（個別申請）":
                foreach (string strGroup in scGroupList)
                {
                    if (strGroup.Contains("経理部") || strGroup.Contains("取締役（個人指定）") || strGroup.Contains("IT戦略部"))
                    {
                        return true;
                    }
                }
                break;

            case "法務・会社庶務申請":
            case "公休日業務　車両使用申請":
            case "公休日私用　車両使用申請":
            case "車両事故・交通違反報告書":
            case "モバイル端末報告書(携帯電話・i-pad・ﾎﾟｹｯﾄWi-Fi）":
                foreach (string strGroup in scGroupList)
                {
                    if (strGroup.Contains("法務・庶務") || strGroup.Contains("取締役（個人指定）"))
                    {
                        return true;
                    }
                }
                break;

            case "サロンリベート登録申請":
                foreach (string strGroup in scGroupList)
                {
                    if (strGroup.Contains("IT戦略部") || strGroup.Contains("取締役（個人指定）"))
                    {
                        return true;
                    }
                }
                break;

            case "セルアウトリベート登録申請":
                foreach (string strGroup in scGroupList)
                {
                    if (strGroup.Contains("IT戦略部") || strGroup.Contains("経理部") || strGroup.Contains("取締役（個人指定）"))
                    {
                        return true;
                    }
                }
                break;

            //case "セルアウトリベート登録申請（NEW）":
            //foreach (string strGroup in scGroupList)
            //{
            //if (strGroup.Contains("IT戦略部") || strGroup.Contains("経理部") || strGroup.Contains("取締役（個人指定）"))
            //{
            //return true;
            //}
            //}
            //break;

            case "器具発注申請":
                foreach (string strGroup in scGroupList)
                {
                    if (strGroup.Contains("購買課") || strGroup.Contains("経理部") || strGroup.Contains("取締役（個人指定）"))
                    {
                        return true;
                    }
                }
                break;
        }

        return false;
    }
}
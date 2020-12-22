using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace _003_신한_EMER
{
    public class CONVERT
    {
        //기본 인코딩 설정
        private static string strEncoding = "ks_c_5601-1987";
        private static string strCardTypeID = "061";
        private static string strCardTypeName = "신한긴급";

        //현 DLL의 카드 타입 코드 반환
        public static string GetCardTypeID()
        {
            return strCardTypeID;
        }

        //현 DLL의 카드 타입명 반환
        public static string GetCardTypeName()
        {
            return strCardTypeName;
        }

        //제휴사코드 반환
        public static string GetCardType(string path)
        {
            System.Text.Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            StreamReader _sr = null;
            string _strReturn = "";
            string _strLine = "";

            //파일 읽기 Stream과 오류시 저장할 쓰기 Stream 생성
            _sr = new StreamReader(path, _encoding);
            _strLine = _sr.ReadLine();
            try
            {
                if (_strLine.Trim() != "")
                {
                    _strReturn = _strLine.Substring(_strLine.Length - 7, 7);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }

            return _strReturn;
        }

        //등록 자료 생성
        //public static string ConvertRegister(string path, string xmlZipcodePath, string xmlZipcodeAreaPath, string xmlPath)
        //public static string ConvertRegister(string path, string xmlZipcodePath, string xmlZipcodeAreaPath, string xmlZipcodePath_new, string xmlZipcodeAreaPath_new, string xmlPath)
        //{
        //    System.Text.Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
        //    ////FileInfo _fi = null;
        //    StreamReader _sr = null;											//파일 읽기 스트림
        //    StreamWriter _swError = null;										//파일 쓰기 스트림
        //    DataSet _dsetZipcode = null, _dsetZipcdeArea = null;				//우편번호 관련 DataSet
        //    DataSet _dsetZipcode_new = null, _dsetZipcdeArea_new = null;		//우편번호 관련 DataSet
        //    DataTable _dtable = null;											//마스터 저장 테이블
        //    DataRow _dr = null;
        //    DataRow[] _drs = null;
        //    byte[] _byteAry = null;
        //    string _strReturn = "";
        //    string _strLine = "";
        //    string _strZipcode = "", _strAreaType = "", _strAreaGroup = "", _strBranch = "", strCard_type_detail = "", strClient_express_code = "";
        //    string _strDeliveryPlaceType = "", strMemo = "";
        //    int _iSeq = 1, _iErrorCount = 0;
        //    Branches _branches = new Branches();
        //    try
        //    {
        //        _dtable = new DataTable("Convert");
        //        //기본 컬럼
        //        _dtable.Columns.Add("degree_arrange_number");
        //        _dtable.Columns.Add("card_area_group");
        //        _dtable.Columns.Add("card_branch");
        //        _dtable.Columns.Add("card_area_type");
        //        _dtable.Columns.Add("area_arrange_number");
        //        //세부 컬럼
        //        _dtable.Columns.Add("customer_SSN");                //민증   dr[5]
        //        _dtable.Columns.Add("card_number");                 //카드번호
        //        _dtable.Columns.Add("card_brand_code");             //브랜드코드
        //        _dtable.Columns.Add("card_urgency_code");           //필터1 609등록구분 2012-03-13 "Q" 체크
        //        _dtable.Columns.Add("card_pt_code");                //필터2 법인구분
        //        _dtable.Columns.Add("customer_name");               //이름   [10]
        //        _dtable.Columns.Add("card_agree_code");             //특별동의서징구여부 p-파워콤, G-gs칼텍스
        //        _dtable.Columns.Add("client_insert_type");          //대리수령동의여부 코드
        //        _dtable.Columns.Add("card_delivery_place_code");    //수령지코드 1자택, 2직장
        //        _dtable.Columns.Add("card_bank_ID");                //관리지점
        //        _dtable.Columns.Add("card_mobile_tel");             //dr[15] 휴대폰
        //        _dtable.Columns.Add("card_zipcode");                //우편번호
        //        _dtable.Columns.Add("card_issue_type_code");        //발급사유(발급구분)
        //        _dtable.Columns.Add("card_count");                  //카드매수 
        //        _dtable.Columns.Add("card_design_code");            //카드형태 본인, 가족

        //        _dtable.Columns.Add("card_cooperation_code");       //dr[20] 제휴코드 
        //        _dtable.Columns.Add("client_register_date");        //작업일자
        //        _dtable.Columns.Add("client_number");               //작업seq
        //        _dtable.Columns.Add("client_quick_work_date");      //특송작업일자
        //        _dtable.Columns.Add("client_send_number");          //특송순번
        //        _dtable.Columns.Add("card_vip_code");               //dr[25] VIP구분   
        //        _dtable.Columns.Add("family_SSN");                  //가족식별번호
        //        _dtable.Columns.Add("family_name");                 //가족이름
        //        _dtable.Columns.Add("card_client_code_1");          //보이스 형태코드
        //        _dtable.Columns.Add("card_client_no_1");            //메모코드
        //        _dtable.Columns.Add("card_register_type");          //dr[30] 긴급배송여부 코드 "U"    
        //        _dtable.Columns.Add("card_cooperation1");           //BPRECC번호
        //        _dtable.Columns.Add("card_cooperation2");           //BPRECC엔코딩값
        //        _dtable.Columns.Add("card_product_code");           //동의서양식코드
        //        _dtable.Columns.Add("client_express_code");         //일반,긴급,동의 구분코드
        //        _dtable.Columns.Add("client_request_memo");         //dr[35] 메모

        //        //내부변환코드
        //        _dtable.Columns.Add("card_barcode_new");            //케리어바코드
        //        _dtable.Columns.Add("card_issue_type_new");         //발급구분코드_new
        //        _dtable.Columns.Add("card_delivery_place_type");    //dr[38] 내부수령지코드 1자택, 2직장

        //        _dtable.Columns.Add("card_zipcode_new");            //dr[39] 신우편번호
        //        _dtable.Columns.Add("card_zipcode_kind");           //dr[40] 신우편번호구분코드

        //        //파일 읽기 Stream과 오류시 저장할 쓰기 Stream 생성
        //        _sr = new StreamReader(path, _encoding);
        //        _swError = new StreamWriter(path + ".Error", false, _encoding);

        //        //우편번호 관련 정보 DataSet에 담기
        //        _dsetZipcode = new DataSet();
        //        _dsetZipcdeArea = new DataSet();
        //        _dsetZipcode.ReadXml(xmlZipcodePath);
        //        _dsetZipcode.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcode.Tables[0].Columns["zipcode"] };
        //        _dsetZipcdeArea.ReadXml(xmlZipcodeAreaPath);
        //        _dsetZipcdeArea.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcdeArea.Tables[0].Columns["zipcode"] };

        //        //신우편번호 관련 정보 DataSet에 담기
        //        _dsetZipcode_new = new DataSet();
        //        _dsetZipcdeArea_new = new DataSet();
        //        _dsetZipcode_new.ReadXml(xmlZipcodePath_new);
        //        _dsetZipcode_new.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcode_new.Tables[0].Columns["zipcode_new"] };
        //        _dsetZipcdeArea_new.ReadXml(xmlZipcodeAreaPath_new);
        //        _dsetZipcdeArea_new.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcdeArea_new.Tables[0].Columns["zipcode_new"] };

        //        while ((_strLine = _sr.ReadLine()) != null)
        //        {
        //            //인코딩, byte 배열로 담기
        //            if (_iSeq == 1)
        //            {
        //                strCard_type_detail = _strLine.Substring(_strLine.Length - 7, 7);
        //            }

        //            _byteAry = _encoding.GetBytes(_strLine);

        //            //_swError.WriteLine(_strLine);
        //            _strDeliveryPlaceType = _encoding.GetString(_byteAry, 48, 1);

        //            _dr = _dtable.NewRow();
        //            _dr[0] = _iSeq;

        //            _dr[5] = _encoding.GetString(_byteAry, 0, 13).Replace('X', 'x').Replace('*', 'x');
        //            _dr[6] = _encoding.GetString(_byteAry, 13, 12);
        //            _dr[7] = _encoding.GetString(_byteAry, 25, 1);
        //            _dr[8] = _encoding.GetString(_byteAry, 26, 1);
        //            _dr[9] = _encoding.GetString(_byteAry, 27, 1);
        //            _dr[10] = _encoding.GetString(_byteAry, 28, 18);
        //            _dr[11] = _encoding.GetString(_byteAry, 46, 1);
        //            _dr[12] = _encoding.GetString(_byteAry, 47, 1);
        //            _dr[13] = _strDeliveryPlaceType;
        //            _dr[14] = _encoding.GetString(_byteAry, 49, 5).Replace(" ", "").Trim();
        //            if (_strDeliveryPlaceType.Equals("1"))
        //            {
        //                _dr[15] = _encoding.GetString(_byteAry, 54, 16).Replace("-", "");
        //                _dr[16] = _encoding.GetString(_byteAry, 70, 16).Replace("-", "");
        //            }
        //            else
        //            {
        //                _dr[15] = _encoding.GetString(_byteAry, 70, 16).Replace("-", "");
        //                _dr[16] = _encoding.GetString(_byteAry, 54, 16).Replace("-", "");
        //            }
        //            _dr[17] = _encoding.GetString(_byteAry, 86, 12);

        //            _strZipcode = _encoding.GetString(_byteAry, 98, 8).Replace(" ", "").Trim();

        //            if (_strZipcode.Length == 5)
        //            {
        //                _dr[39] = _strZipcode;
        //                _dr[40] = "1";
        //            }
        //            _dr[18] = _encoding.GetString(_byteAry, 98, 8).Replace(" ", "");

        //            _dr[19] = _encoding.GetString(_byteAry, 106, 100);
        //            _dr[20] = _encoding.GetString(_byteAry, 206, 100);
        //            _dr[21] = _encoding.GetString(_byteAry, 306, 1);
        //            _dr[22] = _encoding.GetString(_byteAry, 307, 1);
        //            _dr[23] = _encoding.GetString(_byteAry, 308, 1);

        //            _dr[24] = _encoding.GetString(_byteAry, 309, 8).Replace(" ", "");
        //            _dr[25] = _encoding.GetString(_byteAry, 317, 100);
        //            _dr[26] = _encoding.GetString(_byteAry, 417, 100);
        //            _dr[27] = _encoding.GetString(_byteAry, 517, 26);
        //            _dr[28] = _encoding.GetString(_byteAry, 543, 6);
        //            _dr[29] = _encoding.GetString(_byteAry, 549, 8);
        //            _dr[30] = _encoding.GetString(_byteAry, 557, 5);
        //            _dr[31] = _encoding.GetString(_byteAry, 562, 8);
        //            _dr[32] = _encoding.GetString(_byteAry, 570, 6);
        //            _dr[33] = _encoding.GetString(_byteAry, 576, 1);
                    
        //            strClient_express_code = "2005";

        //            _dr[34] = strClient_express_code;

        //            switch (_dr[29].ToString().Trim())
        //            {
        //                case "1": strMemo = "오전 배송 요망";
        //                    break;
        //                case "2": strMemo = "오후 배송 요망";
        //                    break;
        //                case "3": strMemo = "연락후 배송 요망";
        //                    break;
        //                case "4": strMemo = "부재 시 연락 요망";
        //                    break;
        //                case "5": strMemo = "오후6시 이후 배송 요망";
        //                    break;
        //                case "6": strMemo = "배우자 배송 요망(일반 only)";
        //                    break;
        //                case "7": strMemo = "반드시 본인 전달 요망";
        //                    break;
        //                case "8": strMemo = "신분증 확인 철저 요망";
        //                    break;
        //                case "9": strMemo = "민원 주의(친절배송 요망)";
        //                    break;
        //                default:
        //                    strMemo = _dr[29].ToString().Trim();
        //                    break;
        //            }
        //            //2012-03-14 태희철 수정 PDA메모에 익일긴급 표시되게 등록 시 수정
        //            if (_encoding.GetString(_byteAry, 26, 1) == "Q")
        //            {
        //                _dr[35] = strMemo + " 익일긴급";
        //            }
        //            else
        //            {
        //                _dr[35] = strMemo;
        //            }

        //            if (_encoding.GetString(_byteAry, 44, 1) == "0" && strClient_express_code == "2002")
        //            {
        //                if (strMemo.Trim() == "")
        //                {
        //                    _dr[35] = "본인수령요청";
        //                }
        //                else
        //                {
        //                    _dr[35] = "본인수령요청 / " + strMemo;
        //                }
        //            }



        //            //2011-08-23 태희철 수정
        //            ///[내용] 신한의 경우 카드번호가 12자리와 11자리 두 분류로 되어있음
        //            ///11자리 일 경우 중간에 공백이 생겨 바코드 리딩이 안됨.
        //            ///공백을 제거하고 업데이트 하여 케리어 바코드를 22자리로 생성
        //            ///리딩 할 경우에도 공백을 제거 하므로 22자리로 인식 가능
        //            _dr[36] = "0" + strClient_express_code + _dr[6].ToString().Trim() + _strZipcode;

        //            //내부변환코드
        //            _dr[37] = _dr[17];
        //            _dr[38] = _dr[13];

        //            if (_strZipcode != "")
        //            {
        //                //지역 분류 선택
        //                if (_strZipcode.Trim().Length == 5)
        //                {
        //                    _drs = _dsetZipcdeArea_new.Tables[0].Select("zipcode_new = " + _strZipcode);
        //                }
        //                else
        //                {
        //                    _drs = _dsetZipcdeArea.Tables[0].Select("zipcode = " + _strZipcode);
        //                }

        //                //_drs = _dsetZipcdeArea.Tables[0].Select("zipcode=" + _strZipcode);

        //                if (_drs.Length < 1)
        //                {
        //                    _strAreaGroup = "012";
        //                    _strBranch = "012";
        //                }
        //                else
        //                {
        //                    _strAreaGroup = _drs[0][0].ToString();
        //                    _strBranch = _drs[0][1].ToString();
        //                }

        //                //우편번호 유효성 검사
        //                if (_strZipcode.Trim().Length == 5)
        //                {
        //                    _drs = _dsetZipcode_new.Tables[0].Select("zipcode_new = " + _strZipcode);
        //                }
        //                else
        //                {
        //                    _drs = _dsetZipcode.Tables[0].Select("zipcode = " + _strZipcode);
        //                }

        //                //_drs = _dsetZipcode.Tables[0].Select("zipcode=" + _strZipcode);

        //                if (_drs.Length > 0)
        //                {
        //                    _strAreaType = _drs[0]["area_type_code"].ToString();
        //                }
        //                else
        //                {
        //                    _strAreaType = "";
        //                }

        //                //우편번호를 찾지 못 했다면 Error 파일에 기록
        //                if (_strAreaType.Equals(""))
        //                {
        //                    _swError.WriteLine(_strLine);
        //                    _iErrorCount++;
        //                }

        //                _dr[1] = _strAreaGroup;
        //                _dr[2] = _strBranch;
        //                _dr[3] = _strAreaType;
        //                _dr[4] = _branches.GetCount(_strBranch);
        //                _dtable.Rows.Add(_dr);
        //            }
        //            else
        //            {
        //                _swError.WriteLine(_strLine);
        //                _iErrorCount++;
        //            }
        //            _iSeq++;
        //        }

        //        //변환에 성공했다면
        //        if (_iErrorCount < 1)
        //        {
        //            _swError.Close();
        //            _sr.Close();
        //            _dtable.WriteXml(xmlPath);
        //            //_fi = new FileInfo(path + ".Error");
        //            //_fi.Delete();
        //            _strReturn = string.Format("{0}건의 데이터 변환 성공", _iSeq - 1);
        //        }
        //        else
        //        {
        //            _strReturn = string.Format("{0}건 변환, 우편번호 미등록 {1}건 실패", _iSeq - 1, _iErrorCount);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        if (_swError != null)
        //        {
        //            _strReturn = string.Format("{0}번 째 변환 중 오류", _iSeq);
        //            _swError.WriteLine(_strLine);
        //            _swError.WriteLine(ex.Message);
        //        }
        //    }
        //    finally
        //    {
        //        if (_swError != null) _swError.Close();
        //        if (_sr != null) _sr.Close();
        //    }
        //    return _strReturn;
        //}

        public static string ConvertRegister(string path, string xmlZipcodePath, string xmlZipcodeAreaPath, string xmlZipcodePath_new, string xmlZipcodeAreaPath_new, string xmlPath)
        {
            System.Text.Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            ////FileInfo _fi = null;
            StreamReader _sr = null;											//파일 읽기 스트림
            StreamWriter _swError = null;										//파일 쓰기 스트림
            DataSet _dsetZipcode = null, _dsetZipcdeArea = null;				//우편번호 관련 DataSet
            DataSet _dsetZipcode_new = null, _dsetZipcdeArea_new = null;		//우편번호 관련 DataSet
            DataTable _dtable = null;											//마스터 저장 테이블
            DataRow _dr = null;
            DataRow[] _drs = null;
            byte[] _byteAry = null;
            string _strReturn = "";
            string _strLine = "";
            string _strZipcode = "", _strAreaType = "", _strAreaGroup = "", _strBranch = "", strCard_type_detail = "", strClient_express_code = "";
            string _strDeliveryPlaceType = "", strMemo = "", strCustomer_ssn = "";
            int _iSeq = 1, _iErrorCount = 0;
            Branches _branches = new Branches();
            try
            {
                _dtable = new DataTable("Convert");
                //기본 컬럼
                _dtable.Columns.Add("degree_arrange_number");
                _dtable.Columns.Add("card_area_group");
                _dtable.Columns.Add("card_branch");
                _dtable.Columns.Add("card_area_type");
                _dtable.Columns.Add("area_arrange_number");
                //세부 컬럼
                _dtable.Columns.Add("customer_SSN");                //민증   dr[5]
                _dtable.Columns.Add("card_number");                 //카드번호
                _dtable.Columns.Add("card_brand_code");             //브랜드코드
                _dtable.Columns.Add("card_urgency_code");           //필터1 609등록구분 2012-03-13 "Q" 체크
                _dtable.Columns.Add("card_pt_code");                //필터2 법인구분
                _dtable.Columns.Add("customer_name");               //이름   [10]
                _dtable.Columns.Add("card_agree_code");             //특별동의서징구여부 p-파워콤, G-gs칼텍스
                _dtable.Columns.Add("client_insert_type");          //대리수령동의여부 코드
                _dtable.Columns.Add("card_delivery_place_code");    //수령지코드 1자택, 2직장
                _dtable.Columns.Add("card_bank_ID");                //관리지점
                _dtable.Columns.Add("card_mobile_tel");             //dr[15] 휴대폰
                _dtable.Columns.Add("card_zipcode");                //우편번호
                _dtable.Columns.Add("card_issue_type_code");        //발급사유(발급구분)
                _dtable.Columns.Add("card_count");                  //카드매수 
                _dtable.Columns.Add("card_design_code");            //카드형태 본인, 가족

                _dtable.Columns.Add("card_cooperation_code");       //dr[20] 제휴코드 
                _dtable.Columns.Add("client_register_date");        //작업일자
                _dtable.Columns.Add("client_number");               //작업seq
                _dtable.Columns.Add("client_quick_work_date");      //특송작업일자
                _dtable.Columns.Add("client_send_number");          //특송순번
                _dtable.Columns.Add("card_vip_code");               //dr[25] VIP구분   
                _dtable.Columns.Add("family_SSN");                  //가족식별번호
                _dtable.Columns.Add("family_name");                 //가족이름
                _dtable.Columns.Add("card_client_code_1");          //보이스 형태코드
                _dtable.Columns.Add("card_client_no_1");            //메모코드
                _dtable.Columns.Add("card_register_type");          //dr[30] 긴급배송여부 코드 "U"    
                _dtable.Columns.Add("card_cooperation1");           //BPRECC번호
                _dtable.Columns.Add("card_cooperation2");           //BPRECC엔코딩값
                _dtable.Columns.Add("card_product_code");           //동의서양식코드
                _dtable.Columns.Add("client_express_code");         //일반,긴급,동의 구분코드
                _dtable.Columns.Add("client_request_memo");         //dr[35] 메모

                //내부변환코드
                _dtable.Columns.Add("card_barcode_new");            //케리어바코드
                _dtable.Columns.Add("card_issue_type_new");         //발급구분코드_new
                _dtable.Columns.Add("card_delivery_place_type");    //dr[38] 내부수령지코드 1자택, 2직장

                _dtable.Columns.Add("card_zipcode_new");            //dr[39] 신우편번호
                _dtable.Columns.Add("card_zipcode_kind");           //dr[40] 신우편번호구분코드

                _dtable.Columns.Add("choice_agree1");               //dr[41] 필수동의
                _dtable.Columns.Add("choice_agree3");               //dr[42] 선택동의

                _dtable.Columns.Add("customer_memo");              //dr[43] 팝업메모문구
                _dtable.Columns.Add("card_is_for_owner_only");     //dr[44] 본인만배송

                _dtable.Columns.Add("client_enterprise_code");     //dr[45] 동의서양식코드2
                _dtable.Columns.Add("card_level_code");             //dr[46] 카드상품번호2(제공문구코드)
                _dtable.Columns.Add("card_address");                //dr[47] 동이상주소
                _dtable.Columns.Add("card_bank_account_tel");       //dr[48]실번호 뒷4자리
                _dtable.Columns.Add("change_add");                  //dr[49]신분증여부

                //파일 읽기 Stream과 오류시 저장할 쓰기 Stream 생성
                _sr = new StreamReader(path, _encoding);
                _swError = new StreamWriter(path + ".Error", false, _encoding);

                //우편번호 관련 정보 DataSet에 담기
                _dsetZipcode = new DataSet();
                _dsetZipcdeArea = new DataSet();
                _dsetZipcode.ReadXml(xmlZipcodePath);
                _dsetZipcode.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcode.Tables[0].Columns["zipcode"] };
                _dsetZipcdeArea.ReadXml(xmlZipcodeAreaPath);
                _dsetZipcdeArea.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcdeArea.Tables[0].Columns["zipcode"] };

                //신우편번호 관련 정보 DataSet에 담기
                _dsetZipcode_new = new DataSet();
                _dsetZipcdeArea_new = new DataSet();
                _dsetZipcode_new.ReadXml(xmlZipcodePath_new);
                _dsetZipcode_new.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcode_new.Tables[0].Columns["zipcode_new"] };
                _dsetZipcdeArea_new.ReadXml(xmlZipcodeAreaPath_new);
                _dsetZipcdeArea_new.Tables[0].PrimaryKey = new DataColumn[] { _dsetZipcdeArea_new.Tables[0].Columns["zipcode_new"] };

                while ((_strLine = _sr.ReadLine()) != null)
                {
                    //인코딩, byte 배열로 담기
                    if (_iSeq == 1)
                    {
                        strCard_type_detail = _strLine.Substring(_strLine.Length - 7, 7);
                    }

                    _byteAry = _encoding.GetBytes(_strLine);

                    //_swError.WriteLine(_strLine);
                    _strDeliveryPlaceType = _encoding.GetString(_byteAry, 45, 1);

                    _dr = _dtable.NewRow();
                    _dr[0] = _iSeq;

                    strCustomer_ssn = _encoding.GetString(_byteAry, 0, 10).Replace('X', 'x').Replace('*', 'x');
                    if (strCustomer_ssn.Length != 0)
                    {
                        strCustomer_ssn = strCustomer_ssn + "xxx";
                    }

                    _dr[5] = strCustomer_ssn;
                    _dr[6] = _encoding.GetString(_byteAry, 10, 12);
                    _dr[7] = _encoding.GetString(_byteAry, 22, 1);
                    _dr[8] = _encoding.GetString(_byteAry, 23, 1);
                    _dr[9] = _encoding.GetString(_byteAry, 24, 1);
                    _dr[10] = _encoding.GetString(_byteAry, 25, 18);
                    _dr[11] = _encoding.GetString(_byteAry, 43, 1);
                    _dr[12] = _encoding.GetString(_byteAry, 44, 1);
                    _dr[13] = _strDeliveryPlaceType;
                    _dr[14] = _encoding.GetString(_byteAry, 46, 4).Replace(" ", "").Trim();
                    _dr[15] = _encoding.GetString(_byteAry, 50, 12);

                    _strZipcode = _encoding.GetString(_byteAry, 62, 8).Replace(" ", "").Trim();
                    _dr[16] = _strZipcode;

                    if (_strZipcode.Trim().Length == 5)
                    {
                        _dr[39] = _strZipcode;
                        _dr[40] = "1";
                    }

                    _dr[17] = _encoding.GetString(_byteAry, 70, 1);
                    _dr[18] = _encoding.GetString(_byteAry, 71, 1);
                    _dr[19] = _encoding.GetString(_byteAry, 72, 1);

                    _dr[20] = _encoding.GetString(_byteAry, 73, 6);
                    _dr[21] = _encoding.GetString(_byteAry, 79, 8);
                    _dr[22] = _encoding.GetString(_byteAry, 87, 5);
                    _dr[23] = _encoding.GetString(_byteAry, 92, 8);
                    _dr[24] = _encoding.GetString(_byteAry, 100, 6);
                    _dr[25] = _encoding.GetString(_byteAry, 107, 1);
                    _dr[26] = _encoding.GetString(_byteAry, 108, 13);
                    _dr[27] = _encoding.GetString(_byteAry, 121, 12);
                    _dr[28] = _encoding.GetString(_byteAry, 133, 1);
                    _dr[29] = _encoding.GetString(_byteAry, 134, 1);

                    _dr[30] = _encoding.GetString(_byteAry, 135, 1);
                    _dr[31] = _encoding.GetString(_byteAry, 136, 25);
                    _dr[32] = _encoding.GetString(_byteAry, 161, 50);
                    _dr[33] = _encoding.GetString(_byteAry, 211, 4);

                    //중요
                    strClient_express_code = "2005";

                    _dr[34] = strClient_express_code;

                    switch (_dr[29].ToString().Trim())
                    {
                        case "1": strMemo = "오전 배송 요망";
                            break;
                        case "2": strMemo = "오후 배송 요망";
                            break;
                        case "3": strMemo = "연락후 배송 요망";
                            break;
                        case "4": strMemo = "부재 시 연락 요망";
                            break;
                        case "5": strMemo = "오후6시 이후 배송 요망";
                            break;
                        case "6": strMemo = "배우자 배송 요망(일반 only)";
                            break;
                        case "7": strMemo = "반드시 본인 전달 요망";
                            break;
                        case "8": strMemo = "신분증 확인 철저 요망";
                            break;
                        case "9": strMemo = "민원 주의(친절배송 요망)";
                            break;
                        default:
                            strMemo = _dr[29].ToString().Trim();
                            break;
                    }
                    //2012-03-14 태희철 수정 PDA메모에 익일긴급 표시되게 등록 시 수정
                    if (_encoding.GetString(_byteAry, 26, 1) == "Q")
                    {
                        _dr[35] = strMemo + " 익일긴급";
                    }
                    else
                    {
                        _dr[35] = strMemo;
                    }

                    //동의서 외 일반, 긴급 포함
                    if (_encoding.GetString(_byteAry, 44, 1) == "0" && strClient_express_code != "2120")
                    {
                        _dr[35] = "본인수령요청";
                    }
                    else if (_encoding.GetString(_byteAry, 44, 1) == "2" && strClient_express_code != "2120")
                    {
                        _dr[44] = "1";
                        _dr[49] = "1";
                        if (strMemo == "")
                        {
                            _dr[43] = "본인지정배송";
                        }
                        else
                        {
                            _dr[43] = "본인지정배송 " + strMemo;
                        }

                    }



                    //2011-08-23 태희철 수정
                    ///[내용] 신한의 경우 카드번호가 12자리와 11자리 두 분류로 되어있음
                    ///11자리 일 경우 중간에 공백이 생겨 바코드 리딩이 안됨.
                    ///공백을 제거하고 업데이트 하여 케리어 바코드를 22자리로 생성
                    ///리딩 할 경우에도 공백을 제거 하므로 22자리로 인식 가능
                    _dr[36] = "0" + strClient_express_code + _dr[6].ToString().Trim() + _strZipcode;

                    //내부변환코드
                    _dr[37] = _dr[17];
                    _dr[38] = _dr[13];

                    if (_encoding.GetString(_byteAry, 215, 1) == "1" || _encoding.GetString(_byteAry, 215, 1) == "0")
                    {
                        _dr[41] = _encoding.GetString(_byteAry, 215, 1);
                    }
                    else
                    {
                        _dr[41] = "9";
                    }
                    //_dr[41] = _encoding.GetString(_byteAry, 215, 1);
                    _dr[42] = _encoding.GetString(_byteAry, 216, 10);

                    //_encoding.GetString(_byteAry, 226, 1);
                    _dr[45] = _encoding.GetString(_byteAry, 227, 4);
                    _dr[46] = _encoding.GetString(_byteAry, 231, 6);
                    //_encoding.GetString(_byteAry, 237, 59);
                    //동이상주소 2018.04.30
                    _dr[47] = _encoding.GetString(_byteAry, 296, 100);
                    //전화번호 4자리 2018.05.28
                    _dr[48] = _encoding.GetString(_byteAry, 396, 4);

                    if (_strZipcode != "")
                    {
                        //지역 분류 선택
                        if (_strZipcode.Trim().Length == 5)
                        {
                            _drs = _dsetZipcdeArea_new.Tables[0].Select("zipcode_new = " + _strZipcode);
                        }
                        else
                        {
                            _drs = _dsetZipcdeArea.Tables[0].Select("zipcode = " + _strZipcode);
                        }

                        //_drs = _dsetZipcdeArea.Tables[0].Select("zipcode=" + _strZipcode);

                        if (_drs.Length < 1)
                        {
                            _strAreaGroup = "012";
                            _strBranch = "012";
                        }
                        else
                        {
                            _strAreaGroup = _drs[0][0].ToString();
                            _strBranch = _drs[0][1].ToString();
                        }

                        //우편번호 유효성 검사
                        if (_strZipcode.Trim().Length == 5)
                        {
                            _drs = _dsetZipcode_new.Tables[0].Select("zipcode_new = " + _strZipcode);
                        }
                        else
                        {
                            _drs = _dsetZipcode.Tables[0].Select("zipcode = " + _strZipcode);
                        }

                        //_drs = _dsetZipcode.Tables[0].Select("zipcode=" + _strZipcode);

                        if (_drs.Length > 0)
                        {
                            _strAreaType = _drs[0]["area_type_code"].ToString();
                        }
                        else
                        {
                            _strAreaType = "";
                        }

                        //우편번호를 찾지 못 했다면 Error 파일에 기록
                        if (_strAreaType.Equals(""))
                        {
                            _swError.WriteLine(_strLine);
                            _iErrorCount++;
                        }

                        _dr[1] = _strAreaGroup;
                        _dr[2] = _strBranch;
                        _dr[3] = _strAreaType;
                        _dr[4] = _branches.GetCount(_strBranch);
                        _dtable.Rows.Add(_dr);
                    }
                    else
                    {
                        _swError.WriteLine(_strLine);
                        _iErrorCount++;
                    }
                    _iSeq++;
                }

                //변환에 성공했다면
                if (_iErrorCount < 1)
                {
                    _swError.Close();
                    _sr.Close();
                    _dtable.WriteXml(xmlPath);
                    //_fi = new FileInfo(path + ".Error");
                    //_fi.Delete();
                    _strReturn = string.Format("{0}건의 데이터 변환 성공", _iSeq - 1);
                }
                else
                {
                    _strReturn = string.Format("{0}건 변환, 우편번호 미등록 {1}건 실패", _iSeq - 1, _iErrorCount);
                }
            }
            catch (Exception ex)
            {
                if (_swError != null)
                {
                    _strReturn = string.Format("{0}번 째 변환 중 오류", _iSeq);
                    _swError.WriteLine(_strLine);
                    _swError.WriteLine(ex.Message);
                }
            }
            finally
            {
                if (_swError != null) _swError.Close();
                if (_sr != null) _sr.Close();
            }
            return _strReturn;
        }


        //마감 자료
        public static string ConvertResult(System.Data.DataTable dtable, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            StreamWriter _sw01 = null, _sw00 = null, _sw11 = null, _sw10 = null;				//파일 쓰기 스트림
            int i = 0, iCnt = 0, iRe_cnt = 0;
            StringBuilder _strLine = new StringBuilder("");
            string _strReturn = "", strStatus = "", strDegree = "", _strReturnCode = "";
            DataRow[] _drs = null;
            try
            {
                _sw01 = new StreamWriter(fileName + strDegree + ".OLD_01", true, _encoding);
                _sw00 = new StreamWriter(fileName + strDegree + ".OLD_00", true, _encoding);
                _sw10 = new StreamWriter(fileName + ".00", true, _encoding);
                _sw11 = new StreamWriter(fileName + ".01", true, _encoding);

                _drs = dtable.Select("", "delivery_result_editdate");

                //헤더 부분
                string temp_head = GetStringAsLength("HDKJ" + DateTime.Now.ToString("yyyyMMdd"), 12, false, ' ');
                _sw11.WriteLine(GetStringAsLength(temp_head, 300, true, ' '));

                
                for (i = 0; i < _drs.Length; i++)
                {
                    strStatus = _drs[i]["card_delivery_status"].ToString();
                    strDegree = _drs[i]["degree_code"].ToString();
                    _strReturnCode = _drs[i]["delivery_return_reason_last"].ToString();

                    _strLine = new StringBuilder(GetStringAsLength(_drs[i]["card_number"].ToString(), 12, true, ' '));
                    _strLine.Append(GetStringAsLength(_drs[i]["card_brand_code"].ToString(), 1, true, ' '));
                    _strLine.Append(GetStringAsLength("", 3, true, ' '));

                    #region 배송      
                    if (strStatus == "1")
                    {
                        if (_drs[i]["receiver_code_change"].ToString() == "001" || _drs[i]["receiver_code"].ToString() == "01")
                        {
                            _strLine.Append(GetStringAsLength("Y1", 2, true, ' '));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength("Y2", 2, true, ' '));
                        }

                        _strLine.Append(GetStringAsLength(RemoveDash(_drs[i]["card_delivery_date"].ToString()), 8, true, ' '));
                        //if (_drs[i]["receiver_SSN"].ToString().Length == 0)
                        //{
                        //    _strLine.Append(GetStringAsLength("", 6, true, '0'));
                        //    _strLine.Append(GetStringAsLength("", 8, true, ' '));
                        //}
                        //else
                        //{   
                        //    _strLine.Append(GetStringAsLength(ConvertLGSSN(_drs[i]["receiver_SSN"].ToString().Replace("x", "0")), 14, true, ' '));
                        //}

                        //민증번호
                        if (_drs[i]["card_result_status"].ToString() == "61")
                        {
                            _strLine.Append(GetStringAsLength(_drs[i]["customer_ssn"].ToString().Substring(7, 3), 14, true, ' '));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength(ConvertLGSSN(_drs[i]["receiver_SSN"].ToString().Replace("x", "0")), 14, true, ' '));
                        }

                        _strLine.Append(GetStringAsLength(_drs[i]["receiver_tel"].ToString(), 15, true, ' '));
                        if (_drs[i]["card_issue_type"].ToString() == "5")
                        {
                            _strLine.Append(GetStringAsLength(RemoveDash(_drs[i]["client_send_date"].ToString()), 8, true, ' '));
                            _strLine.Append(GetStringAsLength("", 5, false, ' '));
                            _strLine.Append(GetStringAsLength(RemoveDash(_drs[i]["card_in_date"].ToString()), 8, true, ' '));
                            _strLine.Append(GetStringAsLength("", 6, true, ' '));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength(RemoveDash(_drs[i]["client_register_date"].ToString()), 8, true, ' '));
                            _strLine.Append(GetStringAsLength(_drs[i]["client_number"].ToString(), 5, false, '0'));
                            _strLine.Append(GetStringAsLength(RemoveDash(_drs[i]["client_quick_work_date"].ToString()), 8, true, ' '));
                            _strLine.Append(GetStringAsLength(_drs[i]["client_send_number"].ToString(), 6, true, ' '));
                        }
                        _strLine.Append(GetStringAsLength("", 1, true, ' '));
                        _strLine.Append(GetStringAsLength(_drs[i]["receiver_name"].ToString(), 19, true, ' '));
                        _strLine.Append(GetStringAsLength(_drs[i]["receiver_code_change"].ToString(), 3, false, ' '));
                        _strLine.Append(GetStringAsLength(" ", 1, true, ' '));
                    }
                    #endregion

                    #region 반송     
                    else if (strStatus == "2" || strStatus == "3")
                    {
                        _strLine.Append(GetStringAsLength(ReturnType(_strReturnCode), 2, true, ' '));
                        _strLine.Append(GetStringAsLength(RemoveDash(_drs[i]["delivery_return_date_last"].ToString()), 8, true, ' '));
                        _strLine.Append(GetStringAsLength("", 14, true, ' '));
                        _strLine.Append(GetStringAsLength("", 15, true, ' '));

                        if (_drs[i]["client_register_date"].ToString() == "")
                            _strLine.Append(GetStringAsLength(_drs[i]["client_send_date"].ToString().Replace("-", ""), 8, true, ' '));  //제작일자
                        else
                            _strLine.Append(GetStringAsLength(_drs[i]["client_register_date"].ToString().Replace("-", ""), 8, true, ' '));  //제작일자

                        _strLine.Append(GetStringAsLength(_drs[i]["client_number"].ToString(), 5, true, ' '));                          //제작순번

                        if (_drs[i]["client_quick_work_date"].ToString() == "")
                            _strLine.Append(GetStringAsLength(_drs[i]["card_in_date"].ToString().Replace("-", ""), 8, true, ' '));//특송접수일자
                        else
                            _strLine.Append(GetStringAsLength(_drs[i]["client_quick_work_date"].ToString().Replace("-", ""), 8, true, ' '));//특송접수일자

                        _strLine.Append(GetStringAsLength(_drs[i]["client_send_number"].ToString(), 6, true, ' '));                     //특송접수번호

                        _strLine.Append(GetStringAsLength("", 20, true, ' '));
                        _strLine.Append(GetStringAsLength("", 3, false, ' '));
                        _strLine.Append(GetStringAsLength(" ", 1, true, ' '));
                    }
                    #endregion

                    //결번은 DB에서 필터됨

                    #region 분실     
                    else if (strStatus == "6")
                    {
                        _strLine.Append(GetStringAsLength("LL", 2, true, ' '));
                        _strLine.Append(GetStringAsLength(RemoveDash(_drs[i]["delivery_result_regdate"].ToString()), 8, true, ' '));
                        _strLine.Append(GetStringAsLength("", 14, true, ' '));
                        _strLine.Append(GetStringAsLength("", 15, true, ' '));

                        if (_drs[i]["client_register_date"].ToString() == "")
                            _strLine.Append(GetStringAsLength(_drs[i]["client_send_date"].ToString().Replace("-", ""), 8, true, ' '));  //제작일자
                        else
                            _strLine.Append(GetStringAsLength(_drs[i]["client_register_date"].ToString().Replace("-", ""), 8, true, ' '));  //제작일자

                        _strLine.Append(GetStringAsLength(_drs[i]["client_number"].ToString(), 5, true, ' '));                          //제작순번

                        if (_drs[i]["client_quick_work_date"].ToString() == "")
                            _strLine.Append(GetStringAsLength(_drs[i]["card_in_date"].ToString().Replace("-", ""), 8, true, ' '));//특송접수일자
                        else
                            _strLine.Append(GetStringAsLength(_drs[i]["client_quick_work_date"].ToString().Replace("-", ""), 8, true, ' '));//특송접수일자

                        _strLine.Append(GetStringAsLength(_drs[i]["client_send_number"].ToString(), 6, true, ' '));                     //특송접수번호

                        _strLine.Append(GetStringAsLength("", 20, true, ' '));
                        _strLine.Append(GetStringAsLength("", 3, false, ' '));
                        _strLine.Append(GetStringAsLength(" ", 1, true, ' '));
                    }
                    #endregion

                    #region 재방     
                    else if (strStatus == "7")
                    {
                        _strLine.Append(GetStringAsLength("JB", 2, true, ' '));
                        _strLine.Append(GetStringAsLength("", 8, true, ' '));
                        _strLine.Append(GetStringAsLength("", 14, true, ' '));
                        _strLine.Append(GetStringAsLength("", 15, true, ' '));

                        if (_drs[i]["client_register_date"].ToString() == "")
                            _strLine.Append(GetStringAsLength(_drs[i]["client_send_date"].ToString().Replace("-", ""), 8, true, ' '));  //제작일자
                        else
                            _strLine.Append(GetStringAsLength(_drs[i]["client_register_date"].ToString().Replace("-", ""), 8, true, ' '));  //제작일자

                        _strLine.Append(GetStringAsLength(_drs[i]["client_number"].ToString(), 5, true, ' '));                          //제작순번

                        if (_drs[i]["client_quick_work_date"].ToString() == "")
                            _strLine.Append(GetStringAsLength(_drs[i]["card_in_date"].ToString().Replace("-", ""), 8, true, ' '));//특송접수일자
                        else
                            _strLine.Append(GetStringAsLength(_drs[i]["client_quick_work_date"].ToString().Replace("-", ""), 8, true, ' '));//특송접수일자

                        _strLine.Append(GetStringAsLength(_drs[i]["client_send_number"].ToString(), 6, true, ' '));                     //특송접수번호

                        _strLine.Append(GetStringAsLength("", 20, true, ' '));
                        _strLine.Append(GetStringAsLength("030", 3, false, ' '));
                        _strLine.Append(GetStringAsLength(" ", 1, true, ' '));
                    }
                    #endregion

                    else 
                    {
                        _strLine.Append(GetStringAsLength("", 2, true, ' '));
                        _strLine.Append(GetStringAsLength("", 8, true, ' '));
                        _strLine.Append(GetStringAsLength("", 14, true, ' '));
                        _strLine.Append(GetStringAsLength("", 15, true, ' '));

                        if (_drs[i]["client_register_date"].ToString() == "")
                            _strLine.Append(GetStringAsLength(_drs[i]["client_send_date"].ToString().Replace("-", ""), 8, true, ' '));  //제작일자
                        else
                            _strLine.Append(GetStringAsLength(_drs[i]["client_register_date"].ToString().Replace("-", ""), 8, true, ' '));  //제작일자

                        _strLine.Append(GetStringAsLength(_drs[i]["client_number"].ToString(), 5, true, ' '));                          //제작순번

                        if (_drs[i]["client_quick_work_date"].ToString() == "")
                            _strLine.Append(GetStringAsLength(_drs[i]["card_in_date"].ToString().Replace("-", ""), 8, true, ' '));//특송접수일자
                        else
                            _strLine.Append(GetStringAsLength(_drs[i]["client_quick_work_date"].ToString().Replace("-", ""), 8, true, ' '));//특송접수일자

                        _strLine.Append(GetStringAsLength(_drs[i]["client_send_number"].ToString(), 6, true, ' '));                     //특송접수번호

                        _strLine.Append(GetStringAsLength("", 20, true, ' '));
                        _strLine.Append(GetStringAsLength("", 3, false, ' '));
                        _strLine.Append(GetStringAsLength(" ", 1, true, ' '));

                    }

                    if (strStatus == "1")
                    {   
                        _strLine.Append(GetStringAsLength("Y", 1, true, ' '));
                        _strLine.Append(GetStringAsLength("Y", 1, true, ' '));
                        _strLine.Append(GetStringAsLength("Y", 1, true, ' '));
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength("", 1, true, ' '));
                        _strLine.Append(GetStringAsLength("", 1, true, ' '));
                        _strLine.Append(GetStringAsLength("", 1, true, ' '));
                    }


                    _strLine.Append(GetStringAsLength("", 2, true, ' '));
                    _strLine.Append(GetStringAsLength("", 7, true, ' '));

                    _strLine.Append(GetStringAsLength(_drs[i]["code"].ToString(), 1, true, ' '));                    //신분증코드
                    _strLine.Append(GetStringAsLength(RemoveDash(_drs[i]["date"].ToString()), 9, true, ' '));                    //발급일자
                    _strLine.Append(GetStringAsLength(_drs[i]["org"].ToString(), 30, true, ' '));                    //발급기관


                    if (strStatus == "0" || strStatus == "7")
                    {   
                        _sw00.WriteLine(_strLine.ToString());
                    }
                    else
                    {   
                        _sw01.WriteLine(_strLine.ToString());
                    }
                    //2013.07.25 태희철 구마감 끝[S]


                    //2013.07.25 태희철 신마감 시작[S]
                    _strLine = new StringBuilder("DT");                                          //시작부
                    _strLine.Append(GetStringAsLength(_drs[i]["card_number"].ToString().Replace("-", ""), 16, true, ' '));

                    if (strStatus == "1") //배송
                    {
                        if ((_drs[i]["receiver_code_change"].ToString() == "001") ||
                            //2012.07.06 태희철 수정
                            //(_drs[i]["card_result_status"].ToString() == "61"))
                            (_drs[i]["receiver_code"].ToString() == "01"))
                        {
                            _strLine.Append(GetStringAsLength("Y1", 2, true, ' '));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength("Y2", 2, true, ' '));
                        }
                    }
                    else if (strStatus == "7") //재방
                        _strLine.Append(GetStringAsLength("JB", 2, true, ' '));
                    else if (strStatus == "" || strStatus == "4" || strStatus == "6") //결번
                        _strLine.Append(GetStringAsLength("LL", 2, true, ' '));
                    else if (strStatus == "2" || strStatus == "3") //반송
                        _strLine.Append(GetStringAsLength(ReturnType(_strReturnCode), 2, true, ' '));
                    else //기타
                        _strLine.Append(GetStringAsLength("", 2, true, ' '));

                    //전달일자
                    if (strStatus == "1") //배송
                        _strLine.Append(GetStringAsLength(_drs[i]["card_delivery_date"].ToString().Replace("-", ""), 8, true, ' '));
                    else if (strStatus == "2" || strStatus == "3") //반송
                        _strLine.Append(GetStringAsLength(_drs[i]["delivery_return_date_last"].ToString().Replace("-", ""), 8, true, ' '));
                    else if (strStatus == "6")
                        _strLine.Append(GetStringAsLength(_drs[i]["delivery_result_regdate"].ToString().Replace("-", ""), 8, true, ' '));
                    else
                        _strLine.Append(GetStringAsLength("", 8, true, ' '));
                    //민증번호
                    if (_drs[i]["card_result_status"].ToString() == "61")
                    {
                        _strLine.Append(GetStringAsLength(_drs[i]["customer_ssn"].ToString().Substring(7, 3), 14, true, ' '));
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength(ConvertLGSSN(_drs[i]["receiver_SSN"].ToString().Replace("x", "0")), 14, true, ' '));
                    }
                    //_strLine.Append(GetStringAsLength(_drs[i]["receiver_SSN"].ToString().Replace("-", "").Replace("x", ""), 14, true, ' '));
                    //전화번호
                    _strLine.Append(GetStringAsLength(_drs[i]["receiver_tel"].ToString().Replace("-", ""), 15, true, ' '));         

                    //제작일자
                    if (_drs[i]["client_register_date"].ToString() == "")
                    {
                        _strLine.Append(GetStringAsLength(_drs[i]["client_send_date"].ToString().Replace("-", ""), 8, true, ' '));  
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength(_drs[i]["client_register_date"].ToString().Replace("-", ""), 8, true, ' '));
                    }

                    _strLine.Append(GetStringAsLength(_drs[i]["client_number"].ToString(), 5, true, ' '));                          //제작순번
                    //특송접수일자
                    if (_drs[i]["client_quick_work_date"].ToString() == "")
                        _strLine.Append(GetStringAsLength(_drs[i]["card_in_date"].ToString().Replace("-", ""), 8, true, ' '));
                    else
                        _strLine.Append(GetStringAsLength(_drs[i]["client_quick_work_date"].ToString().Replace("-", ""), 8, true, ' '));

                    _strLine.Append(GetStringAsLength(_drs[i]["client_send_number"].ToString(), 6, true, ' '));                     //특송접수번호
                    //수령인성명
                    if (strStatus == "1" || _drs[i]["card_issue_type"].ToString() == "5")
                        _strLine.Append(GetStringAsLength(_drs[i]["receiver_name"].ToString(), 40, true, ' '));
                    else
                        _strLine.Append(GetStringAsLength("", 40, true, ' '));
                    //관계코드 - 은행 요청 코드
                    if (strStatus == "1") 
                        _strLine.Append(GetStringAsLength(_drs[i]["receiver_code_change"].ToString(), 3, true, ' '));                   
                    else
                        _strLine.Append(GetStringAsLength("", 3, true, ' '));

                    _strLine.Append(GetStringAsLength("", 1, true, ' '));                                                           //예비

                    if (strStatus == "1")
                    {
                        _strLine.Append(GetStringAsLength(ConvertAgree(_drs[i]["card_agree1"].ToString()), 1, true, ' '));
                        _strLine.Append(GetStringAsLength(ConvertAgree(_drs[i]["card_agree2"].ToString()), 1, true, ' '));
                        _strLine.Append(GetStringAsLength(ConvertAgree(_drs[i]["card_agree3"].ToString()), 1, true, ' '));
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength("", 1, true, ' '));
                        _strLine.Append(GetStringAsLength("", 1, true, ' '));
                        _strLine.Append(GetStringAsLength("", 1, true, ' '));
                    }
                    //특송발송카드 BIN구분코드
                    _strLine.Append(GetStringAsLength(_drs[i]["card_client_no_1"].ToString(), 2, true, ' '));
                    //제휴사코드
                    _strLine.Append(GetStringAsLength(_drs[i]["client_express_code"].ToString(), 4, true, ' '));
                    _strLine.Append(GetStringAsLength(_drs[i]["code"].ToString(), 1, true, ' '));                    //신분증코드
                    _strLine.Append(GetStringAsLength(RemoveDash(_drs[i]["date"].ToString()), 9, true, ' '));                    //발급일자
                    _strLine.Append(GetStringAsLength(_drs[i]["org"].ToString(), 30, true, ' '));                    //발급기관

                    _strLine.Append(GetStringAsLength("", 123, true, ' '));                                                     //예비

                    if (strStatus == "0" || strStatus == "7")
                    {
                        iRe_cnt++;
                        _sw10.WriteLine(_strLine.ToString());
                    }
                    else
                    {
                        iCnt++;
                        _sw11.WriteLine(_strLine.ToString());
                    }
                }

                //2013.07.22 태희철 수정 [S] 신마감사용
                _strLine = new StringBuilder(GetStringAsLength("TR" + GetStringAsLength(iCnt.ToString(), 11, false, '0'), 300, true, ' '));
                _sw11.WriteLine(_strLine.ToString());
                //2013.07.22 태희철 수정 [E] 신마감사용

                _strReturn = string.Format("총 {0}건 / 결과완료 {1}건 / 미배송 {2}건 다운 완료", i, iCnt, iRe_cnt);
            }
            catch (Exception)
            {
                _strReturn = string.Format("{0}번째 데이터 인계 중 오류 발생", i + 1);
            }
            finally
            {
                if (_sw00 != null) _sw00.Close();
                if (_sw01 != null) _sw01.Close();
                if (_sw10 != null) _sw10.Close();
                if (_sw11 != null) _sw11.Close();
            }
            return _strReturn;
        }

        //일일마감자료
        public static string ConvertResultDay(System.Data.DataTable dtable, string fileName)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            StreamWriter _sw01 = null, _sw11 = null;				//파일 쓰기 스트림
            int i = 0, iCnt = 0, iRe_cnt = 0;
            StringBuilder _strLine = new StringBuilder("");
            string _strReturn = "", strStatus = "";
            DataRow[] _drs = null;
            try
            {   
                _sw01 = new StreamWriter(fileName + ".OLD", true, _encoding);
                _sw11 = new StreamWriter(fileName + ".NEW", true, _encoding);

                _drs = dtable.Select("", "delivery_result_editdate");

                //헤더 부분
                string temp_head = GetStringAsLength("HDKJ" + DateTime.Now.ToString("yyyyMMdd"), 12, false, ' ');

                _sw11.WriteLine(GetStringAsLength(temp_head, 300, true, ' '));

                
                for (i = 0; i < _drs.Length; i++)
                {
                    strStatus = _drs[i]["card_delivery_status"].ToString();

                    #region 배송      
                    if (strStatus == "1")
                    {
                        _strLine = new StringBuilder(GetStringAsLength(_drs[i]["card_number"].ToString(), 12, true, ' '));
                        _strLine.Append(GetStringAsLength(_drs[i]["card_brand_code"].ToString(), 1, true, ' '));
                        _strLine.Append(GetStringAsLength("", 3, true, ' '));

                        if (_drs[i]["receiver_code_change"].ToString() == "001" || _drs[i]["receiver_code"].ToString() == "01")
                        {
                            _strLine.Append(GetStringAsLength("Y1", 2, true, ' '));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength("Y2", 2, true, ' '));
                        }

                        _strLine.Append(GetStringAsLength(RemoveDash(_drs[i]["card_delivery_date"].ToString()), 8, true, ' '));
                        //if (_drs[i]["receiver_SSN"].ToString().Length == 0)
                        //{
                        //    _strLine.Append(GetStringAsLength("", 6, true, '0'));
                        //    _strLine.Append(GetStringAsLength("", 8, true, ' '));
                        //}
                        //else
                        //{   
                        //    _strLine.Append(GetStringAsLength(ConvertLGSSN(_drs[i]["receiver_SSN"].ToString().Replace("x", "0")), 14, true, ' '));
                        //}
                        //민증번호
                        if (_drs[i]["card_result_status"].ToString() == "61")
                        {
                            _strLine.Append(GetStringAsLength(_drs[i]["customer_ssn"].ToString().Substring(7, 3), 14, true, ' '));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength(ConvertLGSSN(_drs[i]["receiver_SSN"].ToString().Replace("x", "0")), 14, true, ' '));
                        }
                        _strLine.Append(GetStringAsLength(_drs[i]["receiver_tel"].ToString(), 15, true, ' '));
                        if (_drs[i]["card_issue_type"].ToString() == "5")
                        {
                            _strLine.Append(GetStringAsLength(RemoveDash(_drs[i]["client_send_date"].ToString()), 8, true, ' '));
                            _strLine.Append(GetStringAsLength("", 5, false, ' '));
                            _strLine.Append(GetStringAsLength(RemoveDash(_drs[i]["card_in_date"].ToString()), 8, true, ' '));
                            _strLine.Append(GetStringAsLength("", 6, true, ' '));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength(RemoveDash(_drs[i]["client_register_date"].ToString()), 8, true, ' '));
                            _strLine.Append(GetStringAsLength(_drs[i]["client_number"].ToString(), 5, false, '0'));
                            _strLine.Append(GetStringAsLength(RemoveDash(_drs[i]["client_quick_work_date"].ToString()), 8, true, ' '));
                            _strLine.Append(GetStringAsLength(_drs[i]["client_send_number"].ToString(), 6, true, ' '));
                        }
                        _strLine.Append(GetStringAsLength("", 1, true, ' '));
                        _strLine.Append(GetStringAsLength(_drs[i]["receiver_name"].ToString(), 19, true, ' '));
                        _strLine.Append(GetStringAsLength(_drs[i]["receiver_code_change"].ToString(), 3, false, ' '));
                        _strLine.Append(GetStringAsLength(" ", 1, true, ' '));

                        _strLine.Append(GetStringAsLength("Y", 1, true, ' '));
                        _strLine.Append(GetStringAsLength("Y", 1, true, ' '));
                        _strLine.Append(GetStringAsLength("Y", 1, true, ' '));


                        _strLine.Append(GetStringAsLength("", 2, true, ' '));
                        _strLine.Append(GetStringAsLength("", 7, true, ' '));

                        _strLine.Append(GetStringAsLength(_drs[i]["code"].ToString(), 1, true, ' '));                    //신분증코드
                        _strLine.Append(GetStringAsLength(RemoveDash(_drs[i]["date"].ToString()), 9, true, ' '));                    //발급일자
                        _strLine.Append(GetStringAsLength(_drs[i]["org"].ToString(), 30, true, ' '));                    //발급기관

                        _sw01.WriteLine(_strLine.ToString());
                    }
                    #endregion
                    //2013.07.25 태희철 구마감 끝[S]

                    if (strStatus == "1") //배송
                    {
                        //2013.07.25 태희철 신마감 시작[S]
                        _strLine = new StringBuilder("DT");                                          //시작부
                        _strLine.Append(GetStringAsLength(_drs[i]["card_number"].ToString().Replace("-", ""), 16, true, ' '));

                        if ((_drs[i]["receiver_code_change"].ToString() == "001") ||
                            //2012.07.06 태희철 수정
                            //(_drs[i]["card_result_status"].ToString() == "61"))
                            (_drs[i]["receiver_code"].ToString() == "01"))
                        {
                            _strLine.Append(GetStringAsLength("Y1", 2, true, ' '));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength("Y2", 2, true, ' '));
                        }
                        //전달일자
                        _strLine.Append(GetStringAsLength(_drs[i]["card_delivery_date"].ToString().Replace("-", ""), 8, true, ' '));
                        //민증번호
                        //민증번호
                        if (_drs[i]["card_result_status"].ToString() == "61")
                        {
                            _strLine.Append(GetStringAsLength(_drs[i]["customer_ssn"].ToString().Substring(7, 3), 14, true, ' '));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength(ConvertLGSSN(_drs[i]["receiver_SSN"].ToString().Replace("x", "0")), 14, true, ' '));
                        }
                        //_strLine.Append(GetStringAsLength(_drs[i]["receiver_SSN"].ToString().Replace("-", "").Replace("x", ""), 14, true, ' '));
                        //전화번호
                        _strLine.Append(GetStringAsLength(_drs[i]["receiver_tel"].ToString().Replace("-", ""), 15, true, ' '));
                        //제작일자
                        if (_drs[i]["client_register_date"].ToString() == "")
                        {
                            _strLine.Append(GetStringAsLength(_drs[i]["client_send_date"].ToString().Replace("-", ""), 8, true, ' '));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength(_drs[i]["client_register_date"].ToString().Replace("-", ""), 8, true, ' '));
                        }

                        //제작순번
                        _strLine.Append(GetStringAsLength(_drs[i]["client_number"].ToString(), 5, true, ' '));
                        //특송접수일자
                        if (_drs[i]["client_quick_work_date"].ToString() == "")
                            _strLine.Append(GetStringAsLength(_drs[i]["card_in_date"].ToString().Replace("-", ""), 8, true, ' '));
                        else
                            _strLine.Append(GetStringAsLength(_drs[i]["client_quick_work_date"].ToString().Replace("-", ""), 8, true, ' '));
                        //특송접수번호
                        _strLine.Append(GetStringAsLength(_drs[i]["client_send_number"].ToString(), 6, true, ' '));                     
                        //수령인명
                        _strLine.Append(GetStringAsLength(_drs[i]["receiver_name"].ToString(), 40, true, ' '));
                        //관계코드 - 은행 요청 코드
                        _strLine.Append(GetStringAsLength(_drs[i]["receiver_code_change"].ToString(), 3, true, ' '));
                        _strLine.Append(GetStringAsLength("", 1, true, ' '));                                                           //예비
                        _strLine.Append(GetStringAsLength(ConvertAgree(_drs[i]["card_agree1"].ToString()), 1, true, ' '));
                        _strLine.Append(GetStringAsLength(ConvertAgree(_drs[i]["card_agree2"].ToString()), 1, true, ' '));
                        _strLine.Append(GetStringAsLength(ConvertAgree(_drs[i]["card_agree3"].ToString()), 1, true, ' '));

                        //특송발송카드 BIN구분코드
                        _strLine.Append(GetStringAsLength(_drs[i]["card_client_no_1"].ToString(), 2, true, ' '));
                        //제휴사코드
                        _strLine.Append(GetStringAsLength(_drs[i]["client_express_code"].ToString(), 4, true, ' '));
                        _strLine.Append(GetStringAsLength(_drs[i]["code"].ToString(), 1, true, ' '));                    //신분증코드
                        _strLine.Append(GetStringAsLength(RemoveDash(_drs[i]["date"].ToString()), 9, true, ' '));                    //발급일자
                        _strLine.Append(GetStringAsLength(_drs[i]["org"].ToString(), 30, true, ' '));                    //발급기관

                        _strLine.Append(GetStringAsLength("", 123, true, ' '));                                                     //예비

                        iCnt++;
                        _sw11.WriteLine(_strLine.ToString());
                    }
                    //2013.07.25 태희철 신마감 끝[E]
                }
                //2013.07.22 태희철 수정 [S] 신마감사용
                _strLine = new StringBuilder(GetStringAsLength("TR" + GetStringAsLength(iCnt.ToString(), 11, false, '0'), 300, true, ' '));
                _sw11.WriteLine(_strLine.ToString());
                //2013.07.22 태희철 수정 [E] 신마감사용

                _strReturn = string.Format("{0}건의 인계데이터 다운 완료", iCnt);
            }
            catch (Exception)
            {
                _strReturn = string.Format("{0}번째 데이터 인계 중 오류 발생", i + 1);
            }
            finally
            {
                if (_sw01 != null) _sw01.Close();
                if (_sw11 != null) _sw11.Close();
            }
            return _strReturn;
        }
        //지역 번호 정리
        private static void ArrangeData(ref DataTable dtable)
        {
            string _strAreaGroup = "", _strPrevGroup = "";
            int _iAreaIndex = 1, _iIndex = -1;
            DataRow[] _dr = dtable.Select("", "card_area_group, card_zipcode, customer_name");
            for (int i = 0; i < _dr.Length; i++)
            {
                _iIndex = int.Parse(_dr[i][0].ToString());
                _strAreaGroup = _dr[i][1].ToString();
                if (_strPrevGroup != _strAreaGroup)
                {
                    _strPrevGroup = _strAreaGroup;
                    _iAreaIndex = 1;
                }
                dtable.Rows[_iIndex - 1][3] = _iAreaIndex;
                _iAreaIndex++;
            }
        }

        //문자중 -를 없앤다
        private static string RemoveDash(string value)
        {
            return value.Replace("-", "");
        }

        //문자에 대해 Length보다 길면 substring하고 짧으면 공백을 넣어서 지정 Length 만큼의 문자를 반환
        private static string GetStringAsLength(string Text, int Length, bool blankPositionAtBack, char chBlank)
        {
            string _strReturn = "";
            System.Text.Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);
            byte[] _byteAry = _encoding.GetBytes(Text);

            if (_byteAry.Length > Length)
            {
                _strReturn = _encoding.GetString(_byteAry, 0, Length);
            }
            else if (_byteAry.Length < Length)
            {
                if (blankPositionAtBack)
                {
                    _strReturn = Text;
                }
                else
                {
                    _strReturn = "";
                }
                for (int i = 0; i < (Length - _byteAry.Length); i++)
                {
                    _strReturn += chBlank;
                }
                if (!blankPositionAtBack)
                {
                    _strReturn += Text;
                }
            }
            else
            {
                _strReturn = Text;
            }
            return _strReturn;
        }
        #region 형식함수
        //우편번호 형식
        private static string ConvertZipcode(string value)
        {
            string _strReturn = value;
            if (_strReturn.Length == 6)
            {
                _strReturn = value.Substring(0, 3) + "-" + value.Substring(3, 3);
            }
            return _strReturn;
        }

        //주민등록번호 형식
        private static string ConvertSSN(string value)
        {
            string _strReturn = value;
            if (_strReturn.Length > 6)
            {
                _strReturn = value.Substring(0, 6) + "-" + value.Substring(6, value.Length - 6);
            }
            else if (_strReturn.Length == 6)
            {
                _strReturn = value.Substring(0, 6) + "-" + value.Substring(6, value.Length - 6);
            }
            else
            {
                _strReturn = value;
            }
            return _strReturn;
        }
        //전화번호 형식
        private static string ConvertTel(string value)
        {
            string _strReturn = "";
            if (value.Substring(0, 2) == "02")
            {
                if (value.Length == 9)
                {
                    _strReturn = value.Substring(0, 2) + "-" + value.Substring(2, 3) + "-" + value.Substring(5, value.Length - 5);
                }
                else if (value.Length == 10)
                {
                    _strReturn = value.Substring(0, 2) + "-" + value.Substring(2, 4) + "-" + value.Substring(6, value.Length - 6);
                }
                else if (value.Length < 9)
                {
                    if (value.Length < 6)
                    {
                        _strReturn = value.Substring(0, 2) + "-" + value.Substring(2, value.Length - 2) + "-";
                    }
                    else if (value.Length >= 6)
                    {
                        _strReturn = value.Substring(0, 2) + "-" + value.Substring(2, 3) + "-" + value.Substring(5, value.Length - 5);
                    }
                }
            }
            else
            {
                if (value.Length == 10)
                {
                    _strReturn = value.Substring(0, 3) + "-" + value.Substring(3, 3) + "-" + value.Substring(6, value.Length - 6);
                }
                else if (value.Length == 11)
                {
                    _strReturn = value.Substring(0, 3) + "-" + value.Substring(2, 4) + "-" + value.Substring(7, value.Length - 7);
                }
                else if (value.Length < 10)
                {
                    if (value.Length < 7)
                    {
                        _strReturn = value.Substring(0, 3) + "-" + value.Substring(3, value.Length - 3) + "-";

                    }
                    else if (value.Length >= 7)
                    {
                        _strReturn = value.Substring(0, 3) + "-" + value.Substring(3, 3) + "-" + value.Substring(6, value.Length - 6);
                    }
                }
            }
            return _strReturn;
        }
        #endregion
        private static string ConvertLGSSN(string value)
        {
            string _strReturn = "";
            if (value.Length < 6)
            {
                _strReturn = value.Substring(0, value.Length - 0);
            }
            else if (value.Length == 6)
            {
                _strReturn = value + "-";
            }
            else if (value.Length == 13)
            {
                _strReturn = value.Substring(0, 6) + "-" + value.Substring(6, 7);
            }
            else
            {
                _strReturn = value.Substring(0, 6) + "-" + value.Substring(6, value.Length - 6);
            }
            return _strReturn;
        }

        #region 반송사유코드
        public static string ReturnType(string value)
        {
            string _strReturn = value;

            switch (value)
            {   
                case "5":
                case "6":
                case "13":
                case "17":
                case "25":
                case "28":
                case "31":
                case "34":
                case "98":
                    _strReturn = "GG";
                    break;
                case "3":
                case "20":
                case "21":
                case "88":
                    _strReturn = "BB";
                    break;
                case "1":
                case "12":
                case "18":
                case "19":
                case "23":
                case "24":
                case "33":
                    _strReturn = "CC";
                    break;
                case "2":
                case "11":
                    _strReturn = "FF";
                    break;
                case "8":
                    _strReturn = "JJ";
                    break;
                case "9":
                case "10":
                    _strReturn = "DD";
                    break;
                case "30":
                    _strReturn = "ZZ";
                    break;
                default:
                    _strReturn = "E1";
                    break;
            }
            return _strReturn;
        }
        #endregion

        private static string ConvertAgree(string value)
        {
            string _strReturn = "Y";

            switch (value)
            {
                case "1":
                    _strReturn = "Y";
                    break;
                case "2":
                    _strReturn = "N";
                    break;
                
            }
            return _strReturn;
        }
    }
    internal class Branches : Collection<BranchCount>
    {
        internal int GetCount(string strBranch)
        {
            bool _bFind = false;
            int _return = 1;

            if (base.Count > 0)
            {
                for (int i = 0; i < base.Count; i++)
                {
                    if (base[i].Branch == strBranch)
                    {
                        _return = base[i].Count + 1;
                        base[i].AddCount();
                        _bFind = true;
                        break;
                    }
                }
            }

            if (!_bFind)
            {
                base.Add(new BranchCount(strBranch));
            }

            return _return;
        }
    }

    internal class BranchCount
    {
        string branch = "";
        int count = 1;

        internal BranchCount(string strBranch)
        {
            branch = strBranch;
            count = 1;
        }

        internal string Branch
        {
            get { return branch; }
        }

        internal int Count
        {
            get { return count; }
        }

        internal void AddCount()
        {
            count++;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace _090_농협_TEST
{
    public class CONVERT
    {
        //기본 인코딩 설정
        private static string strEncoding = "ks_c_5601-1987";
        private static string strCardTypeID = "090";
        private static string strCardTypeName = "NH농협_TEST";

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
            byte[] _byteAry = null;
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
        public static string ConvertRegister(string path, string xmlZipcodePath, string xmlZipcodeAreaPath, string xmlZipcodePath_new, string xmlZipcodeAreaPath_new, string xmlPath)
        {

            System.Text.Encoding _encoding = System.Text.Encoding.GetEncoding("ks_c_5601-1987");	//기본 인코딩	
            StreamReader _sr = null;        										//파일 읽기 스트림
            StreamWriter _swError = null;											//파일 쓰기 스트림
            DataSet _dsetZipcode = null, _dsetZipcdeArea = null;    				//우편번호 관련 DataSet
            DataSet _dsetZipcode_new = null, _dsetZipcdeArea_new = null;			//우편번호 관련 DataSet
            DataTable _dtable = null;												//마스터 저장 테이블
            DataRow _dr = null;
            DataRow[] _drs = null;

            DataTable _dtable_NH = null;
            DataRow _dr_NH = null;
            DataRow[] _drs_NH = null;

            byte[] _byteAry = null;
            string _strReturn = "";
            string _strLine = "";
            string _strZipcode = "", _strAreaType = "", _strAreaGroup = "", _strBranch = "", strCard_type_detail = "", _strOwner_only = "", strZipChk = "";
            int _iSeq = 1, _iErrorCount = 0;
            string strDeliveryChk = "";

            Branches _branches = new Branches();
            try
            {
                #region 테이블 구조
                _dtable = new DataTable("Convert");
                //기본 컬럼
                _dtable.Columns.Add("degree_arrange_number");
                _dtable.Columns.Add("card_area_group");
                _dtable.Columns.Add("card_branch");
                _dtable.Columns.Add("card_area_type");
                _dtable.Columns.Add("area_arrange_number");
                //세부 컬럼
                //자료구분      0
                _dtable.Columns.Add("client_send_number");              //dr[5]     배송일련번호  0
                _dtable.Columns.Add("client_express_code");             //배송차수      0
                //업체 코드 : 국제 02
                _dtable.Columns.Add("client_send_date");                //제작일자      0
                _dtable.Columns.Add("client_number");                   //캐리어일련번호0
                _dtable.Columns.Add("card_issue_detail_code");          //카드봉입구분 : 1=일반, 2=BOX
                _dtable.Columns.Add("card_pt_code");                    //dr[10] 긴급발급구분 : 0=일반, 1=긴급, 2=특급
                _dtable.Columns.Add("card_delivery_place_type");        //수령지 구분       0
                _dtable.Columns.Add("card_design_code");                //배송업무기준코드(레이아웃참조)
                _dtable.Columns.Add("client_quick_seq");                //유형코드      0
                _dtable.Columns.Add("card_request_memo");               //유형명        0
                _dtable.Columns.Add("card_client_no_1");                //dr[15]시작카드번호
                _dtable.Columns.Add("card_bank_account_no");            //끝 카드번호
                _dtable.Columns.Add("card_count");                      //카드매수
                _dtable.Columns.Add("card_address_type1");              //주소체계구분코드
                _dtable.Columns.Add("text10");                          //주소체계구분코드~건물일련번호
                _dtable.Columns.Add("card_zipcode");                    //dr[20]우편번호
                _dtable.Columns.Add("card_address_local");              //주소체계구분코드
                _dtable.Columns.Add("card_address_detail");             //

                _dtable.Columns.Add("customer_name");                   //한글이름
                _dtable.Columns.Add("customer_SSN");                    //생년월일
                _dtable.Columns.Add("card_bank_ID");                    //dr[25]영업점코드
                _dtable.Columns.Add("card_bank_name");                  //관리점 이름       0
                _dtable.Columns.Add("card_zipcode3");                   //관리점 우편번호
                _dtable.Columns.Add("card_address3");                   //관리점 주소
                _dtable.Columns.Add("customer_position");               //관리점 번호
                _dtable.Columns.Add("card_bank_account_tel");           //dr[30]실번호 뒷4자리
                _dtable.Columns.Add("card_mobile_tel");                 //휴대폰 번호
                _dtable.Columns.Add("card_tel1");                       //
                _dtable.Columns.Add("card_tel2");                       // 
                _dtable.Columns.Add("card_tel3");                       //
                _dtable.Columns.Add("card_product_name");               //dr[35]제휴코드
                _dtable.Columns.Add("family_name");                     //제휴코드명
                _dtable.Columns.Add("card_consented");                  //동의서 징구여부0
                _dtable.Columns.Add("client_bank_request_no");          //동의서코드
                _dtable.Columns.Add("card_is_for_owner_only");          //
                _dtable.Columns.Add("card_cooperation2");          //dr[40]거래내용
                _dtable.Columns.Add("text1");                      //필수동의
                _dtable.Columns.Add("choice_agree3");              //카드이용권유
                _dtable.Columns.Add("text2");                      //선택동의
                _dtable.Columns.Add("choice_agree2");              //별지동의구분
                _dtable.Columns.Add("text3");                      //dr[45]멤버스동의

                _dtable.Columns.Add("card_zipcode_new");           //dr[46] 신우편번호
                _dtable.Columns.Add("card_zipcode_kind");          // 
                _dtable.Columns.Add("card_zipcode2_new");          //신주소여부
                _dtable.Columns.Add("card_zipcode2_kind");         //
                _dtable.Columns.Add("card_zipcode3_new");          //dr[50]
                _dtable.Columns.Add("card_zipcode3_kind");         //
                _dtable.Columns.Add("customer_memo");              //
                _dtable.Columns.Add("card_number");                //카드넘버
                _dtable.Columns.Add("card_barcode_new");           //카드 뉴바코드
                _dtable.Columns.Add("client_request_memo");        //dr[55]별지메모
                _dtable.Columns.Add("card_issue_type_code");       //dr[56]발급구분



                //2011-09-28 태희철 수정[S]
                //신주소구분
                //건물일련번호
                //도로명주소
                //2011-09-28[E]
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
                #endregion

                #region 테이블 저장
                //파일 읽기 Stream과 오류시 저장할 쓰기 Stream 생성
                _sr = new StreamReader(path, _encoding);
                _swError = new StreamWriter(path + ".Error", false, _encoding);



                while ((_strLine = _sr.ReadLine()) != null)
                {
                    if (_iSeq == 1)
                    {
                        strCard_type_detail = _strLine.Substring(_strLine.Length - 7, 7);
                    }
                    //인코딩, byte 배열로 담기
                    _byteAry = _encoding.GetBytes(_strLine);

                    if (_encoding.GetString(_byteAry, 19, _byteAry.Length - 20).Replace(" ", "") == "")
                        continue;

                    _dr = _dtable.NewRow();
                    _dr[0] = _iSeq;
                    _dr[5] = _encoding.GetString(_byteAry, 2, 17);
                    _dr[6] = _encoding.GetString(_byteAry, 19, 2);
                    _dr[7] = _encoding.GetString(_byteAry, 23, 8);
                    _dr[8] = _encoding.GetString(_byteAry, 31, 9);
                    _dr[9] = _encoding.GetString(_byteAry, 40, 1);
                    _dr[10] = _encoding.GetString(_byteAry, 41, 1);
                    _dr[11] = _encoding.GetString(_byteAry, 42, 1);
                    strDeliveryChk = _encoding.GetString(_byteAry, 43, 3);
                    _dr[12] = strDeliveryChk;
                    //_dr[31] = _encoding.GetString(_byteAry, 536, 1);
                    _dr[13] = _encoding.GetString(_byteAry, 46, 12);
                    _dr[14] = _encoding.GetString(_byteAry, 58, 60);
                    _dr[15] = _encoding.GetString(_byteAry, 118, 19);
                    _dr[16] = _encoding.GetString(_byteAry, 137, 19);
                    _dr[17] = _encoding.GetString(_byteAry, 156, 6);
                    strZipChk = _encoding.GetString(_byteAry, 162, 1);
                    _dr[18] = strZipChk;
                    _dr[19] = _encoding.GetString(_byteAry, 162, 552);
                    //1:지번, 2:도로명주소, 3:건물명주소

                    if (strZipChk == "1")
                    {
                        _strZipcode = _encoding.GetString(_byteAry, 163, 6).Trim();
                    }
                    else
                    {
                        _strZipcode = _encoding.GetString(_byteAry, 179, 5).Trim();
                    }
                    _dr[20] = _strZipcode;

                    _dr[21] = _encoding.GetString(_byteAry, 184, 80);
                    _dr[22] = _encoding.GetString(_byteAry, 264, 100);
                    // 레이아웃상 50byte이나 DB상 40byte
                    // 문제발생 시 DB 수정필요
                    _dr[23] = _encoding.GetString(_byteAry, 714, 40);
                    _dr[24] = _encoding.GetString(_byteAry, 764, 7) + "xxxxxx";
                    _dr[25] = _encoding.GetString(_byteAry, 771, 6);
                    // 레이아웃상 50byte이나 DB상 30byte
                    // 문제발생 시 DB 수정필요
                    _dr[26] = _encoding.GetString(_byteAry, 777, 30);
                    _dr[27] = _encoding.GetString(_byteAry, 827, 6).Trim();
                    _dr[28] = _encoding.GetString(_byteAry, 833, 100).TrimEnd() + _encoding.GetString(_byteAry, 933, 100).TrimEnd();
                    _dr[29] = _encoding.GetString(_byteAry, 1033, 12);
                    _dr[30] = _encoding.GetString(_byteAry, 1045, 4);
                    _dr[31] = _encoding.GetString(_byteAry, 1049, 12);//휴대폰
                    _dr[32] = _encoding.GetString(_byteAry, 1061, 12); //직장
                    _dr[33] = _encoding.GetString(_byteAry, 1073, 12); //자택
                    _dr[34] = _encoding.GetString(_byteAry, 1085, 12); //제3청구지
                    _dr[35] = _encoding.GetString(_byteAry, 1097, 6);  //NH카드상품명
                    _dr[36] = _encoding.GetString(_byteAry, 1103, 50).Replace("'", "");
                    _dr[37] = _encoding.GetString(_byteAry, 1153, 1);
                    _dr[38] = _encoding.GetString(_byteAry, 1154, 17).Trim();
                    //_strOwner_only = _encoding.GetString(_byteAry, 1016, 2);

                    /// strDeliveryChk
                    /// 001 : 일반특송
                    /// 002 : 본인만배송
                    /// 003 : 갱신일반 ~
                    /// 080 : 직접수령

                    if (strDeliveryChk == "002")
                    {
                        _dr[39] = "1";
                        _dr[52] = "본인지정배송 - 신분증확인필요";
                        _dr[56] = "1";
                    }
                    else if (strDeliveryChk == "003")
                    {
                        _dr[39] = "";
                        _dr[52] = "";
                        _dr[56] = "3";
                    }
                    else
                    {
                        _dr[39] = "";
                        _dr[52] = "";
                        _dr[56] = "1";
                    }


                    _dr[40] = _encoding.GetString(_byteAry, 1173, 12);
                    //strchk1 : 필수제공동의
                    //strchk2 : 수집이용권유동의
                    //strchk3 : 선택제공동의
                    //strchk4 : 멤버스제공동의
                    string strchk1 = "", strchk2 = "", strchk3 = "", strchk4 = "";
                    strchk1 = _encoding.GetString(_byteAry, 1185, 5).Replace(' ', '9');
                    //_encoding.GetString(_byteAry, 1190, 3); //공백
                    _dr[41] = strchk1;
                    strchk2 = _encoding.GetString(_byteAry, 1193, 7).Replace(' ', '9');
                    _dr[42] = strchk2;
                    strchk3 = _encoding.GetString(_byteAry, 1200, 3).Replace(' ', '9');
                    _dr[43] = strchk3;
                    _dr[44] = _encoding.GetString(_byteAry, 1203, 1); //별지동의서코드
                    strchk4 = _encoding.GetString(_byteAry, 1204, 9).Replace(' ', '9');
                    _dr[45] = strchk4;

                    _dr[48] = ""; //주소2 사용안함
                    _dr[49] = ""; //주소2 사용안함

                    if (_strZipcode.Length == 5)
                    {
                        _dr[46] = _strZipcode;
                        _dr[47] = "1";
                    }
                    //관리점우편번호
                    if (_dr[27].ToString().Length == 5)
                    {
                        _dr[50] = _dr[27].ToString();
                        _dr[51] = "1";
                    }

                    //동의서 중 별지 구분
                    if (strDeliveryChk == "021" && (_dr[44].ToString() == "1" || _dr[44].ToString() == "2"))
                    {
                        _dr[55] = "별지";
                    }

                    //2011-09-28 태희철 수정[S]
                    //신주소구분
                    //_encoding.GetString(_byteAry, 1018, 1);신주소구분여부 사용안함
                    //건물일련번호
                    //도로명주소
                    //2011-09-28[E]
                    //배송코드
                    _dr[53] = _dr[7].ToString().Substring(1, 7) + _dr[8].ToString();

                    //케리어바코드
                    if (_strZipcode.Length == 5)
                    {
                        _dr[54] = _dr[7].ToString() + _dr[8].ToString() + _strZipcode + " " + _dr[12].ToString() + "02";
                    }
                    else
                    {
                        _dr[54] = _dr[7].ToString() + _dr[8].ToString() + _strZipcode + _dr[12].ToString() + "02";
                    }

                    if (_strZipcode != "")
                    {
                        //지역 분류 선택
                        if (_strZipcode.Length == 5)
                        {
                            _drs = _dsetZipcdeArea_new.Tables[0].Select("zipcode_new = " + _strZipcode);
                        }
                        else
                        {
                            _drs = _dsetZipcdeArea.Tables[0].Select("zipcode = " + _strZipcode);
                        }


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
                        if (_strZipcode.Length == 5)
                        {
                            _drs = _dsetZipcode_new.Tables[0].Select("zipcode_new = " + _strZipcode);
                        }
                        else
                        {
                            _drs = _dsetZipcode.Tables[0].Select("zipcode = " + _strZipcode);
                        }

                        if (_drs.Length > 0)
                        {
                            //시군구 구분 코드 : A,B,C,D...
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

                if (strCard_type_detail.Substring(0, 5) == "09021")
                {
                    //변환에 성공했다면
                    if (_iErrorCount < 1)
                    {
                        _swError.Close();
                        _sr.Close();

                        try
                        {
                            //_dtable의 Row를 재정렬
                            _drs_NH = _dtable.Select("", "card_bank_ID");

                            _dtable_NH = new DataTable("Convert2");
                            //기본 컬럼
                            _dtable_NH.Columns.Add("degree_arrange_number");
                            _dtable_NH.Columns.Add("card_area_group");
                            _dtable_NH.Columns.Add("card_branch");
                            _dtable_NH.Columns.Add("card_area_type");
                            _dtable_NH.Columns.Add("area_arrange_number");
                            //세부 컬럼
                            //자료구분      0
                            _dtable_NH.Columns.Add("client_send_number");              //dr[5]     배송일련번호  0
                            _dtable_NH.Columns.Add("client_express_code");             //배송차수      0
                            //업체 코드 : 국제 02
                            _dtable_NH.Columns.Add("client_send_date");                //제작일자      0
                            _dtable_NH.Columns.Add("client_number");                   //캐리어일련번호0
                            _dtable_NH.Columns.Add("card_issue_detail_code");          //카드봉입구분 : 1=일반, 2=BOX
                            _dtable_NH.Columns.Add("card_pt_code");                    //dr[10] 긴급발급구분 : 0=일반, 1=긴급, 2=특급
                            _dtable_NH.Columns.Add("card_delivery_place_type");        //수령지 구분       0
                            _dtable_NH.Columns.Add("card_design_code");                //배송업무기준코드(레이아웃참조)
                            _dtable_NH.Columns.Add("client_quick_seq");                //유형코드      0
                            _dtable_NH.Columns.Add("card_request_memo");               //유형명        0
                            _dtable_NH.Columns.Add("card_client_no_1");                //dr[15]시작카드번호
                            _dtable_NH.Columns.Add("card_bank_account_no");            //끝 카드번호
                            _dtable_NH.Columns.Add("card_count");                      //카드매수
                            _dtable_NH.Columns.Add("card_address_type1");              //카드매수
                            _dtable_NH.Columns.Add("text10");                          //주소체계구분코드~건물일련번호
                            _dtable_NH.Columns.Add("card_zipcode");                    //dr[20]우편번호
                            _dtable_NH.Columns.Add("card_address_local");              //주소체계구분코드
                            _dtable_NH.Columns.Add("card_address_detail");             //

                            _dtable_NH.Columns.Add("customer_name");                   //한글이름
                            _dtable_NH.Columns.Add("customer_SSN");                    //생년월일
                            _dtable_NH.Columns.Add("card_bank_ID");                    //dr[25]영업점코드
                            _dtable_NH.Columns.Add("card_bank_name");                  //관리점 이름       0
                            _dtable_NH.Columns.Add("card_zipcode3");                   //관리점 우편번호
                            _dtable_NH.Columns.Add("card_address3");                   //관리점 주소
                            _dtable_NH.Columns.Add("customer_position");               //관리점 번호
                            _dtable_NH.Columns.Add("card_bank_account_tel");           //dr[30]실번호 뒷4자리
                            _dtable_NH.Columns.Add("card_mobile_tel");                 //휴대폰 번호
                            _dtable_NH.Columns.Add("card_tel1");                       //
                            _dtable_NH.Columns.Add("card_tel2");                       // 
                            _dtable_NH.Columns.Add("card_tel3");                       //
                            _dtable_NH.Columns.Add("card_product_name");               //dr[35]제휴코드
                            _dtable_NH.Columns.Add("family_name");                     //제휴코드명
                            _dtable_NH.Columns.Add("card_consented");                  //동의서 징구여부0
                            _dtable_NH.Columns.Add("client_bank_request_no");          //동의서코드
                            _dtable_NH.Columns.Add("card_is_for_owner_only");          //
                            _dtable_NH.Columns.Add("card_cooperation2");          //dr[40]인수증구분코드
                            _dtable_NH.Columns.Add("text1");                      //필수동의
                            _dtable_NH.Columns.Add("choice_agree3");              //카드이용권유
                            _dtable_NH.Columns.Add("text2");                      //선택동의
                            _dtable_NH.Columns.Add("choice_agree2");              //별지동의구분
                            _dtable_NH.Columns.Add("text3");                      //멤버스동의

                            _dtable_NH.Columns.Add("card_zipcode_new");           //dr[45] 신우편번호
                            _dtable_NH.Columns.Add("card_zipcode_kind");          // 
                            _dtable_NH.Columns.Add("card_zipcode2_new");          //신주소여부
                            _dtable_NH.Columns.Add("card_zipcode2_kind");         //
                            _dtable_NH.Columns.Add("card_zipcode3_new");          //
                            _dtable_NH.Columns.Add("card_zipcode3_kind");         //dr[50]
                            _dtable_NH.Columns.Add("customer_memo");              //
                            _dtable_NH.Columns.Add("card_number");                //카드넘버
                            _dtable_NH.Columns.Add("card_barcode_new");           //카드 뉴바코드
                            _dtable_NH.Columns.Add("client_request_memo");        //별지메모
                            _dtable_NH.Columns.Add("card_issue_type_code");       //dr[55]발급구분

                            _iSeq = 1;

                            //degree_arrange_number (Total) 값을 재정의 하기 위하여 _branches를 초기화한다.
                            _branches.Clear();
                            //_dtable의 Row를 재정렬하여 _dtable_NH에 담는다
                            for (int i = 0; i < _drs_NH.Length; i++)
                            {
                                _dr_NH = _dtable_NH.NewRow();
                                for (int k = 1; k < _drs_NH[i].ItemArray.Length; k++)
                                {
                                    _dr_NH[0] = _iSeq;
                                    //k == 4 : degree_arrange_number (Total) 값은 재정의를 한다
                                    if (k == 4)
                                    {
                                        _dr_NH[k] = _branches.GetCount(_drs_NH[i].ItemArray[2].ToString());
                                    }
                                    else
                                    {
                                        _dr_NH[k] = _drs_NH[i].ItemArray[k].ToString();
                                    }
                                }
                                _dtable_NH.Rows.Add(_dr_NH);
                                _iSeq++;
                            }
                            //재정렬(card_bank_ID)된 데이터를 보내준다
                            _dtable_NH.WriteXml(xmlPath);
                        }
                        catch (Exception ex)
                        {

                            MessageBox.Show(ex.Message);
                        }
                        _strReturn = string.Format("{0}건의 데이터 변환 성공", _iSeq - 1);
                    }
                    else
                    {
                        _strReturn = string.Format("{0}건 변환, 우편번호 미등록 {1}건 실패", _iSeq - 1, _iErrorCount);
                    }
                }
                else if (strCard_type_detail.Substring(0, 5) == "09031")
                {
                    //변환에 성공했다면
                    if (_iErrorCount < 1)
                    {
                        _swError.Close();
                        _sr.Close();

                        try
                        {
                            //_dtable의 Row를 재정렬
                            _drs_NH = _dtable.Select("", "card_consented DESC, card_bank_ID ASC");

                            _dtable_NH = new DataTable("Convert2");
                            //기본 컬럼
                            _dtable_NH.Columns.Add("degree_arrange_number");
                            _dtable_NH.Columns.Add("card_area_group");
                            _dtable_NH.Columns.Add("card_branch");
                            _dtable_NH.Columns.Add("card_area_type");
                            _dtable_NH.Columns.Add("area_arrange_number");
                            //세부 컬럼
                            //자료구분      0
                            _dtable_NH.Columns.Add("client_send_number");              //dr[5]     배송일련번호  0
                            _dtable_NH.Columns.Add("client_express_code");             //배송차수      0
                            //업체 코드 : 국제 02
                            _dtable_NH.Columns.Add("client_send_date");                //제작일자      0
                            _dtable_NH.Columns.Add("client_number");                   //캐리어일련번호0
                            _dtable_NH.Columns.Add("card_issue_detail_code");          //카드봉입구분 : 1=일반, 2=BOX
                            _dtable_NH.Columns.Add("card_pt_code");                    //dr[10] 긴급발급구분 : 0=일반, 1=긴급, 2=특급
                            _dtable_NH.Columns.Add("card_delivery_place_type");        //수령지 구분       0
                            _dtable_NH.Columns.Add("card_design_code");                //배송업무기준코드(레이아웃참조)
                            _dtable_NH.Columns.Add("client_quick_seq");                //유형코드      0
                            _dtable_NH.Columns.Add("card_request_memo");               //유형명        0
                            _dtable_NH.Columns.Add("card_client_no_1");                //dr[15]시작카드번호
                            _dtable_NH.Columns.Add("card_bank_account_no");            //끝 카드번호
                            _dtable_NH.Columns.Add("card_count");                      //카드매수
                            _dtable_NH.Columns.Add("card_address_type1");              //카드매수
                            _dtable_NH.Columns.Add("text10");                          //주소체계구분코드~건물일련번호
                            _dtable_NH.Columns.Add("card_zipcode");                    //dr[20]우편번호
                            _dtable_NH.Columns.Add("card_address_local");              //주소체계구분코드
                            _dtable_NH.Columns.Add("card_address_detail");             //

                            _dtable_NH.Columns.Add("customer_name");                   //한글이름
                            _dtable_NH.Columns.Add("customer_SSN");                    //생년월일
                            _dtable_NH.Columns.Add("card_bank_ID");                    //dr[25]영업점코드
                            _dtable_NH.Columns.Add("card_bank_name");                  //관리점 이름       0
                            _dtable_NH.Columns.Add("card_zipcode3");                   //관리점 우편번호
                            _dtable_NH.Columns.Add("card_address3");                   //관리점 주소
                            _dtable_NH.Columns.Add("customer_position");               //관리점 번호
                            _dtable_NH.Columns.Add("card_bank_account_tel");           //dr[30]실번호 뒷4자리
                            _dtable_NH.Columns.Add("card_mobile_tel");                 //휴대폰 번호
                            _dtable_NH.Columns.Add("card_tel1");                       //
                            _dtable_NH.Columns.Add("card_tel2");                       // 
                            _dtable_NH.Columns.Add("card_tel3");                       //
                            _dtable_NH.Columns.Add("card_product_name");               //dr[35]제휴코드
                            _dtable_NH.Columns.Add("family_name");                     //제휴코드명
                            _dtable_NH.Columns.Add("card_consented");                  //동의서 징구여부0
                            _dtable_NH.Columns.Add("client_bank_request_no");          //동의서코드
                            _dtable_NH.Columns.Add("card_is_for_owner_only");          //
                            _dtable_NH.Columns.Add("card_cooperation2");          //dr[40]
                            _dtable_NH.Columns.Add("text1");                      //필수동의
                            _dtable_NH.Columns.Add("choice_agree3");              //카드이용권유
                            _dtable_NH.Columns.Add("text2");                      //선택동의
                            _dtable_NH.Columns.Add("choice_agree2");              //별지동의구분
                            _dtable_NH.Columns.Add("text3");                      //멤버스동의

                            _dtable_NH.Columns.Add("card_zipcode_new");           //dr[45] 신우편번호
                            _dtable_NH.Columns.Add("card_zipcode_kind");          // 
                            _dtable_NH.Columns.Add("card_zipcode2_new");          //신주소여부
                            _dtable_NH.Columns.Add("card_zipcode2_kind");         //
                            _dtable_NH.Columns.Add("card_zipcode3_new");          //
                            _dtable_NH.Columns.Add("card_zipcode3_kind");         //dr[50]
                            _dtable_NH.Columns.Add("customer_memo");              //
                            _dtable_NH.Columns.Add("card_number");                //카드넘버
                            _dtable_NH.Columns.Add("card_barcode_new");           //카드 뉴바코드
                            _dtable_NH.Columns.Add("client_request_memo");        //별지메모
                            _dtable_NH.Columns.Add("card_issue_type_code");       //dr[55]발급구분

                            _iSeq = 1;

                            //area_arrange_number (Total) 값을 재정의 하기 위하여 _branches를 초기화한다.
                            _branches.Clear();
                            //_dtable의 Row를 재정렬하여 _dtable_NH에 담는다
                            for (int i = 0; i < _drs_NH.Length; i++)
                            {
                                _dr_NH = _dtable_NH.NewRow();
                                for (int k = 1; k < _drs_NH[i].ItemArray.Length; k++)
                                {
                                    _dr_NH[0] = _iSeq;
                                    //k == 4 : area_arrange_number (Total) 값은 재정의를 한다
                                    if (k == 4)
                                    {
                                        _dr_NH[k] = _branches.GetCount(_drs_NH[i].ItemArray[2].ToString());
                                    }
                                    else
                                    {
                                        _dr_NH[k] = _drs_NH[i].ItemArray[k].ToString();
                                    }
                                }
                                _dtable_NH.Rows.Add(_dr_NH);
                                _iSeq++;
                            }
                            //재정렬(card_bank_ID)된 데이터를 보내준다
                            _dtable_NH.WriteXml(xmlPath);
                        }
                        catch (Exception ex)
                        {

                            MessageBox.Show(ex.Message);
                        }
                        _strReturn = string.Format("{0}건의 데이터 변환 성공", _iSeq - 1);
                    }
                    else
                    {
                        _strReturn = string.Format("{0}건 변환, 우편번호 미등록 {1}건 실패", _iSeq - 1, _iErrorCount);
                    }
                }
                else if (_iErrorCount < 1)
                {
                    _swError.Close();
                    _sr.Close();
                    _dtable.WriteXml(xmlPath);
                    _strReturn = string.Format("{0}건의 데이터 변환 성공", _iSeq - 1);
                }
                else
                {
                    _strReturn = string.Format("{0}건 변환, 우편번호 미등록 {1}건 실패", _iSeq - 1, _iErrorCount);
                }
                #endregion
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
        



        //마감
        public static string ConvertResult(System.Data.DataTable dtable, string fileName)
        {
            return ConvertResultGubun(dtable, fileName, 3);
        }

        //일일마감자료
        public static string ConvertResultDay(System.Data.DataTable dtable, string fileName)
        {
            return "카드사마감일 자료 다운만 가능합니다.";
        }

        private static string ConvertResultGubun(System.Data.DataTable dtable, string fileName, int temp)
        {
            Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            StreamWriter _sw00 = null, _sw01 = null, _sw02 = null;              //파일 쓰기 스트림
            int i = 0;
            StringBuilder _strLine = new StringBuilder("");
            string _strReturn = "", _strStatus = "", strDong = "", strCard_type_detail = "";
            int temp_0 = 0, temp_1 = 0, temp_2 = 0;
            try
            {
                // temp : 2 일일 // 3 = 마감
                string tempday = DateTime.Now.ToString("yyyyMMdd");

                _sw00 = new StreamWriter(fileName + "농협_재방", true, _encoding);
                _sw01 = new StreamWriter(fileName + "KU0.bissu2022." + tempday + ".00.I.01", true, _encoding);
                _sw02 = new StreamWriter(fileName + "KU0.bissu2022." + tempday + ".00.I.02", true, _encoding);

                //_strLine = new StringBuilder(GetStringAsLength("", 11, true, ' '));
                _strLine.Append(GetStringAsLength("FH", 2, true, ' '));
                _strLine.Append(GetStringAsLength(tempday, 8, true, ' '));
                _strLine.Append("{0}");
                _strLine.Append(GetStringAsLength("", 788, true, ' '));

                _strLine = new StringBuilder(string.Format(_strLine.ToString(), GetStringAsLength(dtable.Rows[0]["client_express_code"].ToString(), 2, true, ' ')));

                _sw01.WriteLine(_strLine);
                _sw02.WriteLine(_strLine);
                _sw00.WriteLine(_strLine);

                for (i = 0; i < dtable.Rows.Count; i++)
                {
                    _strStatus = dtable.Rows[i]["card_delivery_status"].ToString();
                    strCard_type_detail = dtable.Rows[i]["card_type_detail"].ToString();
                    //영업점특송 사용
                    strDong = dtable.Rows[i]["card_consented"].ToString();

                    _strLine = new StringBuilder(GetStringAsLength("FD", 2, true, ' '));
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_send_number"].ToString(), 17, true, ' ')); //배송일련번호
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["client_express_code"].ToString(), 2, true, ' ')); //배송차수
                    _strLine.Append(GetStringAsLength("02", 2, true, ' '));                                             //업체코드

                    //2012-01-17 태희철 정리
                    //temp_0, temp_1, temp_2 : 총건수 1:배송, 2:반송
                    string _strStatusTemp = "00";  //배송 결과

                    if (_strStatus == "1" || _strStatus == "7")
                    {
                        if (dtable.Rows[i]["receiver_code"].ToString() == "98")
                        {
                            _strStatusTemp = "05";
                        }
                        else
                        {
                            _strStatusTemp = "00";
                        }
                    }
                    else if (_strStatus == "2" || _strStatus == "3")
                    {
                        _strStatusTemp = "07";
                    }
                    else if (_strStatus == "6")
                    {
                        _strStatusTemp = "08";
                    }
                    else
                    {
                        _strStatusTemp = "03";
                    }

                    if ((strCard_type_detail == "0903201" || strCard_type_detail == "0904101") && _strStatusTemp == "05" && dtable.Rows[i]["card_issue_detail_code"].ToString() == "3")
                        _strLine.Append(GetStringAsLength("06", 2, true, ' '));
                    else
                        _strLine.Append(GetStringAsLength(_strStatusTemp, 2, true, ' '));

                    _strLine.Append(GetStringAsLength(temp.ToString(), 1, true, ' '));      //배송결과수신방법코드   
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_client_no_1"].ToString(), 19, true, ' ')); //시작카드번호
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_bank_account_no"].ToString(), 19, true, ' '));//끝카드번호

                    //2012.06.28 태희철 수정
                    if ((strCard_type_detail == "0903201" || strCard_type_detail == "0904101") && strDong == "1")
                    {
                        _strLine.Append("0"); //동의서징구 결과
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_consented"].ToString(), 1, true, ' ')); //동의서징구 결과
                    }

                    //_strLine.Append(GetStringAsLength(dtable.Rows[i]["card_delivery_place_type"].ToString(), 1, true, ' ')); //수령지

                    if (_strStatusTemp != "05")
                    {
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_delivery_date"].ToString().Replace("-", ""), 8, true, ' ')); //배송완료일
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_name"].ToString(), 40, true, ' ')); //수령인 이름
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_SSN"].ToString(), 7, true, ' '));   //수령인 민증
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["delivery_result_regdate"].ToString().Replace("-", ""), 8, true, ' ')); //배송완료일
                        _strLine.Append(GetStringAsLength("", 40, true, ' ')); //수령인 이름
                        _strLine.Append(GetStringAsLength("", 7, true, ' '));
                    }

                    if (_strStatus == "2" || _strStatus == "3")
                    {
                        _strLine.Append(GetStringAsLength("", 2, true, ' '));
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["return_code_change"].ToString(), 2, true, ' '));
                    }
                    else if (_strStatus == "6")
                    {
                        _strLine.Append(GetStringAsLength("", 2, true, ' '));
                        _strLine.Append(GetStringAsLength("", 2, true, ' '));
                    }
                    else
                    {
                        if (_strStatusTemp != "05")
                            _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_code_change"].ToString(), 2, true, ' '));
                        else
                            _strLine.Append(GetStringAsLength("", 2, true, ' '));

                        _strLine.Append(GetStringAsLength("", 2, true, ' '));
                    }

                    if (_strStatusTemp == "05")
                    {
                        _strLine.Append(GetStringAsLength(dtable.Rows[i]["receiver_SSN"].ToString(), 14, true, ' '));
                        // 농협등기요금
                        //2012.11.20 태희철 수정 등기요금 1960 -> 2200
                        //2014.02.03 태희철 수정 등기요금 2230 -> 2440
                        //2017.04.01 태희철 수정 등기요금 2440 -> 2470
                        //2017.12.20 태희철 수정 등기요금 2440 -> 2670
                        //_strLine.Append(GetStringAsLength("2230", 6, true, ' '));
                        DateTime CardInDate = DateTime.Parse(dtable.Rows[i]["card_in_date"].ToString());
                        DateTime dtDong_date = DateTime.Parse("2017-12-01");

                        if (CardInDate < dtDong_date)
                        {
                            _strLine.Append(GetStringAsLength("2470", 6, true, ' '));
                        }
                        else
                        {
                            _strLine.Append(GetStringAsLength("2670", 6, true, ' '));
                        }
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength("", 14, true, ' '));
                        _strLine.Append(GetStringAsLength("", 6, true, ' '));
                    }

                    //수령지구분값(재청구지구분)
                    //0-변경없음, 1-자택, 2-직장, 3-제3청구지
                    if (dtable.Rows[i]["change_type"].ToString().Trim().Length > 0)
                    {
                        _strLine.Append(GetStringAsLength("Y", 1, true, ' '));
                    }
                    else
                    {
                        _strLine.Append(GetStringAsLength("N", 1, true, ' '));
                    }

                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["text10"].ToString(), 552, true, ' '));
                    _strLine.Append(GetStringAsLength(dtable.Rows[i]["card_cooperation2"].ToString(), 12, true, ' '));

                    _strLine.Append(GetStringAsLength("1111", 4, true, ' '));
                    _strLine.Append(GetStringAsLength("", 21, true, ' '));
                    _strLine.Append(GetStringAsLength("", 66, true, ' '));

                    // 2012-01-17 태희철 수정
                    if (_strStatus == "1")
                    {
                        temp_1++;
                        _sw01.WriteLine(_strLine);
                    }
                    else if (_strStatus == "2" || _strStatus == "3" || _strStatus == "6")
                    {
                        temp_2++;
                        _sw02.WriteLine(_strLine);
                    }
                    else
                    {
                        temp_0++;
                        _sw00.WriteLine(_strLine);
                    }
                }
                _strLine = new StringBuilder("FT");
                _strLine.Append(GetStringAsLength(temp_1.ToString(), 8, false, '0'));
                _strLine.Append(GetStringAsLength("", 790, true, ' '));

                _sw01.WriteLine(_strLine);

                _strLine = new StringBuilder("FT");
                _strLine.Append(GetStringAsLength(temp_2.ToString(), 8, false, '0'));
                _strLine.Append(GetStringAsLength("", 790, true, ' '));

                _sw02.WriteLine(_strLine);

                _strLine = new StringBuilder("FT");
                _strLine.Append(GetStringAsLength(temp_0.ToString(), 8, false, '0'));
                _strLine.Append(GetStringAsLength("", 790, true, ' '));

                _sw00.WriteLine(_strLine);

                // 2012-01-17 태희철 수정

                _strReturn = string.Format("{0}건의 인계데이터 다운 완료", i);
            }
            catch (Exception)
            {
                _strReturn = string.Format("{0}번째 데이터 인계 중 오류 발생", i + 1);
            }
            finally
            {
                if (_sw00 != null) _sw00.Close();
                if (_sw01 != null) _sw01.Close();
                if (_sw02 != null) _sw02.Close();
            }
            return _strReturn;
        }




        #region 기타 기능
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
        #endregion
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

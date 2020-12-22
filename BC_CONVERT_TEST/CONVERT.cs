using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;


namespace BC_CONVERT
{
    public class CONVERT
    {
        //기본 인코딩 설정
        private static string strEncoding = "ks_c_5601-1987";
        private static string strCardTypeID = "001_CONV";
        private static string strCardTypeName = "비씨컨버터";

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
            string _strReturn = "";

            return _strReturn;
        }

        //등록 자료 생성
        public static string ConvertRegister(string path, string xmlZipcodeAreaPath, string xmlZipcodePath)
        {
            System.Text.Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            //FileInfo _fi = null;
            StreamReader _sr = null;																					//파일 읽기 스트림
            StreamWriter _swError = null;
            StreamWriter _sw = null;

            byte[] _byteAry = null;
            string _strBankID = "", _strBankName = "", _strZipcode = "", _strAreaGroup = "", strBank_code = "", strIssue_code = "";
            //2013.06.26 태희철 _strValue : 영업점코드 : 2 = 영업점, 3 = 제3영업점, 4 = 강제영업점
            string _strReturn = "", _strLine = "", _strValue = null;
            DataTable _dtable = null;
            DataSet _dsetZipcodeArea = null;
            //DataRow _dr = null;
            DataRow[] _drs = null;
            int _iCount = 0;
            string _strDong = null, _strVIP = null, _strSC = null, _strYe = null, _strChk_tel = "", strBank_Chk = "";
            try
            {
                _dtable = new DataTable("CONVERT");
                _dtable.Columns.Add("card_bank_ID");
                _dtable.Columns.Add("card_zipcode");
                _dtable.Columns.Add("data");

                _dsetZipcodeArea = new DataSet();
                _dsetZipcodeArea.ReadXml(xmlZipcodeAreaPath);

                //파일 읽기 Stream과 오류시 저장할 쓰기 Stream 생성
                _sr = new StreamReader(path, _encoding);
                _swError = new StreamWriter(path + ".Error", false, _encoding);

                while ((_strLine = _sr.ReadLine()) != null)
                {
                    _byteAry = _encoding.GetBytes(_strLine);

                    
                    _strBankID = _encoding.GetString(_byteAry, 8, 2);
                    _strZipcode = _encoding.GetString(_byteAry, 139, 6);
                    //2011-10-14 태희철 수정
                    //영업점코드 : 2 = 영업점, 3 = 제3영업점, 4 = 강제영업점
                    _strValue = _encoding.GetString(_byteAry, 145, 1);
                    _strDong = _encoding.GetString(_byteAry, 247, 1);
                    //2012-05-15 태희철 추가
                    _strVIP = _encoding.GetString(_byteAry, 296, 1);
                    //2012-05-15 태희철 추가
                    _strSC = _encoding.GetString(_byteAry, 297, 2);
                    //2012-06-04 태희철 추가 Ye치과 구분
                    _strYe = _encoding.GetString(_byteAry, 245, 1);
                    //우리은행 통신사별지
                    _strChk_tel = _encoding.GetString(_byteAry, 1765, 1);

                    strBank_Chk = _encoding.GetString(_byteAry, 250, 3);

                    _drs = _dsetZipcodeArea.Tables[0].Select("zipcode=" + _strZipcode);

                    _iCount++;


                    // [1] 동의서 1~8은 동의서 | 299,1 의 1 or 2 는 긴급은 따로 분류
                    // 2012-05-22 태희철 수정 BLISS, 인피니트 카드 추가
                    //2013.04.08 태희철 수정 인피니트 업무종료 8번 코드도 일반으로 취급
                    //if (_strVIP == "8")
                    //{
                    //    _sw = new StreamWriter("INP_3000", true, _encoding);
                    //}
                    //2013.04.02 태희철 추가[S] 비씨카드 다이아/시그니쳐
                    //if (_byteAry.Length != 1739)
                    //if (_byteAry.Length != 1741)
                    //if (_byteAry.Length != 1765)
                    if (_byteAry.Length != 1766)
                    {
                        _sw = new StreamWriter(path + "총byte오류", true, _encoding);
                        _sw.WriteLine(_strLine);
                    }
                    else if (_strVIP == "6")
                    {
                        if (_strValue == "2" || _strValue == "3" || _strValue == "4")
                        {
                            if (_encoding.GetString(_byteAry, 245, 1) == "3")
                            {
                                _sw = new StreamWriter(path + "bc14000_Dia_본인", true, _encoding);
                                _sw.WriteLine(_strLine + "0B" + "0011113");
                            }
                            else
                            {
                                _sw = new StreamWriter(path + "bc6000_Dia", true, _encoding);
                                _sw.WriteLine(_strLine + "0B" + "0011106");
                            }
                        }
                        //다이아/시그니처 동의서도 일반동의서와 동일하게 구분
                        else if (_strDong == "1" || _strDong == "2" || _strDong == "3" || _strDong == "4"
                        || _strDong == "5" || _strDong == "6" || _strDong == "7" || _strDong == "8" || _strDong == "9")
                        {
                            _strBankName = Convert_Bank_ID(_strBankID);

                            //출력용지 구분파일
                            //_sw = new StreamWriter(path + "Dia_DONG_26000_" + _strBankName, true, _encoding);
                            //_sw.WriteLine(_strLine + "0A" + "1111111");
                            //_sw.Close();

                            ////등록파일
                            //_sw = new StreamWriter(path + "Dia_DONG_26000", true, _encoding);
                            //_sw.WriteLine(_strLine + "0A" + "0012114");

                            switch (Convert_Bank_ID(_strBankID))
                            {
                                case "국민":
                                    _sw = new StreamWriter(path + "bcd2000_국민", true, _encoding);
                                    _sw.WriteLine(_strLine + "0A" + "0012102");
                                    break;
                                case "농협":
                                    _sw = new StreamWriter(path + "bcd8000_농협", true, _encoding);
                                    _sw.WriteLine(_strLine + "0A" + "0012107");
                                    break;
                                case "우리":
                                    if (_strDong == "4")
                                    {
                                        _sw = new StreamWriter(path + "bcd31000_우리국기원", true, _encoding);
                                        _sw.WriteLine(_strLine + "0A" + "0012128");
                                    }
                                    else
                                    {
                                        if (_strChk_tel == "1" ||_strChk_tel == "2" ||_strChk_tel == "3")
                                        {
                                            _sw = new StreamWriter(path + "bcd32000_우리(통신)", true, _encoding);
                                            _sw.WriteLine(_strLine + "0A" + "0012129");
                                        }
                                        else
                                        {
                                            _sw = new StreamWriter(path + "bcd1000_우리", true, _encoding);
                                            _sw.WriteLine(_strLine + "0A" + "0012101");
                                        }
                                    }
                                    break;
                                case "하나":
                                    _sw = new StreamWriter(path + "bcd5000_하나", true, _encoding);
                                    _sw.WriteLine(_strLine + "0A" + "0012104");
                                    break;
                                case "SC제일":
                                    if (_strDong == "5")
                                    {
                                        _sw = new StreamWriter(path + "bcd25000_SC(아이행복)", true, _encoding);
                                        _sw.WriteLine(_strLine + "0A" + "0012122");
                                    }
                                    else if (_strDong == "6")
                                    {
                                        _sw = new StreamWriter(path + "bcd28000_SC(신세계)", true, _encoding);
                                        _sw.WriteLine(_strLine + "0A" + "0012125");
                                    }
                                    else if (_strDong == "9")
                                    {
                                        _sw = new StreamWriter(path + "bcd29000_SC(이마트)", true, _encoding);
                                        _sw.WriteLine(_strLine + "0A" + "0012126");
                                    }
                                    else
                                    {
                                        _sw = new StreamWriter(path + "bcd3000_SC", true, _encoding);
                                        _sw.WriteLine(_strLine + "0A" + "0012103");
                                    }
                                    break;
                                case "기업":
                                    //2014.09.26 태희철 추가 GS칼텍스 보너스 (업무개시 : 09월29일)
                                    if (_strDong == "1")
                                    {
                                        _sw = new StreamWriter(path + "bcd22000_기업칼텍스", true, _encoding);
                                        _sw.WriteLine(_strLine + "0A" + "0012119");
                                    }
                                    else if (_strDong == "2")
                                    {
                                        _sw = new StreamWriter(path + "bcd12000_기업GS", true, _encoding);
                                        _sw.WriteLine(_strLine + "0A" + "0012109");
                                    }
                                    else if (_strDong == "4")
                                    {
                                        _sw = new StreamWriter(path + "bcd13000_기업SK", true, _encoding);
                                        _sw.WriteLine(_strLine + "0A" + "0012110");
                                    }
                                    else if (_strDong == "5")
                                    {
                                        _sw = new StreamWriter(path + "bcd14000_기업에코", true, _encoding);
                                        _sw.WriteLine(_strLine + "0A" + "0012111");
                                    }
                                    else if (_strDong == "9")
                                    {
                                        _sw = new StreamWriter(path + "bcd24000_기업롯데맴버스", true, _encoding);
                                        _sw.WriteLine(_strLine + "0A" + "0012121");
                                    }
                                    else
                                    {
                                        _sw = new StreamWriter(path + "bcd11000_기업공용", true, _encoding);
                                        _sw.WriteLine(_strLine + "0A" + "0012108");
                                    }
                                    break;
                                case "신한":
                                    _sw = new StreamWriter(path + "bcd15000_신한", true, _encoding);
                                    _sw.WriteLine(_strLine + "0A" + "0012112");
                                    break;
                                case "경남":
                                    _sw = new StreamWriter(path + "bcd7000_경남", true, _encoding);
                                    _sw.WriteLine(_strLine + "0A" + "0012106");
                                    break;
                                case "대구":
                                    _sw = new StreamWriter(path + "bcd6000_대구", true, _encoding);
                                    _sw.WriteLine(_strLine + "0A" + "0012105");
                                    break;
                                case "부산":
                                    _sw = new StreamWriter(path + "bcd16000_부산", true, _encoding);
                                    _sw.WriteLine(_strLine + "0A" + "0012113");
                                    break;
                                case "전북":
                                    if (_strDong == "9")
                                    {
                                        _sw = new StreamWriter(path + "bcd4000_전북", true, _encoding);
                                        _sw.WriteLine(_strLine + "0A" + "0012115");
                                    }
                                    else
                                    {
                                        _sw = new StreamWriter(path + "오류_전북4000", true, _encoding);
                                        _sw.WriteLine(_strLine + "0A" + "0012115");
                                    }
                                    //_sw = new StreamWriter(path + "bcd4000_전북", true, _encoding);
                                    //_sw.WriteLine(_strLine + "0A" + "0012115");
                                    break;
                                case "제주":
                                    _sw = new StreamWriter(path + "bcd26000_제주", true, _encoding);
                                    _sw.WriteLine(_strLine + "0A" + "0012123");
                                    break;
                                case "광주":
                                    _sw = new StreamWriter(path + "bcd27000_광주", true, _encoding);
                                    _sw.WriteLine(_strLine + "0A" + "0012124");
                                    break;
                                case "바로":
                                    _sw = new StreamWriter(path + "bcd30000_바로", true, _encoding);
                                    _sw.WriteLine(_strLine + "0A" + "0012127");
                                    break;
                                case "그외":
                                    _sw = new StreamWriter(path + "동의서_영업점코드확인", true, _encoding);
                                    _sw.WriteLine(_strLine + "0A");
                                    break;
                                default:
                                    _sw = new StreamWriter(path + "동의서_영업점코드확인", true, _encoding);
                                    _sw.WriteLine(_strLine + "0A");
                                    break;
                            }
                        }
                        else
                        {
                            if (_encoding.GetString(_byteAry, 245, 1) == "3")
                            {
                                _sw = new StreamWriter(path + "bc14000_Dia_본인", true, _encoding);
                                _sw.WriteLine(_strLine + "0A" + "0011113");
                            }
                            else
                            {
                                _sw = new StreamWriter(path + "bc6000_Dia", true, _encoding);
                                _sw.WriteLine(_strLine + "0A" + "0011106");
                            }
                        }
                    }
                    //2013.05.02 태희철 수정[S] 블리스 차수 통합
                    //2014.11.07 태희철 수정 블리스 동의서의 경우 기업(공용)으로 추가
                    else if (_strVIP == "Z")
                    {
                        if (_strValue == "2" || _strValue == "3" || _strValue == "4")
                        {
                        //    if (_strDong == "1" || _strDong == "2" || _strDong == "3" || _strDong == "4"
                        //|| _strDong == "5" || _strDong == "6" || _strDong == "7" || _strDong == "8" || _strDong == "9")
                        //    {
                        //        _sw = new StreamWriter(path + "bcd11000_기업공용", true, _encoding);
                        //        _sw.WriteLine(_strLine + "0B" + "0012108");
                        //    }
                        //    else
                        //    {
                        //        if (_encoding.GetString(_byteAry, 245, 1) == "3")
                        //        {
                        //            _sw = new StreamWriter(path + "bc13000_BLISS_본인", true, _encoding);
                        //            _sw.WriteLine(_strLine + "0B" + "0011112");
                        //        }
                        //        else
                        //        {
                        //            _sw = new StreamWriter(path + "bc2000_BLISS", true, _encoding);
                        //            _sw.WriteLine(_strLine + "0B" + "0011102");
                        //        }
                        //    }
                            //2016.07.29 특송담당자 수정 요청 동의서 제외처리
                            if (_encoding.GetString(_byteAry, 245, 1) == "3")
                            {
                                _sw = new StreamWriter(path + "bc13000_BLISS_본인", true, _encoding);
                                _sw.WriteLine(_strLine + "0B" + "0011112");
                            }
                            else
                            {
                                _sw = new StreamWriter(path + "bc2000_BLISS", true, _encoding);
                                _sw.WriteLine(_strLine + "0B" + "0011102");
                            }                            
                        }
                        else
                        {
                            if (_strDong == "1" || _strDong == "2" || _strDong == "3" || _strDong == "4"
                        || _strDong == "5" || _strDong == "6" || _strDong == "7" || _strDong == "8" || _strDong == "9")
                            {
                                _sw = new StreamWriter(path + "bcd11000_기업공용", true, _encoding);
                                _sw.WriteLine(_strLine + "0A" + "0012108");
                            }
                            else
                            {
                                if (_encoding.GetString(_byteAry, 245, 1) == "3")
                                {
                                    _sw = new StreamWriter(path + "bc13000_BLISS_본인", true, _encoding);
                                    _sw.WriteLine(_strLine + "0A" + "0011112");
                                }
                                else
                                {
                                    _sw = new StreamWriter(path + "bc2000_BLISS", true, _encoding);
                                    _sw.WriteLine(_strLine + "0A" + "0011102");
                                }
                            }
                            
                        }
                    }
                    else if (_strValue == "2" || _strValue == "3" || _strValue == "4")
                    {
                        _sw = new StreamWriter(path + "be6000", true, _encoding);
                        _sw.WriteLine(_strLine + "0B" + "0013101");
                        //if (_strVIP == "Z")
                        //    _sw = new StreamWriter("BLISS2500", true, _encoding);
                        //else
                        //    _sw = new StreamWriter("be6000", true, _encoding);
                    }
                    //BC_YE치과카드 분류 추가 2012-06-04 태희철 추가
                    else if (_strYe == "2")
                    {
                        _sw = new StreamWriter(path + "BC_YE_1000", true, _encoding);
                        _sw.WriteLine(_strLine + "0A" + "0013401");
                    }
                    // 2012-05-15 태희철 수정 SC은행의 경우 bwe7000차에 합류
                    else if ((_encoding.GetString(_byteAry, 299, 1) != "1" || _encoding.GetString(_byteAry, 299, 1) != "2")
                        && ((_strDong == "1" || _strDong == "2" || _strDong == "3" || _strDong == "4"
                        || _strDong == "5" || _strDong == "6" || _strDong == "7" || _strDong == "8" || _strDong == "9")))
                    {

                        switch (Convert_Bank_ID(_strBankID))
                        {
                            case "국민":
                                _sw = new StreamWriter(path + "bcd2000_국민", true, _encoding);
                                _sw.WriteLine(_strLine + "0A" + "0012102");
                                break;
                            case "농협":
                                _sw = new StreamWriter(path + "bcd8000_농협", true, _encoding);
                                _sw.WriteLine(_strLine + "0A" + "0012107");
                                break;
                            case "우리":
                                if (_strDong == "4")
                                {
                                    _sw = new StreamWriter(path + "bcd31000_우리국기원", true, _encoding);
                                    _sw.WriteLine(_strLine + "0A" + "0012128");
                                }
                                else
                                {
                                    if (_strChk_tel == "1" || _strChk_tel == "2" || _strChk_tel == "3")
                                    {
                                        _sw = new StreamWriter(path + "bcd32000_우리", true, _encoding);
                                        _sw.WriteLine(_strLine + "0A" + "0012129");
                                    }
                                    else
                                    {
                                        _sw = new StreamWriter(path + "bcd1000_우리", true, _encoding);
                                        _sw.WriteLine(_strLine + "0A" + "0012101");
                                    }
                                }
                                break;
                            case "하나":
                                _sw = new StreamWriter(path + "bcd5000_하나", true, _encoding);
                                _sw.WriteLine(_strLine + "0A" + "0012104");
                                break;
                            case "SC제일":
                                if (_strDong == "5")
                                {
                                    _sw = new StreamWriter(path + "bcd25000_SC(아이행복)", true, _encoding);
                                    _sw.WriteLine(_strLine + "0A" + "0012122");   
                                }
                                else if (_strDong == "6")
                                {
                                    _sw = new StreamWriter(path + "bcd28000_SC(신세계)", true, _encoding);
                                    _sw.WriteLine(_strLine + "0A" + "0012125");
                                }
                                else if (_strDong == "9")
                                {
                                    _sw = new StreamWriter(path + "bcd29000_SC(이마트)", true, _encoding);
                                    _sw.WriteLine(_strLine + "0A" + "0012126");
                                }
                                else
                                {
                                    _sw = new StreamWriter(path + "bcd3000_SC", true, _encoding);
                                    _sw.WriteLine(_strLine + "0A" + "0012103");
                                }
                                break;
                            case "기업":
                                if (_strDong == "1")
                                {
                                    _sw = new StreamWriter(path + "bcd22000_기업칼텍스", true, _encoding);
                                    _sw.WriteLine(_strLine + "0A" + "0012119");
                                }
                                else if (_strDong == "2")
                                {
                                    _sw = new StreamWriter(path + "bcd12000_기업GS", true, _encoding);
                                    _sw.WriteLine(_strLine + "0A" + "0012109");
                                }
                                else if (_strDong == "4")
                                {
                                    _sw = new StreamWriter(path + "bcd13000_기업SK", true, _encoding);
                                    _sw.WriteLine(_strLine + "0A" + "0012110");
                                }
                                else if (_strDong == "5")
                                {
                                    _sw = new StreamWriter(path + "bcd14000_기업에코", true, _encoding);
                                    _sw.WriteLine(_strLine + "0A" + "0012111");
                                }
                                else if (_strDong == "9")
                                {
                                    _sw = new StreamWriter(path + "bcd24000_기업롯데맴버스", true, _encoding);
                                    _sw.WriteLine(_strLine + "0A" + "0012121");
                                }
                                else
                                {
                                    _sw = new StreamWriter(path + "bcd11000_기업공용", true, _encoding);
                                    _sw.WriteLine(_strLine + "0A" + "0012108");
                                }
                                break;
                            case "신한":
                                _sw = new StreamWriter(path + "bcd15000_신한", true, _encoding);
                                _sw.WriteLine(_strLine + "0A" + "0012112");
                                break;
                            case "경남":
                                _sw = new StreamWriter(path + "bcd7000_경남", true, _encoding);
                                _sw.WriteLine(_strLine + "0A" + "0012106");
                                break;
                            case "대구":
                                _sw = new StreamWriter(path + "bcd6000_대구", true, _encoding);
                                _sw.WriteLine(_strLine + "0A" + "0012105");
                                break;
                            case "부산":
                                _sw = new StreamWriter(path + "bcd16000_부산", true, _encoding);
                                _sw.WriteLine(_strLine + "0A" + "0012113");
                                break;
                            case "전북":
                                //2014.09.26 태희철 추가 GS칼텍스 보너스 (업무개시 : 09월29일)
                                if (_strDong == "9")
                                {
                                    _sw = new StreamWriter(path + "bcd4000_전북", true, _encoding);
                                    _sw.WriteLine(_strLine + "0A" + "0012115");
                                }
                                else
                                {
                                    _sw = new StreamWriter(path + "오류_전북4000", true, _encoding);
                                    _sw.WriteLine(_strLine + "0A" + "0012115");
                                }
                                //_sw = new StreamWriter(path + "bcd4000_전북", true, _encoding);
                                //_sw.WriteLine(_strLine + "0A" + "0012115");
                                break;
                            case "제주":
                                _sw = new StreamWriter(path + "bcd26000_제주", true, _encoding);
                                _sw.WriteLine(_strLine + "0A" + "0012123");
                                break;
                            case "광주":
                                _sw = new StreamWriter(path + "bcd27000_광주", true, _encoding);
                                _sw.WriteLine(_strLine + "0A" + "0012124");
                                break;
                            case "바로":
                                _sw = new StreamWriter(path + "bcd30000_바로", true, _encoding);
                                _sw.WriteLine(_strLine + "0A" + "0012127");
                                break;
                            case "그외":
                                _sw = new StreamWriter(path + "동의서_영업점코드확인", true, _encoding);
                                _sw.WriteLine(_strLine + "0A");
                                //_sw = GetStreamWriter(_strBankID + ".txt");
                                break;
                            default:
                                _sw = new StreamWriter(path + "동의서_영업점코드확인", true, _encoding);
                                _sw.WriteLine(_strLine + "0A");
                                break;
                        }
                    }
                    else if (_strBankID.Equals("23") && _strSC == "01")
                    {
                        if (_encoding.GetString(_byteAry, 245, 1) == "3")
                        {
                            _sw = new StreamWriter(path + "bwe9000_본인", true, _encoding);
                            _sw.WriteLine(_strLine + "0A" + "0013204");
                        }
                        else
                        {
                            _sw = new StreamWriter(path + "bwe7000", true, _encoding);
                            _sw.WriteLine(_strLine + "0A" + "0013201");
                        }
                    }
                    else
                    {   
                        if (_encoding.GetString(_byteAry, 299, 1) == "2")
                        {
                            //2012-01-03 태희철 수정 5000차를 7000차와 통합
                            //2012-01-16 인수분부터 적용으로 변경
                            //_sw = new StreamWriter("bwe5000", true, _encoding);
                            if (_encoding.GetString(_byteAry, 245, 1) == "3")
                            {
                                _sw = new StreamWriter(path + "bwe9000_본인", true, _encoding);
                                _sw.WriteLine(_strLine + "0A" + "0013204");
                            }
                            else
                            {
                                _sw = new StreamWriter(path + "bwe7000", true, _encoding);
                                _sw.WriteLine(_strLine + "0A" + "0013201");
                            }
                        }
                        else if (_encoding.GetString(_byteAry, 299, 1) == "3")
                        {
                            _sw = new StreamWriter(path + "bc_오류", true, _encoding);
                            _sw.WriteLine(_strLine + "0A" + "0011101");
                        }
                        else if (_encoding.GetString(_byteAry, 298, 1) == "0")
                        {
                            _sw = new StreamWriter(path + "bc10000_농협면세유", true, _encoding);
                            _sw.WriteLine(_strLine + "0A" + "0011105");
                        }
                        else if (_encoding.GetString(_byteAry, 298, 1) == "1")
                        {
                            if (_encoding.GetString(_byteAry, 244, 1) == "1")
                            {
                                _sw = new StreamWriter(path + "bwe19000", true, _encoding);
                                _sw.WriteLine(_strLine + "0A" + "0013202");
                            }
                            else
                            {
                                if (_encoding.GetString(_byteAry, 245, 1) == "3")
                                {
                                    _sw = new StreamWriter(path + "bwe9000_본인", true, _encoding);
                                    _sw.WriteLine(_strLine + "0A" + "0013204");
                                }
                                else
                                {
                                    _sw = new StreamWriter(path + "bwe7000", true, _encoding);
                                    _sw.WriteLine(_strLine + "0A" + "0013201");
                                }
                            }
                        }
                        else if (_encoding.GetString(_byteAry, 298, 1) == "2")
                        {
                            if (_encoding.GetString(_byteAry, 245, 1) == "3")
                            {
                                _sw = new StreamWriter(path + "bbe9000_본인", true, _encoding);
                                _sw.WriteLine(_strLine + "0A" + "0013303");
                            }
                            else
                            {
                                _sw = new StreamWriter(path + "bbe7500", true, _encoding);
                                _sw.WriteLine(_strLine + "0A" + "0013301");
                            }
                        }
                        else if (_encoding.GetString(_byteAry, 245, 1) == "1")
                        {
                            _sw = new StreamWriter(path + "bc500_기업세이브", true, _encoding);
                            _sw.WriteLine(_strLine + "0A" + "0011104");
                        }
                        else if (_encoding.GetString(_byteAry, 245, 1) == "3")
                        {
                            _sw = new StreamWriter(path + "bc9000_일반본인", true, _encoding);
                            _sw.WriteLine(_strLine + "0A" + "0011109");
                        }
                        else
                        {
                            _sw = new StreamWriter(path + "bc100_일반", true, _encoding);
                            _sw.WriteLine(_strLine + "0A" + "0011101");
                        }
                    }

                    _sw.Close();
                }

                _drs = _dtable.Select("", "");
                for (int i = 0; i < _drs.Length; i++)
                {
                    _sw.WriteLine(_strLine);
                }
                _strReturn = "성공";
            }
            catch (Exception ex)
            {
                _strReturn = string.Format("{0}번째 우편번호 오류", _iCount);
                if (_swError != null) _swError.WriteLine(ex.Message);
            }
            finally
            {
                if (_sr != null) _sr.Close();
                if (_sw != null) _sw.Close();
                if (_swError != null) _swError.Close();
            }
            return _strReturn;
        }

        //은행사별 분리
        private static string Convert_Bank_ID(string _strBankID)
        {
            string strBankName = "";
            
            switch (_strBankID)
            {
                case "04":
                case "06":
                case "09":
                case "19":
                case "29":
                case "87":
                    strBankName = "국민";
                    break;
                case "10":
                case "11":
                case "12":
                case "13":
                case "14":
                case "15":
                case "16":
                case "17":
                case "18":
                    strBankName = "농협";
                    break;
                case "20":
                case "22":
                case "24":
                case "83":
                case "84":
                case "90":
                    strBankName = "우리";
                    break;
                case "33":
                case "80":
                case "81":
                case "82":
                    strBankName = "하나";
                    break;
                case "23":
                    strBankName = "SC제일";
                    break;
                case "03":
                    strBankName = "기업";
                    break;
                case "21":
                case "26":
                case "38":
                case "40":
                    strBankName = "신한";
                    break;
                case "39":
                    strBankName = "경남";
                    break;
                case "31":
                    strBankName = "대구";
                    break;
                case "32":
                    strBankName = "부산";
                    break;
                case "34":
                    strBankName = "광주";
                    break;
                case "35":
                    strBankName = "제주";
                    break;
                case "37":
                    strBankName = "전북";
                    break;
                case "50":
                    strBankName = "바로";
                    break;
                default:
                    strBankName = "그외";
                    break;
            }

            return strBankName;
        }

        //은행사별 분리
        private static string Convert_Bank_Code(string strBank_Chk)
        {
            string strBankName = "";

            switch (_strBankID)
            {
                case "04":
                case "06":
                case "09":
                case "19":
                case "29":
                case "87":
                    strBankName = "국민";
                    break;
                case "10":
                case "11":
                case "12":
                case "13":
                case "14":
                case "15":
                case "16":
                case "17":
                case "18":
                    strBankName = "농협";
                    break;
                case "20":
                case "22":
                case "24":
                case "83":
                case "84":
                case "90":
                    strBankName = "우리";
                    break;
                case "33":
                case "80":
                case "81":
                case "82":
                    strBankName = "하나";
                    break;
                case "23":
                    strBankName = "SC제일";
                    break;
                case "03":
                    strBankName = "기업";
                    break;
                case "21":
                case "26":
                case "38":
                case "40":
                    strBankName = "신한";
                    break;
                case "39":
                    strBankName = "경남";
                    break;
                case "31":
                    strBankName = "대구";
                    break;
                case "32":
                    strBankName = "부산";
                    break;
                case "34":
                    strBankName = "광주";
                    break;
                case "35":
                    strBankName = "제주";
                    break;
                case "37":
                    strBankName = "전북";
                    break;
                case "50":
                    strBankName = "바로";
                    break;
                default:
                    strBankName = "그외";
                    break;
            }

            return strBankName;
        }

        private static StreamWriter GetStreamWriter(string path)
        {
            StreamWriter _return = null;
            FileInfo _fi = new FileInfo(path);
            if (_fi.Exists)
            {
                _return = new StreamWriter(path, false, System.Text.Encoding.GetEncoding(949));
            }
            else
            {
                _return = new StreamWriter(path, true, System.Text.Encoding.GetEncoding(949));
            }
            return _return;
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
    }
}
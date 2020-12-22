using System;
using System.Collections.Generic;
using System.Collections;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace _002_국민_CONVERT
{
    public class CONVERT
    {
        //기본 인코딩 설정
        private static string strEncoding = "ks_c_5601-1987";
        private static string strCardTypeID = "002_CONV";
        private static string strCardTypeName = "국민컨버터";

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
            StreamReader _sr = null;																					//파일 읽기 스트림
            StreamWriter _swError = null;
            StreamWriter _sw = null;
            byte[] _byteAry = null;
            /// _strCode : 제휴코드
            /// _strCode1 : 일일 : 1, 추가(갱신) : 2, 정기(갱신) : 3
            /// _strCode2 : 인편
            /// 일반 : I , 동의서 : D, 우편지역동의서 : S, 긴급 : Q, 특송 : P, 프리미엄일반 : N, 프리미엄동의서 : Y
            /// _strCode2 : 반송
            /// 일반 : B , 동의서 : A, 우편지역동의서 : T, 긴급 : G, 특송 : H, 프리미엄일반 : W, 프리미엄동의서 : X
            /// 일반법인 : E, F, 광역법인 : L,M
            string _strLine = "", _strCode = "", _strCode1 = null, _strCode2 = null, _strReturn = "", _strOwner_one = "", _strFamily_one = "", _strFamilyCode = "";
            string _strSendNumberPrev = "", _strSendNumber = "", _strOwner_onePrev = "", strIDchk = "", strOther_chk = "";

            // DataRow[] _drs = null;
            DataTable _dtable = null;
            DataSet _dsetZipcodeArea = null;
            try
            {
                _dtable = new DataTable("CONVERT");
                _dtable.Columns.Add("client_enterprise_code");


                _dsetZipcodeArea = new DataSet();
                _dsetZipcodeArea.ReadXml(xmlZipcodeAreaPath);


                //파일 읽기 Stream과 오류시 저장할 쓰기 Stream 생성
                _sr = new StreamReader(path, _encoding);
                _swError = new StreamWriter(path + ".Error", false, _encoding);


                while ((_strLine = _sr.ReadLine()) != null)
                {
                    _byteAry = _encoding.GetBytes(_strLine);
                    //본인 + 가족카드 구분
                    _strSendNumber = _encoding.GetString(_byteAry, 329, 16);
                    _strCode = _encoding.GetString(_byteAry, 307, 5);
                    _strCode1 = _encoding.GetString(_byteAry, 0, 1);
                    //동의서 구분 : 동의서 = D
                    _strCode2 = _encoding.GetString(_byteAry, 1, 1);

                    //제3자수령가능여부 1 = 동의, 0 = 미동의
                    //0 = 본인만배송
                    _strOwner_one = _encoding.GetString(_byteAry, 398, 1);
                    //가족앞필수교부 : 1 = 동의, 0 = 미동의
                    //1 = 본인만배송
                    _strFamily_one = _encoding.GetString(_byteAry, 399, 1);
                    //가족만발급 (0:본인만1매or본인+가족, 1:가족만1매, 2:가족만2매이상, 3:동시교부)
                    _strFamilyCode = _encoding.GetString(_byteAry, 345, 1);
                    //신분증진위사후확인
                    strIDchk = _encoding.GetString(_byteAry, 426, 1);
                    //대외기관체크코드
                    strOther_chk = _encoding.GetString(_byteAry, 431, 2);

                    //본인 + 가족카드의 경우
                    if (_strSendNumber == _strSendNumberPrev)
                    {
                        _strOwner_one = _strOwner_onePrev;
                    }
                    else
                    {
                        _strSendNumberPrev = _strSendNumber;
                        _strOwner_onePrev = _strOwner_one;
                    }

                    if (_strCode1 == "3")
                    {
                        switch (_strCode2)
                        {
                            case "A":
                                if (strIDchk == "1")
                                {
                                    _sw = new StreamWriter(path + ".21500_재발송대상_정기", true, _encoding);
                                    _sw.WriteLine(_strLine + "0022202");
                                }
                                else
                                {
                                    _sw = new StreamWriter(path + ".20500_재발송동의_정기", true, _encoding);
                                    _sw.WriteLine(_strLine + "0022201");
                                }
                                break;
                            case "B":
                                if (((_strFamilyCode == "0" || _strFamilyCode == "3") && _strOwner_one == "0") || (_strFamilyCode == "1" && _strFamily_one == "1"))
                                {   
                                    if (strOther_chk == "01")
                                    {
                                        _sw = new StreamWriter(path + ".35000_KB증권_재발송(본인)_정기", true, _encoding);
                                        _sw.WriteLine(_strLine + "0021304");
                                    }
                                    else if (strOther_chk == "02")
                                    {
                                        _sw = new StreamWriter(path + ".41000_페이코_재발송(본인)_정기", true, _encoding);
                                        _sw.WriteLine(_strLine + "0021404");
                                    }
                                    else if (strOther_chk == "03")
                                    {
                                        _sw = new StreamWriter(path + ".47000_한화투자증권_재발송(본인)_정기", true, _encoding);
                                        _sw.WriteLine(_strLine + "0021504");
                                    }
                                    else if (strOther_chk == "04")
                                    {
                                        _sw = new StreamWriter(path + ".53000_트레블월넷_재발송(본인)_정기", true, _encoding);
                                        _sw.WriteLine(_strLine + "0021604");
                                    }
                                    else
                                    {
                                        _sw = new StreamWriter(path + ".20000_재발송(본인)_정기", true, _encoding);
                                        _sw.WriteLine(_strLine + "0021202");
                                    }
                                }
                                else
                                {
                                    if (strOther_chk == "01")
                                    {
                                        _sw = new StreamWriter(path + ".34000_KB증권_재발송_정기", true, _encoding);
                                        _sw.WriteLine(_strLine + "0021303");
                                    }
                                    else if (strOther_chk == "02")
                                    {
                                        _sw = new StreamWriter(path + ".40000_페이코_재발송_정기", true, _encoding);
                                        _sw.WriteLine(_strLine + "0021403");
                                    }
                                    else if (strOther_chk == "03")
                                    {
                                        _sw = new StreamWriter(path + ".46000_한화투자증권_재발송_정기", true, _encoding);
                                        _sw.WriteLine(_strLine + "0021503");
                                    }
                                    else if (strOther_chk == "04")
                                    {
                                        _sw = new StreamWriter(path + ".52000_트레블월넷_재발송_정기", true, _encoding);
                                        _sw.WriteLine(_strLine + "0021603");
                                    }
                                    else
                                    {
                                        _sw = new StreamWriter(path + ".20000_정기", true, _encoding);
                                        _sw.WriteLine(_strLine + "0021201");
                                    }
                                }
                                break;
                            case "D":
                                if (strIDchk == "1")
                                {
                                    _sw = new StreamWriter(path + ".14000_갱신대상건_정기", true, _encoding);
                                    _sw.WriteLine(_strLine + "0022112");    
                                }
                                else
                                {
                                    _sw = new StreamWriter(path + ".2500_갱신_정기", true, _encoding);
                                    _sw.WriteLine(_strLine + "0022110");
                                }
                                break;
                            case "G":
                                if (((_strFamilyCode == "0" || _strFamilyCode == "3") && _strOwner_one == "0") || (_strFamilyCode == "1" && _strFamily_one == "1"))
                                {
                                    _sw = new StreamWriter(path + ".21000_오후(재발급)_본인_정기", true, _encoding);
                                    _sw.WriteLine(_strLine + "0023106");
                                }
                                else
                                {
                                    _sw = new StreamWriter(path + ".21000_정기", true, _encoding);
                                    _sw.WriteLine(_strLine + "0023103");
                                }
                                break;
                            case "I":   //갱신
                                if (strOther_chk == "01" && _strOwner_one == "0")
                                {
                                    _sw = new StreamWriter(path + ".37000_KB증권_갱신(본인)_정기", true, _encoding);
                                    _sw.WriteLine(_strLine + "0021306");
                                }
                                else if (strOther_chk == "02" && _strOwner_one == "0")
                                {
                                    _sw = new StreamWriter(path + ".43000_페이코_갱신(본인)_정기", true, _encoding);
                                    _sw.WriteLine(_strLine + "0021406");
                                }
                                else if (strOther_chk == "03" && _strOwner_one == "0")
                                {
                                    _sw = new StreamWriter(path + ".49000_한화투자증권_갱신(본인)_정기", true, _encoding);
                                    _sw.WriteLine(_strLine + "0021506");
                                }
                                else if (strOther_chk == "04" && _strOwner_one == "0")
                                {
                                    _sw = new StreamWriter(path + ".55000_트레블월넷_갱신(본인)_정기", true, _encoding);
                                    _sw.WriteLine(_strLine + "0021606");
                                }
                                else if (((_strFamilyCode == "0" || _strFamilyCode == "3") && _strOwner_one == "0") || (_strFamilyCode == "1" && _strFamily_one == "1"))
                                {   
                                    _sw = new StreamWriter(path + ".2000_갱신(본인)_정기", true, _encoding);
                                    _sw.WriteLine(_strLine + "0021105");
                                }
                                else
                                {
                                    if (strOther_chk == "01")
                                    {
                                        _sw = new StreamWriter(path + ".36000_KB증권_갱신_정기", true, _encoding);
                                        _sw.WriteLine(_strLine + "0021305");
                                    }
                                    else if (strOther_chk == "02")
                                    {
                                        _sw = new StreamWriter(path + ".42000_페이코_갱신_정기", true, _encoding);
                                        _sw.WriteLine(_strLine + "0021405");
                                    }
                                    else if (strOther_chk == "03")
                                    {
                                        _sw = new StreamWriter(path + ".48000_한화투자증권_갱신_정기", true, _encoding);
                                        _sw.WriteLine(_strLine + "0021505");
                                    }
                                    else if (strOther_chk == "04")
                                    {
                                        _sw = new StreamWriter(path + ".54000_트레블월넷_갱신_정기", true, _encoding);
                                        _sw.WriteLine(_strLine + "0021605");
                                    }
                                    else
                                    {
                                        _sw = new StreamWriter(path + ".2000_정기", true, _encoding);
                                        _sw.WriteLine(_strLine + "0021102");
                                    }
                                }
                                break;
                            case "Q":   //긴급오후
                                if (((_strFamilyCode == "0" || _strFamilyCode == "3") && _strOwner_one == "0") || (_strFamilyCode == "1" && _strFamily_one == "1"))
                                {
                                    _sw = new StreamWriter(path + ".1000_오후(본인)_정기", true, _encoding);
                                    _sw.WriteLine(_strLine + "0023105");
                                }
                                else
                                {
                                    _sw = new StreamWriter(path + ".1000_정기", true, _encoding);
                                    _sw.WriteLine(_strLine + "0023102");
                                }
                                break;
                            case "N":   //갱신VVIP 
                                if (((_strFamilyCode == "0" || _strFamilyCode == "3") && _strOwner_one == "0") || (_strFamilyCode == "1" && _strFamily_one == "1"))
                                {
                                    _sw = new StreamWriter(path + ".17000_광역갱신(본인)_정기", true, _encoding);
                                    _sw.WriteLine(_strLine + "0023204");
                                }
                                else
                                {
                                    _sw = new StreamWriter(path + ".17000_정기", true, _encoding);
                                    _sw.WriteLine(_strLine + "0023202");
                                }
                                break;
                            case "Y":   //광역VVIP
                                if (strIDchk=="1")
                                {
                                    _sw = new StreamWriter(path + ".9500_VVIP대상건_정기", true, _encoding);
                                    _sw.WriteLine(_strLine + "0022304");
                                }
                                else
                                {
                                    _sw = new StreamWriter(path + ".7500_광역VVIP_정기", true, _encoding);
                                    _sw.WriteLine(_strLine + "0022302");
                                }
                                break;
                            case "S":   //광역
                                if (strIDchk == "1")
                                {
                                    _sw = new StreamWriter(path + ".8500_광역대상건_정기", true, _encoding);
                                    _sw.WriteLine(_strLine + "0022303");
                                }
                                else
                                {
                                    _sw = new StreamWriter(path + ".61500_광역동의_정기", true, _encoding);
                                    _sw.WriteLine(_strLine + "0022301");
                                }
                                break;
                            case "E":    
                            case "F":   //법인 
                                _sw = new StreamWriter(path + ".7000_법인(본인)_정기", true, _encoding);
                                _sw.WriteLine(_strLine + "0021107");
                                break;
                            case "L":
                            case "M":   //법인 
                                _sw = new StreamWriter(path + ".9000_광역_법인(본인)_정기", true, _encoding);
                                _sw.WriteLine(_strLine + "0023205");
                                break;
                            default:
                                _sw = new StreamWriter(path + ".기타_DONG_정기", true, _encoding);
                                _sw.WriteLine(_strLine);
                                break;
                        }
                    }
                    else if (_strCode1 == "2")
                    {
                        switch (_strCode2)
                        {
                            case "A":
                                if (strIDchk == "1")
                                {
                                    _sw = new StreamWriter(path + ".21500_재발송대상_추가", true, _encoding);
                                    _sw.WriteLine(_strLine + "0022202");
                                }
                                else
                                {
                                    _sw = new StreamWriter(path + ".20500_재발송동의_추가", true, _encoding);
                                    _sw.WriteLine(_strLine + "0022201");
                                }
                                break;
                            case "B":
                                if (strOther_chk == "01" && _strOwner_one == "0")
                                {
                                    _sw = new StreamWriter(path + ".35000_KB증권_재발송(본인)_추가", true, _encoding);
                                    _sw.WriteLine(_strLine + "0021304");
                                }
                                else if (strOther_chk == "02" && _strOwner_one == "0")
                                {
                                    _sw = new StreamWriter(path + ".41000_페이코_재발송(본인)_추가", true, _encoding);
                                    _sw.WriteLine(_strLine + "0021404");
                                }
                                else if (strOther_chk == "03" && _strOwner_one == "0")
                                {
                                    _sw = new StreamWriter(path + ".47000_한화투자증권_재발송(본인)_추가", true, _encoding);
                                    _sw.WriteLine(_strLine + "0021504");
                                }
                                else if (strOther_chk == "04" && _strOwner_one == "0")
                                {
                                    _sw = new StreamWriter(path + ".53000_트레블월넷_재발송(본인)_추가", true, _encoding);
                                    _sw.WriteLine(_strLine + "0021604");
                                }
                                else if (((_strFamilyCode == "0" || _strFamilyCode == "3") && _strOwner_one == "0") || (_strFamilyCode == "1" && _strFamily_one == "1"))
                                {
                                    _sw = new StreamWriter(path + ".20000_재발송(본인)추가", true, _encoding);
                                    _sw.WriteLine(_strLine + "0021202");
                                }
                                else
                                {
                                    if (strOther_chk == "01")
                                    {
                                        _sw = new StreamWriter(path + ".34000_KB증권_재발송_추가", true, _encoding);
                                        _sw.WriteLine(_strLine + "0021303");
                                    }
                                    else if (strOther_chk == "02")
                                    {
                                        _sw = new StreamWriter(path + ".40000_페이코_재발송_추가", true, _encoding);
                                        _sw.WriteLine(_strLine + "0021403");
                                    }
                                    else if (strOther_chk == "03")
                                    {
                                        _sw = new StreamWriter(path + ".46000_한화투자증권_재발송_추가", true, _encoding);
                                        _sw.WriteLine(_strLine + "0021503");
                                    }
                                    else if (strOther_chk == "04")
                                    {
                                        _sw = new StreamWriter(path + ".42000_트레블월넷_재발송_추가", true, _encoding);
                                        _sw.WriteLine(_strLine + "0021603");
                                    }
                                    else
                                    {
                                        _sw = new StreamWriter(path + ".20000_추가", true, _encoding);
                                        _sw.WriteLine(_strLine + "0021201");
                                    }
                                }
                                break;
                            case "D":
                                if (strIDchk == "1")
                                {
                                    _sw = new StreamWriter(path + ".14000_갱신대상_추가", true, _encoding);
                                    _sw.WriteLine(_strLine + "0022112");
                                }
                                else
                                {
                                    _sw = new StreamWriter(path + ".2500_갱신_추가", true, _encoding);
                                    _sw.WriteLine(_strLine + "0022110");
                                }
                                break;
                            case "G":
                                if (((_strFamilyCode == "0" || _strFamilyCode == "3") && _strOwner_one == "0") || (_strFamilyCode == "1" && _strFamily_one == "1"))
                                {
                                    _sw = new StreamWriter(path + ".21000_오후(재발급)_본인_추가", true, _encoding);
                                    _sw.WriteLine(_strLine + "0023106");
                                }
                                else
                                {
                                    _sw = new StreamWriter(path + ".21000_추가", true, _encoding);
                                    _sw.WriteLine(_strLine + "0023103");
                                }
                                break;
                            case "I":   //갱신
                                if (strOther_chk == "01" && _strOwner_one == "0")
                                {
                                    _sw = new StreamWriter(path + ".37000_KB증권_갱신(본인)_추가", true, _encoding);
                                    _sw.WriteLine(_strLine + "0021306");
                                }
                                else if (strOther_chk == "02" && _strOwner_one == "0")
                                {
                                    _sw = new StreamWriter(path + ".43000_페이코_갱신(본인)_추가", true, _encoding);
                                    _sw.WriteLine(_strLine + "0021406");
                                }
                                else if (strOther_chk == "03" && _strOwner_one == "0")
                                {
                                    _sw = new StreamWriter(path + ".49000_한화투자증권_갱신(본인)_추가", true, _encoding);
                                    _sw.WriteLine(_strLine + "0021506");
                                }
                                else if (strOther_chk == "04" && _strOwner_one == "0")
                                {
                                    _sw = new StreamWriter(path + ".55000_트레블월넷_갱신(본인)_추가", true, _encoding);
                                    _sw.WriteLine(_strLine + "0021606");
                                }
                                else if (((_strFamilyCode == "0" || _strFamilyCode == "3") && _strOwner_one == "0") || (_strFamilyCode == "1" && _strFamily_one == "1"))
                                {
                                    _sw = new StreamWriter(path + ".2000_갱신(본인)_추가", true, _encoding);
                                    _sw.WriteLine(_strLine + "0021105");
                                }
                                else
                                {
                                    if (strOther_chk == "01")
                                    {
                                        _sw = new StreamWriter(path + ".36000_KB증권_갱신_추가", true, _encoding);
                                        _sw.WriteLine(_strLine + "0021305");
                                    }
                                    else if (strOther_chk == "02")
                                    {
                                        _sw = new StreamWriter(path + ".42000_페이코_갱신_추가", true, _encoding);
                                        _sw.WriteLine(_strLine + "0021405");
                                    }
                                    else if (strOther_chk == "03")
                                    {
                                        _sw = new StreamWriter(path + ".48000_한화투자증권_갱신_추가", true, _encoding);
                                        _sw.WriteLine(_strLine + "0021505");
                                    }
                                    else if (strOther_chk == "04")
                                    {
                                        _sw = new StreamWriter(path + ".54000_트레블월넷_갱신_추가", true, _encoding);
                                        _sw.WriteLine(_strLine + "0021605");
                                    }
                                    else
                                    {
                                        _sw = new StreamWriter(path + ".2000_추가", true, _encoding);
                                        _sw.WriteLine(_strLine + "0021102");
                                    }
                                }
                                break;
                            case "Q":   //긴급오후
                                if (((_strFamilyCode == "0" || _strFamilyCode == "3") && _strOwner_one == "0") || (_strFamilyCode == "1" && _strFamily_one == "1"))
                                {
                                    _sw = new StreamWriter(path + ".1000_오후(본인)_추가", true, _encoding);
                                    _sw.WriteLine(_strLine + "0023105");
                                }
                                else
                                {
                                    _sw = new StreamWriter(path + ".1000_추가", true, _encoding);
                                    _sw.WriteLine(_strLine + "0023102");
                                }
                                break;
                            case "N":   // VVIP
                                if (((_strFamilyCode == "0" || _strFamilyCode == "3") && _strOwner_one == "0") || (_strFamilyCode == "1" && _strFamily_one == "1"))
                                {
                                    _sw = new StreamWriter(path + ".7000_VVIP(본인)_추가", true, _encoding);
                                    _sw.WriteLine(_strLine + "0023203");
                                }
                                else
                                {
                                    _sw = new StreamWriter(path + ".7000_추가", true, _encoding);
                                    _sw.WriteLine(_strLine + "0023201");
                                }
                                break;
                            case "Y":   // 광역VVIP
                                if (strIDchk == "1")
                                {
                                    _sw = new StreamWriter(path + ".9500_VVIP대상건_추가", true, _encoding);
                                    _sw.WriteLine(_strLine + "0022304");
                                }
                                else
                                {
                                    _sw = new StreamWriter(path + ".7500_광역VVIP_추가", true, _encoding);
                                    _sw.WriteLine(_strLine + "0022302");
                                }
                                break;
                            case "S":   // 광역
                                if (strIDchk == "1")
                                {
                                    _sw = new StreamWriter(path + ".8500_광역대상건_추가", true, _encoding);
                                    _sw.WriteLine(_strLine + "0022303");   
                                }
                                else
                                {
                                    _sw = new StreamWriter(path + ".61500_광역동의_추가", true, _encoding);
                                    _sw.WriteLine(_strLine + "0022301");
                                }
                                break;
                            case "E":    
                            case "F":   //법인 
                                _sw = new StreamWriter(path + ".7000_법인(본인)_추가", true, _encoding);
                                _sw.WriteLine(_strLine + "0021107");
                                break;
                            case "L":
                            case "M":   //법인 
                                _sw = new StreamWriter(path + ".9000_광역_법인(본인)_추가", true, _encoding);
                                _sw.WriteLine(_strLine + "0023205");
                                break;
                            default:
                                _sw = new StreamWriter(path + ".기타_DONG_추가", true, _encoding);
                                _sw.WriteLine(_strLine);
                                break;
                        }
                    }
                    else
                    {
                        switch (_strCode2)
                        {
                            case "A":
                                if (strIDchk == "1")
                                {
                                    _sw = new StreamWriter(path + ".21500_재발송대상_일일", true, _encoding);
                                    _sw.WriteLine(_strLine + "0022202");
                                }
                                else
                                {
                                    _sw = new StreamWriter(path + ".20500_재발송동의_일일", true, _encoding);
                                    _sw.WriteLine(_strLine + "0022201");
                                }
                                break;
                            case "B":
                                if (strOther_chk == "01" && _strOwner_one == "0")
                                {
                                    _sw = new StreamWriter(path + ".35000_KB증권_재발송(본인)_일일", true, _encoding);
                                    _sw.WriteLine(_strLine + "0021304");
                                }
                                else if (strOther_chk == "02" && _strOwner_one == "0")
                                {
                                    _sw = new StreamWriter(path + ".41000_페이코_재발송(본인)_일일", true, _encoding);
                                    _sw.WriteLine(_strLine + "0021404");
                                }
                                else if (strOther_chk == "03" && _strOwner_one == "0")
                                {
                                    _sw = new StreamWriter(path + ".47000_한화투자증권_재발송(본인)_일일", true, _encoding);
                                    _sw.WriteLine(_strLine + "0021504");
                                }
                                else if (strOther_chk == "04" && _strOwner_one == "0")
                                {
                                    _sw = new StreamWriter(path + ".53000_트레블월넷_재발송(본인)_일일", true, _encoding);
                                    _sw.WriteLine(_strLine + "0021604");
                                }
                                else if (((_strFamilyCode == "0" || _strFamilyCode == "3") && _strOwner_one == "0") || (_strFamilyCode == "1" && _strFamily_one == "1"))
                                {
                                    _sw = new StreamWriter(path + ".20000_재발송(본인)일일", true, _encoding);
                                    _sw.WriteLine(_strLine + "0021202");
                                }
                                else
                                {
                                    if (strOther_chk == "01")
                                    {
                                        _sw = new StreamWriter(path + ".34000_KB증권_재발송_일일", true, _encoding);
                                        _sw.WriteLine(_strLine + "0021303");
                                    }
                                    else if (strOther_chk == "02")
                                    {
                                        _sw = new StreamWriter(path + ".40000_페이코_재발송_일일", true, _encoding);
                                        _sw.WriteLine(_strLine + "0021403");
                                    }
                                    else if (strOther_chk == "03")
                                    {
                                        _sw = new StreamWriter(path + ".46000_한화투자증권_재발송_일일", true, _encoding);
                                        _sw.WriteLine(_strLine + "0021503");
                                    }
                                    else if (strOther_chk == "04")
                                    {
                                        _sw = new StreamWriter(path + ".52000_트레블월넷_재발송_일일", true, _encoding);
                                        _sw.WriteLine(_strLine + "0021603");
                                    }
                                    else
                                    {
                                        _sw = new StreamWriter(path + ".20000_일일", true, _encoding);
                                        _sw.WriteLine(_strLine + "0021201");
                                    }
                                }
                                break;
                            case "D":
                                if (strIDchk == "1")
                                {
                                    _sw = new StreamWriter(path + ".1000_동의대상건", true, _encoding);
                                    _sw.WriteLine(_strLine + "0022111");   
                                }
                                else
                                {
                                    _sw = new StreamWriter(path + ".500_일일", true, _encoding);
                                    _sw.WriteLine(_strLine + "0022101");
                                }
                                break;
                            case "G":
                                if (((_strFamilyCode == "0" || _strFamilyCode == "3") && _strOwner_one == "0") || (_strFamilyCode == "1" && _strFamily_one == "1"))
                                {
                                    _sw = new StreamWriter(path + ".21000_오후(재발급)_본인_일일", true, _encoding);
                                    _sw.WriteLine(_strLine + "0023106");
                                }
                                else
                                {
                                    _sw = new StreamWriter(path + ".21000_일일", true, _encoding);
                                    _sw.WriteLine(_strLine + "0023103");
                                }
                                break;
                            case "I":   //일반 
                                if (strOther_chk == "01" && _strOwner_one == "0")
                                {
                                    _sw = new StreamWriter(path + ".33001_KB증권_일반(본인)_일일", true, _encoding);
                                    _sw.WriteLine(_strLine + "0021302");
                                }
                                else if (strOther_chk == "02" && _strOwner_one == "0")
                                {
                                    _sw = new StreamWriter(path + ".39001_페이코_일반(본인)_일일", true, _encoding);
                                    _sw.WriteLine(_strLine + "0021402");
                                }
                                else if (strOther_chk == "03" && _strOwner_one == "0")
                                {
                                    _sw = new StreamWriter(path + ".45001_한화투자증권_일반(본인)_일일", true, _encoding);
                                    _sw.WriteLine(_strLine + "0021502");
                                }
                                else if (strOther_chk == "04" && _strOwner_one == "0")
                                {
                                    _sw = new StreamWriter(path + ".51001_트레블월넷_일반(본인)_일일", true, _encoding);
                                    _sw.WriteLine(_strLine + "0021602");
                                }
                                else if (((_strFamilyCode == "0" || _strFamilyCode == "3") && _strOwner_one == "0") || (_strFamilyCode == "1" && _strFamily_one == "1"))
                                {
                                    _sw = new StreamWriter(path + ".100_일반(본인)_일일", true, _encoding);
                                    _sw.WriteLine(_strLine + "0021104");
                                }
                                else
                                {
                                    if (strOther_chk == "01")
                                    {
                                        _sw = new StreamWriter(path + ".32000_KB증권_일일", true, _encoding);
                                        _sw.WriteLine(_strLine + "0021301");
                                    }
                                    else if (strOther_chk == "02")
                                    {
                                        _sw = new StreamWriter(path + ".38000_페이코_일일", true, _encoding);
                                        _sw.WriteLine(_strLine + "0021401");
                                    }
                                    else if (strOther_chk == "03")
                                    {
                                        _sw = new StreamWriter(path + ".44000_한화투자증권_일일", true, _encoding);
                                        _sw.WriteLine(_strLine + "0021501");
                                    }
                                    else if (strOther_chk == "04")
                                    {
                                        _sw = new StreamWriter(path + ".50000_트레블월넷_일일", true, _encoding);
                                        _sw.WriteLine(_strLine + "0021601");
                                    }
                                    else
                                    {
                                        _sw = new StreamWriter(path + ".100_일일", true, _encoding);
                                        _sw.WriteLine(_strLine + "0021101");
                                    }
                                }

                                break;
                            case "Q":   //긴급오후
                                if (((_strFamilyCode == "0" || _strFamilyCode == "3") && _strOwner_one == "0") || (_strFamilyCode == "1" && _strFamily_one == "1"))
                                {
                                    _sw = new StreamWriter(path + ".1000_오후(본인)_일일", true, _encoding);
                                    _sw.WriteLine(_strLine + "0023105");
                                }
                                else
                                {
                                    _sw = new StreamWriter(path + ".1000_일일", true, _encoding);
                                    _sw.WriteLine(_strLine + "0023102");
                                }
                                break;
                            case "N":   // VVIP
                                if (ConvertStrCode(_strCode) == "미르")
                                {
                                    if (((_strFamilyCode == "0" || _strFamilyCode == "3") && _strOwner_one == "0") || (_strFamilyCode == "1" && _strFamily_one == "1"))
                                    {
                                        _sw = new StreamWriter(path + ".30000_VV미르(본인)_일일", true, _encoding);
                                        _sw.WriteLine(_strLine + "0023402");
                                    }
                                    else
                                    {
                                        _sw = new StreamWriter(path + ".30000_일일", true, _encoding);
                                        _sw.WriteLine(_strLine + "0023401");
                                    }
                                }
                                else
                                {
                                    if (((_strFamilyCode == "0" || _strFamilyCode == "3") && _strOwner_one == "0") || (_strFamilyCode == "1" && _strFamily_one == "1"))
                                    {
                                        _sw = new StreamWriter(path + ".7000_VVIP(본인)_일일", true, _encoding);
                                        _sw.WriteLine(_strLine + "0023203");
                                    }
                                    else
                                    {
                                        _sw = new StreamWriter(path + ".7000_일일", true, _encoding);
                                        _sw.WriteLine(_strLine + "0023201");
                                    }
                                }
                                break;
                            case "Y":   // 광역VVIP-동의
                                if (ConvertStrCode(_strCode) == "미르")
                                {
                                    if (strIDchk == "1")
                                    {
                                        _sw = new StreamWriter(path + ".31500_미르대상건_일일", true, _encoding);
                                        _sw.WriteLine(_strLine + "0022502");
                                    }
                                    else
                                    {
                                        _sw = new StreamWriter(path + ".30500_미르동의_일일", true, _encoding);
                                        _sw.WriteLine(_strLine + "0022501");
                                    }
                                }
                                else
                                {
                                    if (strIDchk == "1")
                                    {
                                        _sw = new StreamWriter(path + ".9500_VVIP대상건_일일", true, _encoding);
                                        _sw.WriteLine(_strLine + "0022304");
                                    }
                                    else
                                    {
                                        _sw = new StreamWriter(path + ".7500_광역VVIP_일일", true, _encoding);
                                        _sw.WriteLine(_strLine + "0022302");
                                    }
                                }
                                break;
                            case "S":   // 광역
                                if (strIDchk == "1")
                                {
                                    _sw = new StreamWriter(path + ".8500_광역대상건_일일", true, _encoding);
                                    _sw.WriteLine(_strLine + "0022303");
                                }
                                else
                                {
                                    _sw = new StreamWriter(path + ".61500_광역동의_일일", true, _encoding);
                                    _sw.WriteLine(_strLine + "0022301");
                                }
                                break;
                            case "E":
                            case "F":   //법인 
                                _sw = new StreamWriter(path + ".7000_법인(본인)_일일", true, _encoding);
                                _sw.WriteLine(_strLine + "0021107");
                                break;
                            case "L":
                            case "M":   //법인 
                                _sw = new StreamWriter(path + ".9000_광역_법인(본인)_일일", true, _encoding);
                                _sw.WriteLine(_strLine + "0023205");
                                break;
                            default:
                                _sw = new StreamWriter(path + ".기타_DONG_일일", true, _encoding);
                                _sw.WriteLine(_strLine);
                                break;
                        }
                    }

                    //_sw.WriteLine(_strLine);
                    _sw.Close();
                }
                _strReturn = "성공";
            }
            catch (Exception ex)
            {
                _strReturn = "실패";
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

        // 제휴동의서 코드 -> 동의서명으로 변환
        private static string ConvertStrCode(string _strCode)
        {
            string strName = null;

            switch (_strCode)
            {   
                case "09301":
                case "09302":
                case "09303":
                case "09304":
                case "09305":
                case "09306":
                    strName = "미르";
                    break;
            }
            return strName;
        }

        private static StreamWriter GetStreamWriter(string path)
        {
            StreamWriter _return = null;
            FileInfo _fi = new FileInfo(path);
            if (_fi.Exists)
            {
                _return = new StreamWriter(path, true, System.Text.Encoding.GetEncoding(949));
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
        private static string[] StringSplit(ref string data)
        {
            int index = data.IndexOf(":");
            string content = data.Substring(index + 1);

            data = data.Substring(0, index);

            string[] returnValue = content.Split(",".ToCharArray());
            return returnValue;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;


namespace BC_CONVERT_RESEND
{
    public class CONVERT
    {
        //기본 인코딩 설정
        private static string strEncoding = "ks_c_5601-1987";
        private static string strCardTypeID = "001_CONV";
        private static string strCardTypeName = "비씨_재배송_컨버트";

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
            StreamReader _sr = null;	//파일 읽기 스트림
            StreamWriter _swError = null;
            StreamWriter _sw = null;

            byte[] _byteAry = null;
            string _strBankID = "", _strZipcode = "", _strAreaGroup = "";
            //2013.06.26 태희철 _strValue : 영업점코드 : 2 = 영업점, 3 = 제3영업점, 4 = 강제영업점
            string _strReturn = "", _strLine = "", _strValue = null;
            DataTable _dtable = null;
            DataSet _dsetZipcodeArea = null;
            //DataRow _dr = null;
            DataRow[] _drs = null;
            int _iCount = 0;
            string _strDong = null, _strVIP = null, _strSC = null, _strYe = null, strBank_code = "", _strChk_add = "", strWoori_Mem = "";
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
                    //_strValue = _encoding.GetString(_byteAry, 540, 2);
                    //영업점코드 : 2 = 영업점, 3 = 제3영업점, 4 = 강제영업점
                    _strValue = _encoding.GetString(_byteAry, 145, 1);
                    _strDong = _encoding.GetString(_byteAry, 247, 1);
                    //2012-05-15 태희철 추가
                    _strVIP = _encoding.GetString(_byteAry, 296, 1);
                    //2012-05-15 태희철 추가
                    _strSC = _encoding.GetString(_byteAry, 297, 2);
                    //2012-06-04 태희철 추가 Ye치과 구분
                    _strYe = _encoding.GetString(_byteAry, 245, 1);
                    //은행구분코드
                    strBank_code = _encoding.GetString(_byteAry, 250, 3);
                    //우리은행(롯데멤버십) : 837011 / 837711
                    strWoori_Mem = _encoding.GetString(_byteAry, 542, 6);
                    //우리은행 통신사별지
                    _strChk_add = _encoding.GetString(_byteAry, 2765, 1);

                    _drs = _dsetZipcodeArea.Tables[0].Select("zipcode=" + _strZipcode);

                    _iCount++;


                    // [1] 동의서 1~8은 동의서 | 299,1 의 1 or 2 는 긴급은 따로 분류
                    //if (_byteAry.Length != 1739)
                    //if (_byteAry.Length != 1765)
                    //
                    if (_byteAry.Length != 2766)
                    {
                        _sw = new StreamWriter(path + "총byte오류_재발송", true, _encoding);
                        _sw.WriteLine(_strLine);
                    }
                    else
                    {
                        switch (_strDong)
                        {
                            case "1":
                            case "2":
                            case "3":
                            case "4":
                            case "5":
                            case "6":
                            case "7":
                            case "8":
                            case "9":
                                if (strBank_code == "020")
                                {
                                    if (_strDong == "1")
                                    {
                                        _sw = new StreamWriter(path + "bcd35000_우리재발송(롯데멤버스)", true, _encoding);
                                        _sw.WriteLine(_strLine + "0A" + "0012404");
                                    }
                                    else if (_strDong == "4")
                                    {
                                        _sw = new StreamWriter(path + "우리재발송(국기원)_동의_19000", true, _encoding);
                                        _sw.WriteLine(_strLine + "0A" + "0012402");
                                    }
                                    else
                                    {
                                        if (_strChk_add == "1" || _strChk_add == "2" || _strChk_add == "3")
                                        {
                                            _sw = new StreamWriter(path + "우리재발송(통신)_20000", true, _encoding);
                                            _sw.WriteLine(_strLine + "0A" + "0012403");
                                        }
                                        else
                                        {
                                            _sw = new StreamWriter(path + "우리재발송_동의_18000", true, _encoding);
                                            _sw.WriteLine(_strLine + "0A" + "0012401");
                                        }
                                    }
                                }
                                else if (strBank_code == "031")
                                {
                                    _sw = new StreamWriter(path + "대구재발송_동의_6500", true, _encoding);
                                    _sw.WriteLine(_strLine + "0A" + "0012201");
                                }
                                else if (strBank_code == "039")
                                {
                                    _sw = new StreamWriter(path + "경남재발송_동의_17000", true, _encoding);
                                    _sw.WriteLine(_strLine + "0A" + "0012301");
                                }
                                else if (strBank_code == "050")
                                {
                                    _sw = new StreamWriter(path + "바로재발송_동의_33000", true, _encoding);
                                    _sw.WriteLine(_strLine + "0A" + "0012501");
                                }
                                else
                                {
                                    _sw = new StreamWriter(path + "기타", true, _encoding);
                                    _sw.WriteLine(_strLine + "0A");
                                }
                                break;
                            default:
                                if (strBank_code == "020")
                                {
                                    if (_strYe == "1" || _strYe == "3")
                                    {
                                        _sw = new StreamWriter(path + "우리재발송_일반본인_12000", true, _encoding);
                                        _sw.WriteLine(_strLine + "0A" + "0011111");
                                    }
                                    else
                                    {
                                        _sw = new StreamWriter(path + "우리재발송_일반_8000", true, _encoding);
                                        _sw.WriteLine(_strLine + "0A" + "0011108");
                                    }
                                }
                                else if (strBank_code == "031")
                                {
                                    if (_strYe == "1" || _strYe == "3")
                                    {
                                        _sw = new StreamWriter(path + "대구재발송_일반본인_34000", true, _encoding);
                                        _sw.WriteLine(_strLine + "0A" + "0011402");
                                    }
                                    else
                                    {
                                        _sw = new StreamWriter(path + "대구재발송_일반_6000", true, _encoding);
                                        _sw.WriteLine(_strLine + "0A" + "0011401");
                                    }
                                    
                                }
                                else if (strBank_code == "039")
                                {
                                    if (_strYe == "1" || _strYe == "3")
                                    {
                                        _sw = new StreamWriter(path + "경남재발송_일반본인_11000", true, _encoding);
                                        _sw.WriteLine(_strLine + "0A" + "0011110");
                                    }
                                    else
                                    {
                                        _sw = new StreamWriter(path + "경남재발송_일반_7000", true, _encoding);
                                        _sw.WriteLine(_strLine + "0A" + "0011107");
                                    }
                                }
                                else if (strBank_code == "050")
                                {
                                    _sw = new StreamWriter(path + "바로재발송_일반_32000", true, _encoding);
                                    _sw.WriteLine(_strLine + "0A" + "0011114");
                                }
                                else if (strBank_code == "042")
                                {
                                    if (_strYe == "1" || _strYe == "3")
                                    {
                                        _sw = new StreamWriter(path + "토스재발송_본인_37000", true, _encoding);
                                        _sw.WriteLine(_strLine + "0A" + "0011116");
                                    }
                                    else
                                    {
                                        _sw = new StreamWriter(path + "토스재발송_일반_36000", true, _encoding);
                                        _sw.WriteLine(_strLine + "0A" + "0011115");
                                    }
                                }
                                else if (strBank_code == "079")
                                {
                                    if (_strYe == "1" || _strYe == "3")
                                    {
                                        _sw = new StreamWriter(path + "차이재발송_본인_39000", true, _encoding);
                                        _sw.WriteLine(_strLine + "0A" + "0011118");
                                    }
                                    else
                                    {
                                        _sw = new StreamWriter(path + "차이재발송_일반_38000", true, _encoding);
                                        _sw.WriteLine(_strLine + "0A" + "0011117");
                                    }
                                }
                                else
                                {
                                    _sw = new StreamWriter(path + "기타", true, _encoding);
                                    _sw.WriteLine(_strLine + "0A");
                                }
                                break;
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
                _strReturn = string.Format("{0}번째 우편번호 오류", _iCount + 1);
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
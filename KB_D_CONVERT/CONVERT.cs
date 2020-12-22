using System;
using System.Collections.Generic;
using System.Collections;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace KB_AM_EMER_CONVERT
{
    public class CONVERT
    {
        //기본 인코딩 설정
        private static string strEncoding = "ks_c_5601-1987";
        private static string strCardTypeID = "002_CONV";
        private static string strCardTypeName = "국민오전긴급_컨버터";

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
            /// 긴급오전동의 EmerDong : 326~328
            string _strLine = "", _strCode = "", _strCode1 = null, _strCode2 = null;
            string _strReturn = "", _strOwner_one = "", _strFamily_one = "", _strFamilyCode = "", strIDchk = "";
            string strCustomer_SSN = "", strCustomer_SSN_old = "", strData_Group = "";
            int iCard_cnt = 0;
            string EmerDong = "";

            bool bFamil_chk = false;

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
                    //_strCode = _encoding.GetString(_byteAry, 526, 5);
                    _strCode = _encoding.GetString(_byteAry, 307, 5);
                    _strCode1 = _encoding.GetString(_byteAry, 0, 1);
                    //동의서 구분 : 동의서 = D
                    _strCode2 = _encoding.GetString(_byteAry, 1, 1);

                    //제3자수령가능여부 1 = 동의, 0 = 미동의
                    _strOwner_one = _encoding.GetString(_byteAry, 398, 1);
                    //가족앞필수교부 : 1 = 동의, 0 = 미동의
                    _strFamily_one = _encoding.GetString(_byteAry, 399, 1);
                    //가족만발급 (0:본인만1매or본인+가족, 1:가족만1매, 2:가족만2매이상)
                    _strFamilyCode = _encoding.GetString(_byteAry, 345, 1);
                    //긴급동의
                    EmerDong = _encoding.GetString(_byteAry, 326, 2);
                    //신분증진위사후확인
                    strIDchk = _encoding.GetString(_byteAry, 426, 1);
                    //본인회원 신분증 번호
                    strCustomer_SSN = _encoding.GetString(_byteAry, 358, 13);
                    ///가족카드 구분
                    ///strCard_cnt > 1, _strFamilyCode = 0,
                    ///strCustomer_SSN : 본인회원 주민번호 동일
                    ///가족 묶음 데이터 수량은 strCard_cnt 동일

                    string strCard_cnt_chk = _encoding.GetString(_byteAry, 304, 1);

                    
                    if (bFamil_chk == true && strCustomer_SSN == strCustomer_SSN_old)
                    {
                        bFamil_chk = true;
                    }
                    else
                    {
                        bFamil_chk = false;
                        iCard_cnt = 0;
                    }

                    if(bFamil_chk == false)
                    {
                        //2016.06.03 태희철 추가
                        //오전긴급동의
                        if (_strCode2 == "P" && EmerDong == "65")
                        {
                            if (strIDchk == "1")
                            {
                                _sw = new StreamWriter(path + ".10001_퀵대상건_동의", true, _encoding);
                                _sw.WriteLine(_strLine + "0022602");
                                strData_Group = "1";
                            }
                            else
                            {
                                _sw = new StreamWriter(path + ".6500_오전(퀵)_동의", true, _encoding);
                                _sw.WriteLine(_strLine + "0022601");
                                strData_Group = "2";
                            }
                        }
                        else if ((_strFamilyCode == "0" && _strOwner_one == "0") || (_strFamilyCode == "1" && _strFamily_one == "1"))
                        {
                            _sw = new StreamWriter(path + ".5000_오전(퀵)_본인", true, _encoding);
                            _sw.WriteLine(_strLine + "0023104");
                            strData_Group = "3";
                        }
                        else
                        {
                            _sw = new StreamWriter(path + ".5000_오전(퀵)", true, _encoding);
                            _sw.WriteLine(_strLine + "0023101");
                            strData_Group = "4";
                        }
                        
                        //_strFamilyCode = 0 본인만 도는 본인+가족
                        // 0 값이 아닌경우 가족 아님
                        if (_strFamilyCode != "0")
                        {
                            bFamil_chk = false;
                        }
                        else if (iCard_cnt != 0)
                        {
                            iCard_cnt--;
                            bFamil_chk = true;
                        }
                        else if (strCard_cnt_chk == "2" || strCard_cnt_chk == "3" || strCard_cnt_chk == "4" || strCard_cnt_chk == "5" ||
                                strCard_cnt_chk == "6" || strCard_cnt_chk == "7" || strCard_cnt_chk == "8" || strCard_cnt_chk == "9"
                        )
                        {
                            iCard_cnt = Int16.Parse(strCard_cnt_chk);
                            iCard_cnt--;
                            bFamil_chk = true;
                        }
                        else
                        {
                            bFamil_chk = false;
                        }
                    }
                    else
                    {
                        if (strData_Group == "1")
                        {
                            _sw = new StreamWriter(path + ".10001_퀵대상건_동의", true, _encoding);
                            _sw.WriteLine(_strLine + "0022602");
                            strData_Group = "1";
                        }
                        else if (strData_Group == "2")
                        {
                            _sw = new StreamWriter(path + ".6500_오전(퀵)_동의", true, _encoding);
                            _sw.WriteLine(_strLine + "0022601");
                            strData_Group = "2";
                        }
                        else if (strData_Group == "3")
                        {
                            _sw = new StreamWriter(path + ".5000_오전(퀵)_본인", true, _encoding);
                            _sw.WriteLine(_strLine + "0023104");
                            strData_Group = "3";
                        }
                        else if (strData_Group == "4")
                        {
                            _sw = new StreamWriter(path + ".5000_오전(퀵)", true, _encoding);
                            _sw.WriteLine(_strLine + "0023101");
                            strData_Group = "4";
                        }

                        if (iCard_cnt != 0)
                        {
                            iCard_cnt--;
                            bFamil_chk = true;
                        }
                    }

                    //가족여부 체크
                    strCustomer_SSN_old = strCustomer_SSN;
                    
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

        //등록 자료 생성 2020.02.04 이전
        public static string ConvertRegister_back(string path, string xmlZipcodeAreaPath, string xmlZipcodePath)
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
            /// 긴급오전동의 EmerDong : 326~328
            string _strLine = "", _strCode = "", _strCode1 = null, _strCode2 = null;
            string _strReturn = "", _strOwner_one = "", _strFamily_one = "", _strFamilyCode = "", strIDchk = "";
            string EmerDong = "";

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
                    //_strCode = _encoding.GetString(_byteAry, 526, 5);
                    _strCode = _encoding.GetString(_byteAry, 307, 5);
                    _strCode1 = _encoding.GetString(_byteAry, 0, 1);
                    //동의서 구분 : 동의서 = D
                    _strCode2 = _encoding.GetString(_byteAry, 1, 1);

                    //제3자수령가능여부 1 = 동의, 0 = 미동의
                    _strOwner_one = _encoding.GetString(_byteAry, 398, 1);
                    //가족앞필수교부 : 1 = 동의, 0 = 미동의
                    _strFamily_one = _encoding.GetString(_byteAry, 399, 1);
                    //가족만발급 (0:본인만1매or본인+가족, 1:가족만1매, 2:가족만2매이상)
                    _strFamilyCode = _encoding.GetString(_byteAry, 345, 1);
                    //긴급동의
                    EmerDong = _encoding.GetString(_byteAry, 326, 2);
                    //신분증진위사후확인
                    strIDchk = _encoding.GetString(_byteAry, 426, 1);

                    //2016.06.03 태희철 추가
                    //오전긴급동의
                    if (_strCode2 == "P" && EmerDong == "65")
                    {
                        if (strIDchk == "1")
                        {
                            _sw = new StreamWriter(path + ".10001_퀵대상건_동의", true, _encoding);
                            _sw.WriteLine(_strLine + "0022602");
                        }
                        else
                        {
                            _sw = new StreamWriter(path + ".6500_오전(퀵)_동의", true, _encoding);
                            _sw.WriteLine(_strLine + "0022601");
                        }
                    }
                    else if ((_strFamilyCode == "0" && _strOwner_one == "0") || (_strFamilyCode == "1" && _strFamily_one == "1"))
                    {
                        _sw = new StreamWriter(path + ".5000_오전(퀵)_본인", true, _encoding);
                        _sw.WriteLine(_strLine + "0023104");
                    }
                    else
                    {
                        _sw = new StreamWriter(path + ".5000_오전(퀵)", true, _encoding);
                        _sw.WriteLine(_strLine + "0023101");
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

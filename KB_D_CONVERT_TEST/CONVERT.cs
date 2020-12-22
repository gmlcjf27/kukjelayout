using System;
using System.Collections.Generic;
using System.Collections;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace KB_D_CONVERT_TEST
{
    public class CONVERT
    {
        //기본 인코딩 설정
        private static string strEncoding = "ks_c_5601-1987";
        private static string strCardTypeID = "002_CONV";
        private static string strCardTypeName = "국민동의서컨버터_TEST";

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
            string _strLine = "", _strCode = "", _strCode1 = null, _strCode2 = null, _strReturn = "";

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
                    _strCode2 = _encoding.GetString(_byteAry, 1, 1);

                    if (_strCode1 == "3")
                    {
                        switch (_strCode2)
                        {
                            case "A":
                                _sw = new StreamWriter(path + ".20500_DONG_정기", true, _encoding);
                                _sw.WriteLine(_strLine + "0022201");
                                break;
                            case "B":
                                _sw = new StreamWriter(path + ".20000_정기", true, _encoding);
                                _sw.WriteLine(_strLine + "0021201");
                                break;
                            case "D":
                                _sw = new StreamWriter(path + ".2500_DONG_정기", true, _encoding);
                                _sw.WriteLine(_strLine + "0022110");
                                break;
                            case "G":
                                _sw = new StreamWriter(path + ".21000_정기", true, _encoding);
                                _sw.WriteLine(_strLine + "0023103");
                                break;
                            case "I":   //갱신
                                _sw = new StreamWriter(path + ".2000_정기", true, _encoding);
                                _sw.WriteLine(_strLine + "0021102");
                                break;
                            case "Q":   //긴급오후
                                _sw = new StreamWriter(path + ".1000_정기", true, _encoding);
                                _sw.WriteLine(_strLine + "0023102");
                                break;
                            case "N":   //갱신VVIP 
                                _sw = new StreamWriter(path + ".17000_정기", true, _encoding);
                                _sw.WriteLine(_strLine + "0023202");
                                break;
                            case "Y":   //광역VVIP
                                _sw = new StreamWriter(path + ".7500_DONG_정기", true, _encoding);
                                _sw.WriteLine(_strLine + "0022302");
                                break;
                            case "S":   //광역
                                _sw = new StreamWriter(path + ".61500_DONG_정기", true, _encoding);
                                _sw.WriteLine(_strLine + "0022301");
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
                                _sw = new StreamWriter(path + ".20500_DONG_추가", true, _encoding);
                                _sw.WriteLine(_strLine + "0022201");
                                break;
                            case "B":
                                _sw = new StreamWriter(path + ".20000_추가", true, _encoding);
                                _sw.WriteLine(_strLine + "0021201");
                                break;
                            case "D":
                                _sw = new StreamWriter(path + ".2500_DONG_추가", true, _encoding);
                                _sw.WriteLine(_strLine + "0022110");
                                break;
                            case "G":
                                _sw = new StreamWriter(path + ".21000_추가", true, _encoding);
                                _sw.WriteLine(_strLine + "0023103");
                                break;
                            case "I":   //갱신
                                _sw = new StreamWriter(path + ".2000_추가", true, _encoding);
                                _sw.WriteLine(_strLine + "0021102");
                                break;
                            case "Q":   //긴급오후
                                _sw = new StreamWriter(path + ".1000_추가", true, _encoding);
                                _sw.WriteLine(_strLine + "0023102");
                                break;
                            case "N":   // VVIP
                                _sw = new StreamWriter(path + ".7000_추가", true, _encoding);
                                _sw.WriteLine(_strLine + "0023201");
                                break;
                            case "Y":   // 광역VVIP
                                _sw = new StreamWriter(path + ".7500_DONG_추가", true, _encoding);
                                _sw.WriteLine(_strLine + "0022302");
                                break;
                            case "S":   // 광역
                                _sw = new StreamWriter(path + ".61500_DONG_추가", true, _encoding);
                                _sw.WriteLine(_strLine + "0022301");
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
                                _sw = new StreamWriter(path + ".20500_DONG_일일", true, _encoding);
                                _sw.WriteLine(_strLine + "0022201");
                                break;
                            case "B":
                                _sw = new StreamWriter(path + ".20000_일일", true, _encoding);
                                _sw.WriteLine(_strLine + "0021201");
                                break;
                            case "D":
                                //switch (ConvertStrCode(_strCode))
                                //{
                                //    case "SK":
                                //        _sw = new StreamWriter(path + ".3500_SK_일일", true, _encoding);
                                //        _sw.WriteLine(_strLine + "0022102");
                                //        break;
                                //    case "GS":
                                //        _sw = new StreamWriter(path + ".4500_GS_일일", true, _encoding);
                                //        _sw.WriteLine(_strLine + "0022103");
                                //        break;
                                //    case "HANA":
                                //        _sw = new StreamWriter(path + ".5500_HANA_일일", true, _encoding);
                                //        _sw.WriteLine(_strLine + "0022104");
                                //        break;
                                //    case "ELAND":
                                //        _sw = new StreamWriter(path + ".6500_ELAND_일일", true, _encoding);
                                //        _sw.WriteLine(_strLine + "0022105");
                                //        break;
                                //    case "HPLUSE":
                                //        _sw = new StreamWriter(path + ".9500_HPLUSE_일일", true, _encoding);
                                //        _sw.WriteLine(_strLine + "0022106");
                                //        break;
                                //    //case "MYONE": _sw = new StreamWriter(path + ".8500_MYONE_일일", true, _encoding); break;
                                //    case "SAN":
                                //        _sw = new StreamWriter(path + ".11500_SAN_일일", true, _encoding);
                                //        _sw.WriteLine(_strLine + "0022107");
                                //        break;
                                //    case "GREEN":
                                //        _sw = new StreamWriter(path + ".12500_GREEN_일일", true, _encoding);
                                //        _sw.WriteLine(_strLine + "0022108");
                                //        break;
                                //    case "U_PLUS":
                                //        _sw = new StreamWriter(path + ".13500_UPLUS_일일", true, _encoding);
                                //        _sw.WriteLine(_strLine + "0022109");
                                //        break;
                                //    default:
                                //        _sw = new StreamWriter(path + ".500_일일", true, _encoding);
                                //        _sw.WriteLine(_strLine + "0022101");
                                //        break;
                                //}
                                //break;

                                //2014.10.01 태희철 제휴구분없이 일반으로 통일
                                _sw = new StreamWriter(path + ".500_일일", true, _encoding);
                                _sw.WriteLine(_strLine + "0022101");
                                break;
                            case "G":
                                _sw = new StreamWriter(path + ".21000_일일", true, _encoding);
                                _sw.WriteLine(_strLine + "0023103");
                                break;
                            case "I":   //갱신
                                _sw = new StreamWriter(path + ".100_일일", true, _encoding);
                                _sw.WriteLine(_strLine + "0021101");
                                break;
                            case "Q":   //긴급오후
                                _sw = new StreamWriter(path + ".1000_일일", true, _encoding);
                                _sw.WriteLine(_strLine + "0023102");
                                break;
                            case "N":   // VVIP
                                if (ConvertStrCode(_strCode) == "미르")
                                {
                                    _sw = new StreamWriter(path + ".30000_일일", true, _encoding);
                                    _sw.WriteLine(_strLine + "0023401");
                                }
                                else
                                {
                                    _sw = new StreamWriter(path + ".7000_일일", true, _encoding);
                                    _sw.WriteLine(_strLine + "0023201");
                                }
                                break;
                            case "Y":   // 광역VVIP-동의
                                if (ConvertStrCode(_strCode) == "미르")
                                {
                                    _sw = new StreamWriter(path + ".30500_DONG_일일", true, _encoding);
                                    _sw.WriteLine(_strLine + "0022501");
                                }
                                else
                                {
                                    _sw = new StreamWriter(path + ".7500_DONG_일일", true, _encoding);
                                    _sw.WriteLine(_strLine + "0022302");
                                }
                                break;
                            case "S":   // 광역
                                _sw = new StreamWriter(path + ".61500_DONG_일일", true, _encoding);
                                _sw.WriteLine(_strLine + "0022301");
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

using System;
using System.Collections.Generic;
using System.Collections;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace _007_카카오뱅크_CONVERT
{
    public class CONVERT
    {
        //기본 인코딩 설정
        private static string strEncoding = "ks_c_5601-1987";
        private static string strCardTypeID = "007_CONV";
        private static string strCardTypeName = "카카오뱅크_컨버터";

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
            string _strReturn = null;
            // _strCode : 제휴코드
            // _strCode2 : 갱신여부 (일일 = 1, 추가(갱신) = 2, 정기(갱신) = 3)
            string _strLine = "", _strCode = "", _strCode2 = "";
            // DataRow[] _drs = null;
            DataTable _dtable = null;
            DataSet _dsetZipcodeArea = null;


            string _strOwner_one = "", strMini = "";

            try
            {
                _dtable = new DataTable("CONVERT");

                _dsetZipcodeArea = new DataSet();
                _dsetZipcodeArea.ReadXml(xmlZipcodeAreaPath);


                //파일 읽기 Stream과 오류시 저장할 쓰기 Stream 생성
                _sr = new StreamReader(path, _encoding);
                _swError = new StreamWriter(path + ".Error", false, _encoding);


                while ((_strLine = _sr.ReadLine()) != null)
                {
                    _byteAry = _encoding.GetBytes(_strLine);
                    //배송데이터형식구분 : 일반 = I, 재발송 = B
                    _strCode = _encoding.GetString(_byteAry, 1, 1);
                    _strCode2 = _encoding.GetString(_byteAry, 0, 1);
                    //제3자수령가능여부 : 1=가능(대리수령가능), 0=불가능(본인만배송)
                    _strOwner_one = _encoding.GetString(_byteAry, 398, 1);
                    //카카오미니 2020.07.29
                    strMini = _encoding.GetString(_byteAry, 402, 1);

                    //2020.08.03 태희철 수정
                    //카카오미니 추가
                    if (strMini == "1")
                    {
                        if (_strCode == "I")
                        {
                            _sw = new StreamWriter(path + ".5000_미니_일반", true, _encoding);
                            _sw.WriteLine(_strLine + "0071201");
                        }
                        else if (_strCode == "B")
                        {
                            _sw = new StreamWriter(path + ".6000_미니_재발송", true, _encoding);
                            _sw.WriteLine(_strLine + "0071202");
                        }
                        else
                        {
                            _sw = new StreamWriter(path + ".기타_오류", true, _encoding);
                            _sw.WriteLine(_strLine + "");
                        }
                    }
                    else if (_strOwner_one == "1")
                    {
                        if (_strCode == "I")
                        {
                            if (_strCode2 == "3")
                            {
                                _sw = new StreamWriter(path + ".7000_갱신_정기", true, _encoding);
                                _sw.WriteLine(_strLine + "0071105");
                            }
                            else if (_strCode2 == "2")
                            {
                                _sw = new StreamWriter(path + ".8000_갱신_추가", true, _encoding);
                                _sw.WriteLine(_strLine + "0071106");
                            }
                            else
                            {
                                _sw = new StreamWriter(path + ".2000_일반", true, _encoding);
                                _sw.WriteLine(_strLine + "0071103");
                            }
                        }
                        else if (_strCode == "B")
                        {
                            _sw = new StreamWriter(path + ".3000_재발송", true, _encoding);
                            _sw.WriteLine(_strLine + "0071104");
                        }
                        else
                        {
                            _sw = new StreamWriter(path + ".기타_오류", true, _encoding);
                            _sw.WriteLine(_strLine + "");
                        }
                    }
                    else
                    {
                        if (_strCode == "I")
                        {
                            if (_strCode2 == "3")
                            {
                                _sw = new StreamWriter(path + ".9000_갱신_일반_본인_정기", true, _encoding);
                                _sw.WriteLine(_strLine + "0071107");
                            }
                            else if (_strCode2 == "2")
                            {
                                _sw = new StreamWriter(path + ".10000_갱신_일반_본인_추가", true, _encoding);
                                _sw.WriteLine(_strLine + "0071108");
                            }
                            else
                            {
                                _sw = new StreamWriter(path + ".100_일반_본인", true, _encoding);
                                _sw.WriteLine(_strLine + "0071101");
                            }
                        }
                        else if (_strCode == "B")
                        {
                            _sw = new StreamWriter(path + ".1000_재발송_본인", true, _encoding);
                            _sw.WriteLine(_strLine + "0071102");
                        }
                        else
                        {
                            _sw = new StreamWriter(path + ".기타_오류", true, _encoding);
                            _sw.WriteLine(_strLine + "");
                        }
                    }
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

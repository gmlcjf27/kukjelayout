using System;
using System.Collections.Generic;
using System.Collections;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace SH_CONVERT
{
    public class CONVERT
    {
        //기본 인코딩 설정
        private static string strEncoding = "ks_c_5601-1987";
        private static string strCardTypeID = "003_CONV_Tax";
        private static string strCardTypeName = "신한컨버터_Tax";

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
            string _strLine = "";
            string _strReturn = "";
            string _col_gubun = "", strChk_code = "";

            try
            {
                //파일 읽기 Stream과 오류시 저장할 쓰기 Stream 생성
                _sr = new StreamReader(path, _encoding);
                _swError = new StreamWriter(path + ".Error", false, _encoding);

                while ((_strLine = _sr.ReadLine()) != null)
                {
                    //2012.10.12 태희철 수정
                    _byteAry = _encoding.GetBytes(_strLine);
                    _col_gubun = _encoding.GetString(_byteAry, 23, 1).Trim().ToUpper();
                    //법인구분 : Z = 법인
                    strChk_code = _encoding.GetString(_byteAry, 24, 1).ToUpper();

                    //본인만배송건
                    if (_encoding.GetString(_byteAry, 44, 1) == "2")
                    {
                        if (strChk_code == "Z")
                        {
                            _sw = new StreamWriter(path + ".4000_법인-본인", true, _encoding);
                            _sw.WriteLine(_strLine + "0031105");
                        }
                        else
                        {
                            _sw = new StreamWriter(path + ".2000_본인만", true, _encoding);
                            _sw.WriteLine(_strLine + "0031103");
                        }
                    }
                    else if (strChk_code == "Z")
                    {
                        _sw = new StreamWriter(path + ".3000_법인", true, _encoding);
                        _sw.WriteLine(_strLine + "0031104");
                    }
                    else if (_col_gubun == "Q")
                    {
                        _sw = new StreamWriter(path + ".1000-q", true, _encoding);
                        _sw.WriteLine(_strLine + "0031102");
                    }
                    else
                    {
                        _sw = new StreamWriter(path + ".100", true, _encoding);
                        _sw.WriteLine(_strLine + "0031101");
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

using System;
using System.Collections.Generic;
using System.Collections;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace _004_우체국_삼성_일반_CONVERT
{
    public class CONVERT
    {
        //기본 인코딩 설정
        private static string strEncoding = "ks_c_5601-1987";
        private static string strCardTypeID = "004_CONVJ";
        private static string strCardTypeName = "우체국_삼성-일반-컨버터";
        //private static char chCSV = ',';
        

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

            StreamReader _sr = null;
            StreamReader _sam = null;//파일 읽기 스트림
            StreamWriter _swError = null;
            StreamWriter _sw = null;
            //StreamWriter _swdong = null;
            byte[] _byteAry = null;
            string _strCode = "", strAdd_check = "";
            string _strLine = "";
            string _strReturn = "";
            //DataTable _dtable = null;
            DataSet _dsetZipcodeArea = null;

            try
            {

                //_dtable = new DataTable("CONVERT");
                ////_dtable.Columns.Add("code");
                //_dtable.Columns.Add("card_design_code");

                _dsetZipcodeArea = new DataSet();
                _dsetZipcodeArea.ReadXml(xmlZipcodeAreaPath);

                //파일 읽기 Stream과 오류시 저장할 쓰기 Stream 생성
                _sr = new StreamReader(path, _encoding);
                _swError = new StreamWriter(path + ".Error", false, _encoding);

                //2012.10.12 태희철 수정
                while ((_strLine = _sr.ReadLine()) != null)
                {
                    _byteAry = _encoding.GetBytes(_strLine);
                    _strCode = _encoding.GetString(_byteAry, 718, 1);

                    if (_strCode == "Y")
                    {
                        _sw = new StreamWriter(path + ".1000_우체국_삼성일반_본인", true, _encoding);
                        _sw.WriteLine(_strLine + "0041102");
                    }
                    else
                    {
                        _sw = new StreamWriter(path + ".100_우체국_삼성일반", true, _encoding);
                        _sw.WriteLine(_strLine + "0041101");
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

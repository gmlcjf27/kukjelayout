using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace _090_농협_CONVERT
{
    public class CONVERT
    {
        //기본 인코딩 설정
        private static string strEncoding = "ks_c_5601-1987";
        private static string strCardTypeID = "089_CONV";
        private static string strCardTypeName = "NH농협컨버터";

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
            string _strZipcode = "", _strReturn = "", _strLine = "";
            //strGubun1=1000차, strGubun2 = 2000차, strDong=동의서, strGubun4 = 영업점특송 BOX
            string strGubun1 = null, strBox = null, strGubun3 = null, strDong_ADD = "", strOffice_chk = "";
            string strGubun4 = "", strGubun5 = "", strDong_chk = "", strDelivery_Place = "";
            DataTable _dtable = null;
            DataSet _dsetZipcodeArea = null;
            //DataRow _dr = null;
            //DataRow[] _drs = null;
            int _iCount = 0;
            try
            {
                _dtable = new DataTable("CONVERT");
                _dtable.Columns.Add("card_bank_ID");
                _dtable.Columns.Add("card_zipcode");
                _dtable.Columns.Add("data");

                //파일 읽기 Stream과 오류시 저장할 쓰기 Stream 생성
                _sr = new StreamReader(path, _encoding);
                _swError = new StreamWriter(path + ".Error", false, _encoding);


                while ((_strLine = _sr.ReadLine()) != null)
                {
                    _byteAry = _encoding.GetBytes(_strLine);
                    _strZipcode = _encoding.GetString(_byteAry, 179, 6);
                    //_drs = _dsetZipcodeArea.Tables[0].Select("zipcode=" + _strZipcode);
                    
                    _iCount++;

                    /// 2019.05.10 차세대 이전
                    /// 긴급및동의구분 : 2 (동의서), strGubun1 = 3 (긴급)
                    /// strGubun1 = _encoding.GetString(_byteAry, 41, 1);
                    /// strGubun2 = 7 (영업점특송), 8:Box배송, 9:정복배송 2017.07.04 추가
                    /// strGubun2 = _encoding.GetString(_byteAry, 40, 1);
                    /// strDelivery_Place = _encoding.GetString(_byteAry, 536, 1);
                    /// strOffice_chk = _encoding.GetString(_byteAry, 1294, 7);
                    /// strDong_ADD = _encoding.GetString(_byteAry, 1323, 1);


                    strGubun3 = _encoding.GetString(_byteAry, 0, 2); //DATA 구분
                    //2019.05.10 수정
                    //strGubun2 = 1 : 일반, 2 : BOX
                    //카드봉입분류
                    strBox = _encoding.GetString(_byteAry, 40, 1);
                    //긴급발급구분
                    //strGubun1 = 0 : 일반, 1 : 긴급, 2: 특급
                    strGubun1 = _encoding.GetString(_byteAry, 41, 1);
                    //배송수령지구분
                    //1 : 자택, 2 : 직장, 3 : 제3청구지, 4 : 영업점, 5 : 지자체, 6 : 직접수령
                    strDelivery_Place = _encoding.GetString(_byteAry, 42, 1);
                    //strDong_chk = 동의서징구여부 : 1 징구, 0 : 징구아님
                    strDong_chk = _encoding.GetString(_byteAry, 1153, 1);
                    //거래내용
                    strOffice_chk = _encoding.GetString(_byteAry, 1173, 7);
                    //별지동의서코드
                    //1 : GS리테일/OK캐쉬백, 2 : GS리테일
                    strDong_ADD = _encoding.GetString(_byteAry, 1203, 1);
                    //영업점특송Box 일반 : 804, 영업점특송BOX 동의 : 814
                    strGubun4 = strBox + strDong_chk + strDelivery_Place;
                    //배송업무기준코드(VIP,갱신 등)
                    strGubun5 = _encoding.GetString(_byteAry, 43, 3); //DATA 구분
                    

                    if (strGubun3 == "FD")
                    {
                        switch (strGubun5)
                        {
                            case "001":
                            case "003":
                                _sw = new StreamWriter(path + "NH_100", true, _encoding);
                                _sw.WriteLine(_strLine + "0901101");
                                break;
                            case "002":
                                _sw = new StreamWriter(path + "NH_일반본인_3000", true, _encoding);
                                _sw.WriteLine(_strLine + "0901102");
                                break;
                            case "011":
                                _sw = new StreamWriter(path + "NH_영업점_2000", true, _encoding);
                                _sw.WriteLine(_strLine + "0903201");
                                break;
                            case "021":
                                if (strOffice_chk == "1070302")
                                {
                                    _sw = new StreamWriter(path + "NH_DONG_기업", true, _encoding);
                                    _sw.WriteLine(_strLine + "0902104");
                                }
                                else if (strDong_ADD == "1")
                                {
                                    _sw = new StreamWriter(path + "NH_DONG_4500_OK캐쉬백", true, _encoding);
                                    _sw.WriteLine(_strLine + "0902105");
                                }
                                else if (strDong_ADD == "2")
                                {
                                    _sw = new StreamWriter(path + "NH_DONG_5500_GS리테일", true, _encoding);
                                    _sw.WriteLine(_strLine + "0902106");
                                }
                                else
                                {
                                    _sw = new StreamWriter(path + "NH_DONG_500", true, _encoding);
                                    _sw.WriteLine(_strLine + "0902101");
                                }
                                break;
                            case "024":
                                _sw = new StreamWriter(path + "NH_광역동의", true, _encoding);
                                _sw.WriteLine(_strLine + "0902102");
                                break;
                            case "031":
                                _sw = new StreamWriter(path + "NH_일반BOX_4000", true, _encoding);
                                _sw.WriteLine(_strLine + "0901103");
                                break;
                            case "032":
                                _sw = new StreamWriter(path + "NH_영업점_6000_영업점BOX", true, _encoding);
                                _sw.WriteLine(_strLine + "0903202");
                                break;
                            case "033":
                                _sw = new StreamWriter(path + "NH_DONG_2500_Box", true, _encoding);
                                _sw.WriteLine(_strLine + "0902103");
                                break;
                            case "034":
                                _sw = new StreamWriter(path + "NH_긴급BOX_4000", true, _encoding);
                                _sw.WriteLine(_strLine + "0903103");
                                break;
                            case "041":
                            case "042":
                                _sw = new StreamWriter(path + "NH_긴급_1000", true, _encoding);
                                _sw.WriteLine(_strLine + "0903101");
                                break;
                            case "051":
                                _sw = new StreamWriter(path + "NH_지점_재발송", true, _encoding);
                                _sw.WriteLine(_strLine + "0904101");
                                break;
                            case "071":
                                _sw = new StreamWriter(path + "NH_긴급VIP_3000", true, _encoding);
                                _sw.WriteLine(_strLine + "0903102");
                                break;
                            default:
                                _sw = new StreamWriter(path + "구분코드오류", true, _encoding);
                                _sw.WriteLine(_strLine + "");
                                break;
                        }
                        _sw.Close();
                    }
                }

                _strReturn = "성공";
            }
            catch (Exception ex)
            {
                _strReturn = string.Format("{0}번째 데이터 변환 중 오류", _iCount + 1);
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

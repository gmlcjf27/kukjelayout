using System;
using System.Collections.Generic;
using System.Collections;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace _003_신한_동의_CONVERT
{
    public class CONVERT
    {
        //기본 인코딩 설정
        private static string strEncoding = "ks_c_5601-1987";
        private static string strCardTypeID = "003_CONV";
        private static string strCardTypeName = "신한컨버터";

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
        //2012.07.05 태희철 수정
        public static string ConvertRegister(string path, string xmlZipcodeAreaPath, string xmlZipcodePath)
        {
            System.Text.Encoding _encoding = System.Text.Encoding.GetEncoding(strEncoding);	//기본 인코딩	
            StreamReader _sr = null;																					//파일 읽기 스트림
            //StreamReader _sh = null;
            StreamWriter _swError = null;
            StreamWriter _sw = null;
            byte[] _byteAry = null;
            string _strCode = "", _strDeliveryPlaceType = "";
            string _strLine = "";
            string _strReturn = "";
            //DataTable dt_code = null;
            //DataSet _dsetZipcodeArea = null;
            //string strDegree = null;

            try
            {

                //_dsetZipcodeArea = new DataSet();
                //_dsetZipcodeArea.ReadXml(xmlZipcodeAreaPath);


                _sr = new StreamReader(path, _encoding);
                _swError = new StreamWriter(path + ".Error", false, _encoding);

                //dt_code = dset.Tables[0];

                while ((_strLine = _sr.ReadLine()) != null)
                {
                    _byteAry = _encoding.GetBytes(_strLine);

                    _strCode = _encoding.GetString(_byteAry, 211, 4).Trim().ToUpper();
                    _strDeliveryPlaceType = _encoding.GetString(_byteAry, 45, 1);
                    //제휴코드 변경 4byte -> 6byte
                    //2012.10.10 태희철 수정
                    //제휴코드 변경 4byte -> 6byte
                    //2012.10.10 태희철 수정
                    //선반납 구분
                    if (_strDeliveryPlaceType.Trim() == "")
                    {
                        _sw = new StreamWriter(path + ".500", true, _encoding);
                        _sw.WriteLine(_strLine + "0032101");
                    }
                    else
                    {
                        switch (_strCode)
                        {
                            case "BA01": //1
                            case "BA16": //대한노인회동의서
                            case "BA17": //현대아이파크
                                _sw = new StreamWriter(path + ".500", true, _encoding);
                                _sw.WriteLine(_strLine + "0032101");
                                break;
                            case "BA03": //A
                            case "BB27": //2014.10.21 태희철 신규
                                _sw = new StreamWriter(path + ".12500_U_하이", true, _encoding);
                                _sw.WriteLine(_strLine + "0032112");
                                break;
                            case "BA04": //B
                            case "BB28": //2014.10.21 태희철 신규
                                _sw = new StreamWriter(path + ".13500_U_스폰서", true, _encoding);
                                _sw.WriteLine(_strLine + "0032113");
                                break;
                            case "BA05": //D
                            case "BB29": //2014.10.21 태희철 신규
                                _sw = new StreamWriter(path + ".6500_U_스마트", true, _encoding);
                                _sw.WriteLine(_strLine + "0032106");
                                break;
                            case "BA06": //N
                                _sw = new StreamWriter(path + ".23000_LGU_Class", true, _encoding);
                                _sw.WriteLine(_strLine + "0032123");
                                break;
                            case "BA07": //C
                            case "BB25": //2014.10.21 태희철 신규
                                _sw = new StreamWriter(path + ".16500_국민연금", true, _encoding);
                                _sw.WriteLine(_strLine + "0032116");
                                break;
                            case "BB05": //E
                                _sw = new StreamWriter(path + ".18000_CJ", true, _encoding);
                                _sw.WriteLine(_strLine + "0032118");
                                break;
                            case "BA08": //F
                                _sw = new StreamWriter(path + ".5500_해피오토", true, _encoding);
                                _sw.WriteLine(_strLine + "0032105");
                                break;
                            case "BA09": //G
                                _sw = new StreamWriter(path + ".11000_해피라이프", true, _encoding);
                                _sw.WriteLine(_strLine + "0032111");
                                break;
                            case "BB06": //대한항공
                            case "BC01": //J
                                _sw = new StreamWriter(path + ".19000_대한항공", true, _encoding);
                                _sw.WriteLine(_strLine + "0032119");
                                break;
                            case "BB07": //아시아나
                            case "BC02": //K
                                _sw = new StreamWriter(path + ".20000_아시아나", true, _encoding);
                                _sw.WriteLine(_strLine + "0032120");
                                break;
                            case "BA10": //M
                            case "BB30": //2014.10.21 태희철 신규
                                _sw = new StreamWriter(path + ".21000_LTE", true, _encoding);
                                _sw.WriteLine(_strLine + "0032121");
                                break;
                            case "BB08": //O
                                _sw = new StreamWriter(path + ".24000_현대백화점(체크)", true, _encoding);
                                _sw.WriteLine(_strLine + "0032124");
                                break;
                            case "BB09": //P
                                _sw = new StreamWriter(path + ".22000_하나투어", true, _encoding);
                                _sw.WriteLine(_strLine + "0032122");
                                break;
                            case "BB10": //Q
                                _sw = new StreamWriter(path + ".25000_이랜드", true, _encoding);
                                _sw.WriteLine(_strLine + "0032125");
                                break;
                            case "BC03": //R
                                _sw = new StreamWriter(path + ".26000_이랜드_대한", true, _encoding);
                                _sw.WriteLine(_strLine + "0032126");
                                break;
                            case "BC04": //S
                                _sw = new StreamWriter(path + ".27000_이랜드_아시아나", true, _encoding);
                                _sw.WriteLine(_strLine + "0032127");
                                break;
                            case "BB11": //T
                                _sw = new StreamWriter(path + ".28000_GS리테일", true, _encoding);
                                _sw.WriteLine(_strLine + "0032128");
                                break;
                            case "BB12": //U
                                _sw = new StreamWriter(path + ".29000_AK신한", true, _encoding);
                                _sw.WriteLine(_strLine + "0032129");
                                break;
                            case "BB13": //V
                                _sw = new StreamWriter(path + ".30000_현대백화점(신용)", true, _encoding);
                                _sw.WriteLine(_strLine + "0032130");
                                break;
                            case "BA12": //X
                            case "BB31": //2014.10.21 태희철 신규
                                _sw = new StreamWriter(path + ".31000_SKTSave", true, _encoding);
                                _sw.WriteLine(_strLine + "0032131");
                                break;
                            case "BB01": //3
                            case "BB14": //W
                                _sw = new StreamWriter(path + ".3500_GS", true, _encoding);
                                _sw.WriteLine(_strLine + "0032103");
                                break;
                            case "BB02": //4
                                _sw = new StreamWriter(path + ".4500_SK엔크린", true, _encoding);
                                _sw.WriteLine(_strLine + "0032104");
                                break;
                            //case "4": _sw = new StreamWriter(path + ".7500_SK오일", true, _encoding);
                            //    _sw.WriteLine(_strLine + "0032107");
                            //    break;
                            case "BB03": //6
                                _sw = new StreamWriter(path + ".9500_SK행복", true, _encoding);
                                _sw.WriteLine(_strLine + "0032109");
                                break;
                            case "BA02": //7
                            case "BB26": //2014.10.21 태희철 신규
                                _sw = new StreamWriter(path + ".10500_LGT하이세이브", true, _encoding);
                                _sw.WriteLine(_strLine + "0032110");
                                break;
                            case "BB04": //9 신세계
                                _sw = new StreamWriter(path + ".17000_U_Emart", true, _encoding);
                                _sw.WriteLine(_strLine + "0032117");
                                break;
                            case "BA13":
                                _sw = new StreamWriter(path + ".32000_한샘Style", true, _encoding);
                                _sw.WriteLine(_strLine + "0032132");
                                break;
                            case "BB15": //해피포인트
                                _sw = new StreamWriter(path + ".33000_해피포인트", true, _encoding);
                                _sw.WriteLine(_strLine + "0032133");
                                break;
                            case "BA14": //LGU+ Saveㅂ
                            case "BB32": //2014.10.21 태희철 신규
                                _sw = new StreamWriter(path + ".34000_Save", true, _encoding);
                                _sw.WriteLine(_strLine + "0032134");
                                break;
                            case "BB18": //코웨이 별지
                                _sw = new StreamWriter(path + ".35000_payFree", true, _encoding);
                                _sw.WriteLine(_strLine + "0032135");
                                break;
                            case "BB19": //미샤 별지
                                _sw = new StreamWriter(path + ".36000_미샤", true, _encoding);
                                _sw.WriteLine(_strLine + "0032136");
                                break;
                            case "BB20": //현대아이파크(아이멤버스) 별지
                                _sw = new StreamWriter(path + ".37000_아이파크", true, _encoding);
                                _sw.WriteLine(_strLine + "0032137");
                                break;
                            case "BB16": //홈플러스(훼밀리카드)
                                _sw = new StreamWriter(path + ".38000_홈플러스", true, _encoding);
                                _sw.WriteLine(_strLine + "0032138");
                                break;
                            case "BB17": //화물택시
                                _sw = new StreamWriter(path + ".39000_화물택시", true, _encoding);
                                _sw.WriteLine(_strLine + "0032139");
                                break;
                            case "BB21": //LG전자
                                _sw = new StreamWriter(path + ".40000_LG전자", true, _encoding);
                                _sw.WriteLine(_strLine + "0032140");
                                break;
                            case "BB22": //교보문고
                                _sw = new StreamWriter(path + ".41000_교보문고", true, _encoding);
                                _sw.WriteLine(_strLine + "0032141");
                                break;
                            case "BB23": //베니건스
                                _sw = new StreamWriter(path + ".42000_베니건스", true, _encoding);
                                _sw.WriteLine(_strLine + "0032142");
                                break;
                            case "BA15": //Simple Free
                                _sw = new StreamWriter(path + ".43000_Simple", true, _encoding);
                                _sw.WriteLine(_strLine + "0032143");
                                break;
                            case "BB24": //대한노인회
                                _sw = new StreamWriter(path + ".44000_대한노인회", true, _encoding);
                                _sw.WriteLine(_strLine + "0032144");
                                break;
                            case "BB33": //(우대용 교통카드[시니어/실버])
                                _sw = new StreamWriter(path + ".45000_우대교통", true, _encoding);
                                _sw.WriteLine(_strLine + "0032145");
                                break;
                            case "BB34": //EnClean Bonus Card(파란색)
                                _sw = new StreamWriter(path + ".46000_개인택시", true, _encoding);
                                _sw.WriteLine(_strLine + "0032146");
                                break;
                            case "BB35": //EnClean Bonus Card(녹색)
                                _sw = new StreamWriter(path + ".47000_개인화물복지", true, _encoding);
                                _sw.WriteLine(_strLine + "0032147");
                                break;
                            case "BB36": //SOIL
                                _sw = new StreamWriter(path + ".48000_SOIL", true, _encoding);
                                _sw.WriteLine(_strLine + "0032148");
                                break;
                            case "BB37": //GS리테일팝
                                _sw = new StreamWriter(path + ".49000_GS리테일팝", true, _encoding);
                                _sw.WriteLine(_strLine + "0032149");
                                break;
                            case "BB38": //Olleh 슈퍼세이브
                                _sw = new StreamWriter(path + ".14000_olleh", true, _encoding);
                                _sw.WriteLine(_strLine + "0032114");
                                break;
                            case "BB39": //프리드라이프
                                _sw = new StreamWriter(path + ".50000_프리드라이프", true, _encoding);
                                _sw.WriteLine(_strLine + "0032150");
                                break;
                            case "BB40": //씽화물복지(GS&POINT)
                                _sw = new StreamWriter(path + ".51000_씽화물(GS)", true, _encoding);
                                _sw.WriteLine(_strLine + "0032151");
                                break;
                            case "BC05": //씽화물복지(GS&POINT) + 현대오일뱅크(파란색)
                                _sw = new StreamWriter(path + ".52000_씽화물(SOIL)", true, _encoding);
                                _sw.WriteLine(_strLine + "0032152");
                                break;
                            case "BB41": //통일과나눔
                                _sw = new StreamWriter(path + ".53000_통일과나눔", true, _encoding);
                                _sw.WriteLine(_strLine + "0032153");
                                break;
                            case "BB42": //제주항공
                                _sw = new StreamWriter(path + ".54000_제주항공", true, _encoding);
                                _sw.WriteLine(_strLine + "0032154");
                                break;
                            case "BB43": //대명빅플러스
                                _sw = new StreamWriter(path + ".55000_대명빅플러스", true, _encoding);
                                _sw.WriteLine(_strLine + "0032155");
                                break;
                            case "BB44": //예다함BigPlus
                                _sw = new StreamWriter(path + ".56000_예다함빅플러스", true, _encoding);
                                _sw.WriteLine(_strLine + "0032156");
                                break;
                            case "BB45": //SRT
                                _sw = new StreamWriter(path + ".57000_SRT", true, _encoding);
                                _sw.WriteLine(_strLine + "0032157");
                                break;
                            case "BB47": //현대오일뱅크
                                _sw = new StreamWriter(path + ".58000_현대오일뱅크", true, _encoding);
                                _sw.WriteLine(_strLine + "0032158");
                                break;
                            case "BB46": //S-oil보너스
                                _sw = new StreamWriter(path + ".59000_SOIL_보너스", true, _encoding);
                                _sw.WriteLine(_strLine + "0032159");
                                break;
                            case "BB48": //신세계
                                _sw = new StreamWriter(path + ".60000_신세계", true, _encoding);
                                _sw.WriteLine(_strLine + "0032160");
                                break;
                            case "BC07": //신세계(대한항공)
                                _sw = new StreamWriter(path + ".61000_신세계(대한항공)", true, _encoding);
                                _sw.WriteLine(_strLine + "0032161");
                                break;
                            case "BC08": //신세계(아시아나)
                                _sw = new StreamWriter(path + ".62000_신세계(아시아나)", true, _encoding);
                                _sw.WriteLine(_strLine + "0032162");
                                break;
                            case "BB49": //오일뱅크화물복지
                                _sw = new StreamWriter(path + ".63000_오일뱅크화물복지", true, _encoding);
                                _sw.WriteLine(_strLine + "0032163");
                                break;
                            case "BB50": //홈플러스신한카드
                                _sw = new StreamWriter(path + ".64000_홈플러스신한카드", true, _encoding);
                                _sw.WriteLine(_strLine + "0032164");
                                break;
                            case "BC09": //홈플러스OKCashbag
                                _sw = new StreamWriter(path + ".65000_홈플러스OKCashbag", true, _encoding);
                                _sw.WriteLine(_strLine + "0032165");
                                break;
                            case "BB51": //홈플러스OKCashbag
                                _sw = new StreamWriter(path + ".66000_마이홈플러스", true, _encoding);
                                _sw.WriteLine(_strLine + "0032166");
                                break;
                            case "BC10": //홈플러스OKCashbag
                                _sw = new StreamWriter(path + ".67000_마이홈플러스OK", true, _encoding);
                                _sw.WriteLine(_strLine + "0032167");
                                break;
                            case "BB52": //SRT체크
                                _sw = new StreamWriter(path + ".68000_SRT체크", true, _encoding);
                                _sw.WriteLine(_strLine + "0032168");
                                break;
                            case "BC11": //마이홈플러스OK(임직원)
                                _sw = new StreamWriter(path + ".69000_마이홈플러스OK(임직원)", true, _encoding);
                                _sw.WriteLine(_strLine + "0032169");
                                break;
                            case "BB53": //마이홈플러스(임직원)
                                _sw = new StreamWriter(path + ".70000_마이홈플러스(임직원)", true, _encoding);
                                _sw.WriteLine(_strLine + "0032170");
                                break;
                            case "BB55": //워너원 체크
                                _sw = new StreamWriter(path + ".71000_워너원", true, _encoding);
                                _sw.WriteLine(_strLine + "0032171");
                                break;
                            case "BB57": //CJONE2
                                _sw = new StreamWriter(path + ".72000_CJONE2", true, _encoding);
                                _sw.WriteLine(_strLine + "0032172");
                                break;
                            case "BB60": //11번가SKPAY
                                _sw = new StreamWriter(path + ".75000_11번가SKPAY", true, _encoding);
                                _sw.WriteLine(_strLine + "0032175");
                                break;
                            case "BB61": //메가마트
                                _sw = new StreamWriter(path + ".85000_메가마트", true, _encoding);
                                _sw.WriteLine(_strLine + "0032176");
                                break;
                            case "BB63": //롯데멤버스
                                _sw = new StreamWriter(path + ".76000_롯데멤버스", true, _encoding);
                                _sw.WriteLine(_strLine + "0032177");
                                break;
                            case "BB64": //롯데마트 2020.08.20
                                _sw = new StreamWriter(path + ".77000_롯데마트", true, _encoding);
                                _sw.WriteLine(_strLine + "0032178");
                                break;
                            case "BC06": //홈플러스대한항공
                                _sw = new StreamWriter(path + ".73000_홈플러스대한항공", true, _encoding);
                                _sw.WriteLine(_strLine + "0032173");
                                break;
                            case "BC12": //성남아동수당
                                _sw = new StreamWriter(path + ".74000_성남아동수당", true, _encoding);
                                _sw.WriteLine(_strLine + "0032174");
                                break;
                            default:
                                _sw = new StreamWriter(path + ".그외", true, _encoding);
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
                if (_sw != null) _sw.Close();
                if (_swError != null) _swError.Close();
                //if (dt_code != null) dt_code.Clear();
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

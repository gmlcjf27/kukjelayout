using System;
using System.Collections.Generic;
using System.Collections;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace _004_우체국_삼성_동_CONVERT
{
    public class CONVERT
    {
        //기본 인코딩 설정
        private static string strEncoding = "ks_c_5601-1987";
        private static string strCardTypeID = "004_CONVJ";
        private static string strCardTypeName = "우체국_삼성-동-컨버터";
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
            DataTable _dtable = null;
            DataSet _dsetZipcodeArea = null;

            try
            {

                _dtable = new DataTable("CONVERT");
                //_dtable.Columns.Add("code");
                _dtable.Columns.Add("card_design_code");

                _dsetZipcodeArea = new DataSet();
                _dsetZipcodeArea.ReadXml(xmlZipcodeAreaPath);

                //파일 읽기 Stream과 오류시 저장할 쓰기 Stream 생성
                _sr = new StreamReader(path, _encoding);
                _swError = new StreamWriter(path + ".Error", false, _encoding);

                //2012.10.12 태희철 수정
                while ((_strLine = _sr.ReadLine()) != null)
                {
                    _byteAry = _encoding.GetBytes(_strLine);
                    _strCode = _encoding.GetString(_byteAry, 471, 5);
                    strAdd_check = _encoding.GetString(_byteAry, 700, 4);

                    if (strAdd_check.Substring(0, 2) == "YA")
                    {
                        _sw = new StreamWriter(path + ".23000-신세계(별지)", true, _encoding);
                        _sw.WriteLine(_strLine + "0042123");
                    }
                    else if (strAdd_check.Substring(0, 2) == "YB")
                    {
                        _sw = new StreamWriter(path + ".24000-GS칼텍스(별지)", true, _encoding);
                        _sw.WriteLine(_strLine + "0042124");
                    }
                    else if (strAdd_check.Substring(0, 2) == "YC")
                    {
                        _sw = new StreamWriter(path + ".25000-SOIL멤버십(별지)", true, _encoding);
                        _sw.WriteLine(_strLine + "0042125");
                    }
                    else if (strAdd_check.Substring(0, 2) == "YD")
                    {
                        _sw = new StreamWriter(path + ".22000-국민행복(별지)", true, _encoding);
                        _sw.WriteLine(_strLine + "0042122");
                    }
                    //2019.11.20 수정
                    else if (strAdd_check == "YE01")
                    {
                        _sw = new StreamWriter(path + ".26000-화물복지(별지)", true, _encoding);
                        _sw.WriteLine(_strLine + "0042126");
                    }
                    //2019.11.20 수정
                    else if (strAdd_check == "YE02")
                    {
                        _sw = new StreamWriter(path + ".27000-큰수레화물복지(별지)", true, _encoding);
                        _sw.WriteLine(_strLine + "0042127");
                    }
                    else
                    {
                        switch (ConvertStrCode(_strCode))
                        {
                            case "SFC":
                                _sw = new StreamWriter(path + ".2000-SFC", true, _encoding);
                                _sw.WriteLine(_strLine + "0042102");
                                break;
                            case "LIFE":
                                _sw = new StreamWriter(path + ".3000-Life", true, _encoding);
                                _sw.WriteLine(_strLine + "0042103");
                                break;
                            case "JACHE":
                                _sw = new StreamWriter(path + ".4000-JACHE", true, _encoding);
                                _sw.WriteLine(_strLine + "0042104");
                                break;
                            case "CHECK":
                                _sw = new StreamWriter(path + ".5000-check", true, _encoding);
                                _sw.WriteLine(_strLine + "0042105");
                                break;
                            case "GSGS":
                                _sw = new StreamWriter(path + ".6000-GSGS", true, _encoding);
                                _sw.WriteLine(_strLine + "0042106");
                                break;
                            case "SOIL":
                                _sw = new StreamWriter(path + ".7000-SOIL", true, _encoding);
                                _sw.WriteLine(_strLine + "0042107");
                                break;
                            case "CAEND":
                                _sw = new StreamWriter(path + ".8000-화재", true, _encoding);
                                _sw.WriteLine(_strLine + "0042108");
                                break;
                            case "CMA":
                                _sw = new StreamWriter(path + ".9000-CMA", true, _encoding);
                                _sw.WriteLine(_strLine + "0042109");
                                break;
                            case "TOUR":
                                _sw = new StreamWriter(path + ".10000-TOUR", true, _encoding);
                                _sw.WriteLine(_strLine + "0042110");
                                break;
                            case "CJONE":
                                _sw = new StreamWriter(path + ".11000-CJONE", true, _encoding);
                                _sw.WriteLine(_strLine + "0042111");
                                break;
                            case "SK":
                                _sw = new StreamWriter(path + ".12000-SK", true, _encoding);
                                _sw.WriteLine(_strLine + "0042112");
                                break;
                            case "HOME_P":
                                _sw = new StreamWriter(path + ".13000-HOME_P", true, _encoding);
                                _sw.WriteLine(_strLine + "0042113");
                                break;
                            //2012.11.13 태희철 추가 6/6+카드 14000차
                            case "6_Plus":
                                _sw = new StreamWriter(path + ".14000-6_Plus", true, _encoding);
                                _sw.WriteLine(_strLine + "0042114");
                                break;
                            //2013.01.15 태희철 추가 4카드S-OIL 15000차
                            case "4_SOIL":
                                _sw = new StreamWriter(path + ".15000-4_SOIL", true, _encoding);
                                _sw.WriteLine(_strLine + "0042115");
                                break;
                            //2013.05.24 태희철 추가 4카드S-OIL 15000차
                            case "SM2_MNO":
                                _sw = new StreamWriter(path + ".16000-SM2_MNO", true, _encoding);
                                _sw.WriteLine(_strLine + "0042116");
                                break;
                            case "뷰티":
                                _sw = new StreamWriter(path + ".17000-뷰티", true, _encoding);
                                _sw.WriteLine(_strLine + "0042117");
                                break;
                            case "전자랜드":
                                _sw = new StreamWriter(path + ".18000-전자랜드", true, _encoding);
                                _sw.WriteLine(_strLine + "0042118");
                                break;
                            case "해피포인트":
                                _sw = new StreamWriter(path + ".19000-해피포인트", true, _encoding);
                                _sw.WriteLine(_strLine + "0042119");
                                break;
                            case "S클래스":
                                _sw = new StreamWriter(path + ".20000-S클래스", true, _encoding);
                                _sw.WriteLine(_strLine + "0042120");
                                break;
                            case "손보사3종":
                                _sw = new StreamWriter(path + ".21000-손보사3종", true, _encoding);
                                _sw.WriteLine(_strLine + "0042121");
                                break;
                            //2019.11.20 상품이미지코드로 대체
                            case "화물복지":
                                _sw = new StreamWriter(path + ".26000-화물복지(별지)", true, _encoding);
                                _sw.WriteLine(_strLine + "0042126");
                                break;
                            //2019.11.20 상품이미지코드로 대체
                            case "큰수레화물복지":
                                _sw = new StreamWriter(path + ".27000-큰수레화물복지(별지)", true, _encoding);
                                _sw.WriteLine(_strLine + "0042127");
                                break;
                            //case "국민행복":
                            //    _sw = new StreamWriter(path + ".22000-국민행복", true, _encoding);
                            //    _sw.WriteLine(_strLine + "0042122");
                            //    break;
                            default:
                                _sw = new StreamWriter(path + ".DONG", true, _encoding);
                                _sw.WriteLine(_strLine + "0042101");
                                break;
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

        private static string ConvertStrCode(string _strCode)
        {
            string strReturn = null;
            switch (_strCode)
            {
                //2000-SFC
                #region SFC     
                case "13932":
                case "13935":
                case "13936":
                case "13937":
                case "13938":
                case "13939":
                case "13940":
                case "13941":
                case "13942":
                case "13943":
                case "13944":
                case "13945":
                case "13955":
                case "13957":
                case "13958":
                case "13959":
                case "13960":
                case "13961":
                case "13962":
                case "13963":
                case "13964":
                case "13965":
                case "13966":
                case "13968":
                case "13969":
                case "13970":
                case "13971":
                case "13972":
                case "13973":
                case "13974":
                case "13975":
                case "13976":
                case "13977":
                case "13978":
                case "13979":
                case "13980":
                case "13981":
                case "13982":
                case "13983":
                case "13984":
                case "13985":
                case "13986":
                case "13987":
                case "13988":
                case "13989":
                case "13990":
                case "13992":
                case "13993":
                case "13994":
                case "13995":
                case "13996":
                case "13997":
                case "13998":
                case "13999":
                case "14000":
                case "14001":
                case "14002":
                case "14003":
                case "14004":
                case "14005":
                case "14006":
                case "14007":
                case "14008":
                case "14009":
                case "14010":
                case "14011":
                case "14012":
                case "14013":
                case "14027":
                case "14127":
                case "14128":
                case "14129":
                case "14130":
                case "14131":
                case "14132":
                case "14133":
                case "14134":
                case "14135":
                case "14136":
                case "14137":
                case "14138":
                case "14139":
                case "14140":
                case "14141":
                case "14142":
                case "14143":
                case "14144":
                case "14145":
                case "14146":
                case "14147":
                case "14148":
                case "14149":
                case "14150":
                case "14151":
                case "14152":
                case "14153":
                case "14154":
                case "14155":
                case "14156":
                case "14157":
                case "14158":
                case "14159":
                case "14160":
                case "14161":
                case "14162":
                case "14163":
                case "14164":
                case "14165":
                case "14166":
                case "14167":
                case "14168":
                case "14169":
                case "14170":
                case "14171":
                case "14172":
                case "14173":
                case "14174":
                case "14175":
                case "14176":
                case "14177":
                case "14178":
                case "14179":
                case "14180":
                case "14181":
                case "14182":
                case "14183":
                case "14184":
                case "14185":
                case "14186":
                case "14187":
                case "14188":
                case "14189":
                case "14190":
                case "14197":
                case "14198":
                case "14208":
                case "14232":
                case "14233":
                case "14234":
                case "14235":
                case "14236":
                case "14237":
                case "14238":
                case "14239":
                case "14240":
                case "14241":
                case "14242":
                case "14243":
                case "14244":
                case "14245":
                case "14249":
                case "14261":
                case "14280":
                case "14281":
                case "14282":
                case "14283":
                case "14284":
                case "14285":
                case "14286":
                case "14287":
                case "14356":
                case "14412":
                case "14414":
                case "14416":
                case "14417":
                case "14421":
                case "14422":
                case "14423":
                case "14424":
                case "14436":
                case "14437":
                case "14438":
                case "14439":
                case "14457":
                case "14458":
                case "14671":
                case "14672":
                case "14673":
                case "14674":
                case "14742":
                case "14743":
                case "14744":
                case "14745":
                case "15035":
                case "15036":
                case "15037":
                case "15038":
                case "15043":
                case "15044":
                case "15094":
                case "15095":
                case "15096":
                case "15097":
                case "15245":
                case "15447":
                case "15582":
                case "15627":
                case "15628":
                case "15629":
                case "15630":
                case "15631":
                case "15632":
                case "15633":
                case "15634":
                case "15635":
                case "15636":
                case "15637":
                case "15638":
                case "15639":
                case "15640":
                case "15641":
                case "15642":
                case "15643":
                case "15644":
                case "15645":
                case "15646":
                case "15647":
                case "15648":
                case "15649":
                case "15650":
                case "15651":
                case "15652":
                case "15653":
                case "15654":
                case "15655":
                case "15656":
                case "15657":
                case "15658":
                case "15659":
                case "15660":
                case "15661":
                case "15662":
                case "15680":
                case "15681":
                case "15682":
                case "15683":
                case "15684":
                case "15685":
                case "15686":
                case "15687":
                case "15688":
                case "15689":
                case "15690":
                case "15691":
                case "15712":
                case "15713":
                case "15714":
                case "15715":
                case "16580":
                case "16581":
                case "16582":
                case "16583":
                case "16584":
                case "16585":
                case "16586":
                case "16587":
                case "16588":
                case "16589":
                case "16590":
                case "16591":
                case "16592":
                case "16593":
                case "16594":
                case "16595":
                case "16596":
                case "16597":
                case "16598":
                case "16599":
                case "16766":
                case "16767":
                case "16768":
                case "16769":
                case "16770":
                case "16771":
                case "16772":
                case "16773":
                case "16859":
                //2013.08.16 태희철 추가
                case "17181":
                case "17182":
                case "17183":
                case "17184":
                case "17185":
                case "17186":
                case "17976":
                case "17977":
                    strReturn = "SFC";
                    break;
                #endregion
                //3000-LIFE
                case "15839":
                case "15840":
                case "15845":
                case "15846":
                case "15849":
                case "15850":
                case "15851":
                case "15852":
                case "15853":
                case "15854":
                case "15855":
                case "15856":
                case "15857":
                case "15858":
                    strReturn = "LIFE";
                    break;
                //4000-JACHE
                case "10000":
                case "10963":
                case "10966":
                case "10977":
                    strReturn = "JACHE";
                    break;
                //5000-CHECK
                #region 5000-CHECK     
                case "10139":
                case "10140":
                case "10141":
                case "10143":
                case "10144":
                case "10145":
                case "12385":
                case "12386":
                case "12387":
                case "12388":
                case "12389":
                case "12390":
                case "12391":
                case "12392":
                case "12393":
                case "12394":
                case "12395":
                case "12396":
                case "12397":
                case "12398":
                case "12399":
                case "12400":
                case "12404":
                case "12405":
                case "12406":
                case "12407":
                case "12408":
                case "12409":
                case "12412":
                case "12413":
                case "13553":
                case "13595":
                case "13596":
                case "13635":
                case "13701":
                case "13702":
                case "13754":
                case "13758":
                case "13769":
                case "13776":
                case "13777":
                case "13780":
                case "13791":
                case "13792":
                case "13796":
                case "13797":
                case "13798":
                case "13799":
                case "13805":
                case "13806":
                case "13832":
                case "13901":
                case "13902":
                case "13903":
                case "13923":
                case "14205":
                case "14206":
                case "14219":
                case "14220":
                case "14221":
                case "14222":
                case "14394":
                case "14428":
                case "14429":
                case "14431":
                case "14478":
                case "14485":
                case "14486":
                case "14490":
                case "14532":
                case "14533":
                case "14535":
                case "14536":
                case "14537":
                case "14538":
                case "14539":
                case "14540":
                case "14575":
                case "14576":
                case "14635":
                case "14636":
                case "14645":
                case "14646":
                case "14651":
                case "14652":
                case "14667":
                case "14668":
                case "14684":
                case "14686":
                case "14737":
                case "14740"://2013.04.04 태희철 추가
                case "14741"://2013.04.04 태희철 추가
                case "14746":
                case "14789":
                case "14804":
                case "14829":
                case "14849":
                case "14868":
                case "14875":
                case "14876":
                case "14877":
                case "14909":
                case "14917":
                case "14918":
                case "14925":
                case "14926":
                case "14946":
                case "14947":
                case "14952":
                case "14980":
                case "15052":
                case "15105":
                case "15106":
                case "15107":
                case "15108":
                case "15109":
                case "15110":
                case "15111":
                case "15112":
                case "15150":
                case "15170":
                case "15171":
                case "15195":
                case "15288"://2013.04.04 태희철 추가
                case "15289"://2013.04.04 태희철 추가
                case "15290"://2013.04.04 태희철 추가
                case "15315":
                case "15338":
                case "15339":
                case "15340":
                case "15341":
                case "15342":
                case "15343":
                case "15376":
                case "15377":
                case "15378":
                case "15379":
                case "15380":
                case "15390":
                case "15391":
                case "15392":
                case "15393":
                case "15394":
                case "15397":
                case "15405":
                case "15406":
                case "15407":
                case "15417":
                case "15418":
                case "15419":
                case "15423":
                case "15424":
                case "15445":
                case "15446":
                case "15460"://2013.04.04 태희철 추가
                case "15465":
                case "15529":
                case "15530":
                case "15550":
                case "15551":
                case "15552":
                case "15567":
                case "15568":
                case "15569":
                case "15570"://2013.04.04 태희철 추가
                case "15572":
                case "15591"://2013.04.04 태희철 추가
                case "15592"://2013.04.04 태희철 추가
                case "15593"://2013.04.04 태희철 추가
                case "15595":
                case "15669":
                case "15735"://2013.04.04 태희철 추가
                case "15867"://2013.04.04 태희철 추가
                case "15868"://2013.04.04 태희철 추가
                case "15869"://2013.04.04 태희철 추가
                case "15870"://2013.04.04 태희철 추가
                case "15909":
                case "16142"://2013.04.04 태희철 추가
                case "17059"://2013.04.04 태희철 추가
                case "17086"://2013.04.04 태희철 추가
                case "17087"://2013.04.04 태희철 추가
                case "16888":
                case "16721":
                case "16722":
                    strReturn = "CHECK";
                    break;
                #endregion
                //6000-GSGS
                #region 6000-GSGS 
                case "14834":
                case "14835":
                case "14862":
                case "14863":
                case "14896":
                case "14897":
                case "14898":
                case "14899":
                case "14920":
                case "14921":
                case "14950":
                case "14951":
                case "14972":
                case "14973":
                case "14974":
                case "14975":
                case "15011":
                case "15027":
                case "15030":
                case "15046":
                case "15047":
                case "15113":
                case "15114":
                case "15115":
                case "15116":
                case "15117":
                case "15118":
                case "15119":
                case "15120":
                case "15121":
                case "15126":
                case "15127":
                case "15146":
                case "15147":
                case "15148":
                case "15149":
                case "15151":
                case "15152":
                case "15161":
                case "15162":
                case "15240":
                case "15241":
                case "15242":
                case "15246":
                case "15247":
                case "15248":
                case "15275":
                case "15276":
                case "15277":
                case "15278":
                case "15286":
                case "15287":
                case "15336":
                case "15506":
                case "15507":
                case "15543":
                case "15544":
                case "15573":
                case "15574":
                case "15589":
                case "15590":
                case "15696":
                case "15697":
                case "15698":
                case "15699":
                case "16611":
                case "16612":
                    strReturn = "GSGS";
                    break;
                #endregion
                //7000-SOIL
                #region 7000-SOIL     
                case "14969":
                case "14970":
                case "14976":
                case "14977":
                case "14978":
                case "14979":
                case "15019":
                case "15020":
                case "15028":
                case "15029":
                case "15057":
                case "15058":
                case "15061":
                case "15062":
                case "15063":
                case "15064":
                case "15065":
                case "15066":
                case "15067":
                case "15068":
                case "15129":
                case "15130":
                case "15132":
                case "15133":
                case "15153":
                case "15154":
                case "15163":
                case "15164":
                case "15182":
                case "15183":
                case "15188":
                case "15189":
                case "15191":
                case "15192":
                case "15193":
                case "15194":
                case "15222":
                case "15223":
                case "15224":
                case "15225":
                case "15232":
                case "15236":
                case "15237":
                case "15238":
                case "15249":
                case "15250":
                case "15251":
                case "15253":
                case "15254":
                case "15255":
                case "15256":
                case "15257":
                case "15258":
                case "15273":
                case "15274":
                case "15302":
                case "15303":
                case "15304":
                case "15305":
                case "15337":
                case "15433":
                case "15434":
                case "15435":
                case "15436":
                case "15461":
                case "15462":
                case "15480":
                case "15481":
                case "15575":
                case "15576":
                case "15583":
                case "15584":
                //case "15705": -> 손보사3종으로 이동
                //case "15706": 2013.11.11 태희철 수정
                case "15885":
                case "15886":
                case "15945":
                case "15946":
                case "16059":
                case "16060":
                case "16061":
                case "16062":
                case "16098":
                case "16099":
                case "16100":
                case "16101":
                case "16110":
                case "16111":
                case "16112":
                case "16113":
                case "16116":
                case "16117":
                    strReturn = "SOIL";
                    break;
                #endregion
                //8000-caend
                #region 8000-CAEND
                case "16103":
                case "16104":
                case "16363":
                case "16364":
                case "16365":
                case "16366":
                case "15721":
                case "15722":
                case "16420":
                case "16421":
                case "16422":
                case "16423":
                case "16437":
                case "16438":
                case "16439":
                case "16440":
                case "16572":
                case "16573":
                case "16574":
                case "16575":
                case "16650":
                case "16651":
                case "16659":
                case "16660":
                case "16652":
                case "16653":
                    strReturn = "CAEND";
                    break;
                #endregion
                //9000-CMA
                case "15425":
                case "15512":
                case "15513":
                case "15670":
                case "15671":
                    strReturn = "CMA";
                    break;
                //10000-TOUR
                case "16328":
                case "16329":
                    strReturn = "TOUR";
                    break;
                //11000-CJONE
                case "16334":
                case "16335":
                case "16119":
                case "16120":
                case "16151":
                case "16152":
                case "17167":   //2013.05.20 태희철 추가
                case "17168":   //2013.05.20 태희철 추가
                    strReturn = "CJONE";
                    break;
                //12000-SK
                case "16782":
                case "16783":
                case "16784":
                case "16785":
                    strReturn = "SK";
                    break;
                //13000-HOMEP
                case "17046":
                case "17047":
                case "17048":
                case "17049":
                case "17050":
                case "17051":
                case "17124":       //2013.04.01 태희철 추가
                case "17125":
                    strReturn = "HOME_P";
                    break;
                //DEAR
                case "15366":
                case "15367":
                case "15371":
                case "15381":
                    strReturn = "DEAR";
                    break;
                //2012.11.13 태희철 추가
                //6/6+
                case "16924":
                case "16925":
                case "16926":
                case "16927":
                case "16928":
                case "16929":
                case "16930":
                case "16931":
                case "16932":
                case "16933":
                case "16934":
                case "16935":
                case "16936":
                case "16937":
                case "17110":
                case "17112":
                    strReturn = "6_Plus";
                    break;
                case "17101":       //2013.01.07 태희철 추가
                case "17102":
                case "17103":
                case "17104":
                    strReturn = "4_SOIL";
                    break;
                case "17202":       //2013.05.20 태희철 추가
                case "17203":
                case "17204":
                case "17205":
                case "17297":
                case "17298":
                    strReturn = "SM2_MNO";
                    break;
                case "17266":       //2013.07.22 뷰티 신규 추가
                case "17267":
                case "17295":
                case "17296":
                    strReturn = "뷰티";
                    break;
                case "17299":       //2013.07.24 전자랜드 신규 추가
                case "17300":
                    strReturn = "전자랜드";
                    break;
                case "17236":       //2013.08.22 해피포인트 신규 추가
                    strReturn = "해피포인트";
                    break;
                case "16714":       //2013.10.17 S클래스 신규 추가
                case "16715":
                case "16716":
                case "16717":
                    strReturn = "S클래스";
                    break;
                case "15705":
                case "15706":
                case "15736":
                case "15737":
                case "15925":
                case "15926":
                case "15932":
                case "15933":
                case "15956":
                case "15957":
                case "16047":
                case "16048":
                case "16053":
                case "16054":
                case "16188":
                case "16189":
                case "17316":
                case "17317":
                    strReturn = "손보사3종";
                    break;
                case "17736":       //2015.05.06 국민행복 신규 추가
                case "17737":
                case "17740":
                    strReturn = "국민행복";
                    break;
                case "18190":       //2016.12.26 화물복지 신규 추가
                case "18191":
                case "18194":
                case "18195":
                case "18196":       //2019.10.08 화물복지 코드 추가
                    strReturn = "화물복지";
                    break;
                case "18192":       //2016.12.26 큰수레화물복지 신규 추가
                case "18193":
                    strReturn = "큰수레화물복지";
                    break;
                default:
                    strReturn = "";
                    break;
            }
            return strReturn;
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

        //private static string[] StringSplit(ref string data)
        //{
        //    int index = data.IndexOf(":");
        //    string content = data.Substring(index + 1);

        //    data = data.Substring(0, index);

        //    string[] returnValue = content.Split(",".ToCharArray());
        //    return returnValue;
        //}
       
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

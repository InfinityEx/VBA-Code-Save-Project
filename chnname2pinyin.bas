Function pinyin(p As String) As String

i = Asc(p)

Select Case i

Case -20319 To -20318: pinyin = "a"

Case -20317 To -20305: pinyin = "ai"

Case -20304 To -20296: pinyin = "an"

Case -20295 To -20293: pinyin = "ang"

Case -20292 To -20284: pinyin = "ao"

Case -20283 To -20266: pinyin = "ba"

Case -20265 To -20258: pinyin = "bai"

Case -20257 To -20243: pinyin = "ban"

Case -20242 To -20231: pinyin = "bang"

Case -20230 To -20052: pinyin = "bao"

Case -20051 To -20037: pinyin = "bei"

Case -20036 To -20033: pinyin = "ben"

Case -20032 To -20027: pinyin = "beng"

Case -20026 To -20003: pinyin = "bi"

Case -20002 To -19991: pinyin = "bian"

Case -19990 To -19987: pinyin = "biao"

Case -19986 To -19983: pinyin = "bie"

Case -19982 To -19977: pinyin = "bin"

Case -19976 To -19806: pinyin = "bing"

Case -19805 To -19786: pinyin = "bo"

Case -19785 To -19776: pinyin = "bu"

Case -19775 To -19775: pinyin = "ca"

Case -19774 To -19764: pinyin = "cai"

Case -19763 To -19757: pinyin = "can"

Case -19756 To -19752: pinyin = "cang"

Case -19751 To -19747: pinyin = "cao"

Case -19746 To -19742: pinyin = "ce"

Case -19741 To -19740: pinyin = "ceng"

Case -19739 To -19729: pinyin = "cha"

Case -19728 To -19726: pinyin = "chai"

Case -19725 To -19716: pinyin = "chan"

Case -19715 To -19541: pinyin = "chang"

Case -19540 To -19532: pinyin = "chao"

Case -19531 To -19526: pinyin = "che"

Case -19525 To -19516: pinyin = "chen"

Case -19515 To -19501: pinyin = "cheng"

Case -19500 To -19485: pinyin = "chi"

Case -19484 To -19480: pinyin = "chong"

Case -19479 To -19468: pinyin = "chou"

Case -19467 To -19290: pinyin = "chu"

Case -19289 To -19289: pinyin = "chuai"

Case -19288 To -19282: pinyin = "chuan"

Case -19281 To -19276: pinyin = "chuang"

Case -19275 To -19271: pinyin = "chui"

Case -19270 To -19264: pinyin = "chun"

Case -19263 To -19262: pinyin = "chuo"

Case -19261 To -19250: pinyin = "ci"

Case -19249 To -19244: pinyin = "cong"

Case -19243 To -19243: pinyin = "cou"

Case -19242 To -19239: pinyin = "cu"

Case -19238 To -19236: pinyin = "cuan"

Case -19235 To -19228: pinyin = "cui"

Case -19227 To -19225: pinyin = "cun"

Case -19224 To -19219: pinyin = "cuo"

Case -19218 To -19213: pinyin = "da"

Case -19212 To -19039: pinyin = "dai"

Case -19038 To -19024: pinyin = "dan"

Case -19023 To -19019: pinyin = "dang"

Case -19018 To -19007: pinyin = "dao"

Case -19006 To -19004: pinyin = "de"

Case -19003 To -18997: pinyin = "deng"

Case -18996 To -18978: pinyin = "di"

Case -18977 To -18962: pinyin = "dian"

Case -18961 To -18953: pinyin = "diao"

Case -18952 To -18784: pinyin = "die"

Case -18783 To -18775: pinyin = "ding"

Case -18774 To -18774: pinyin = "diu"

Case -18773 To -18527: pinyin = "dong"

Case -18526 To -18519: pinyin = "fa"

Case -18518 To -18502: pinyin = "fan"

Case -18501 To -18491: pinyin = "fang"

Case -18490 To -18479: pinyin = "fei"

Case -18478 To -18464: pinyin = "fen"

Case -18463 To -18449: pinyin = "feng"

Case -18448 To -18448: pinyin = "fo"

Case -18447 To -18447: pinyin = "fou"

Case -18446 To -18240: pinyin = "fu"

Case -18239 To -18238: pinyin = "ga"

Case -18237 To -18232: pinyin = "gai"

Case -18231 To -18221: pinyin = "gan"

Case -18220 To -18212: pinyin = "gang"

Case -18211 To -18202: pinyin = "gao"

Case -18201 To -18185: pinyin = "ge"

Case -18184 To -18184: pinyin = "gei"

Case -18183 To -18182: pinyin = "gen"

Case -18181 To -18013: pinyin = "geng"

Case -18012 To -17998: pinyin = "gong"

Case -17997 To -17989: pinyin = "gou"

Case -17988 To -17971: pinyin = "gu"

Case -17970 To -17965: pinyin = "gua"

Case -17964 To -17962: pinyin = "guai"

Case -17961 To -17951: pinyin = "guan"

Case -17950 To -17948: pinyin = "guang"

Case -17947 To -17932: pinyin = "gui"

Case -17931 To -17929: pinyin = "gun"

Case -17928 To -17923: pinyin = "guo"

Case -17922 To -17760: pinyin = "ha"

Case -17759 To -17753: pinyin = "hai"

Case -17752 To -17734: pinyin = "han"

Case -17733 To -17731: pinyin = "hang"

Case -17730 To -17722: pinyin = "hao"

Case -17721 To -17704: pinyin = "he"

Case -17703 To -17702: pinyin = "hei"

Case -17701 To -17698: pinyin = "hen"

Case -17697 To -17693: pinyin = "heng"

Case -17692 To -17684: pinyin = "hong"

Case -17683 To -17677: pinyin = "hou"

Case -17676 To -17497: pinyin = "hu"

Case -17496 To -17488: pinyin = "hua"

Case -17487 To -17483: pinyin = "huai"

Case -17482 To -17469: pinyin = "huan"

Case -17468 To -17455: pinyin = "huang"

Case -17454 To -17434, -5427: pinyin = "hui"

Case -17433 To -17428: pinyin = "hun"

Case -17427 To -17418: pinyin = "huo"

Case -17417 To -17203: pinyin = "ji"

Case -17202 To -17186: pinyin = "jia"

Case -17185 To -16984, -6721: pinyin = "jian"

Case -16983 To -16971: pinyin = "jiang"

Case -16970 To -16943, -6481: pinyin = "jiao"

Case -16942 To -16916, -3423: pinyin = "jie"

Case -16915 To -16734: pinyin = "jin"

Case -16733 To -16709: pinyin = "jing"

Case -16708 To -16707: pinyin = "jiong"

Case -16706 To -16690: pinyin = "jiu"

Case -16689 To -16665: pinyin = "ju"

Case -16664 To -16658: pinyin = "juan"

Case -16657 To -16648: pinyin = "jue"

Case -16647 To -16475: pinyin = "jun"

Case -16474 To -16471: pinyin = "ka"

Case -16470 To -16466: pinyin = "kai"

Case -16465 To -16460: pinyin = "kan"

Case -16459 To -16453: pinyin = "kang"

Case -16452 To -16449: pinyin = "kao"

Case -16448 To -16434: pinyin = "ke"

Case -16433 To -16430: pinyin = "ken"

Case -16429 To -16428: pinyin = "keng"

Case -16427 To -16424: pinyin = "kong"

Case -16423 To -16420: pinyin = "kou"

Case -16419 To -16413: pinyin = "ku"

Case -16412 To -16408: pinyin = "kua"

Case -16407 To -16404: pinyin = "kuai"

Case -16403 To -16402: pinyin = "kuan"

Case -16401 To -16394: pinyin = "kuang"

Case -16393 To -16221: pinyin = "kui"

Case -16220 To -16217: pinyin = "kun"

Case -16216 To -16213: pinyin = "kuo"

Case -16212 To -16206: pinyin = "la"

Case -16205 To -16203: pinyin = "lai"

Case -16202 To -16188: pinyin = "lan"

Case -16187 To -16181: pinyin = "lang"

Case -16180 To -16172: pinyin = "lao"

Case -16171 To -16170: pinyin = "le"

Case -16169 To -16159: pinyin = "lei"

Case -16158 To -16156: pinyin = "leng"

Case -16155 To -15960: pinyin = "li"

Case -15959 To -15959: pinyin = "lia"

Case -15958 To -15945: pinyin = "lian"

Case -15944 To -15934: pinyin = "liang"

Case -15933 To -15921: pinyin = "liao"

Case -15920 To -15916: pinyin = "lie"

Case -15915 To -15904: pinyin = "lin"

Case -15903 To -15890: pinyin = "ling"

Case -15889 To -15879: pinyin = "liu"

Case -15878 To -15708: pinyin = "long"

Case -15707 To -15702: pinyin = "lou"

Case -15701 To -15682: pinyin = "lu"

Case -15681 To -15668: pinyin = "lv"

Case -15667 To -15662: pinyin = "luan"

Case -15661 To -15660: pinyin = "lue"

Case -15659 To -15653: pinyin = "lun"

Case -15652 To -15641: pinyin = "luo"

Case -15640 To -15632: pinyin = "ma"

Case -15631 To -15626: pinyin = "mai"

Case -15625 To -15455: pinyin = "man"

Case -15454 To -15449: pinyin = "mang"

Case -15448 To -15437: pinyin = "mao"

Case -15436 To -15436: pinyin = "me"

Case -15435 To -15420: pinyin = "mei"

Case -15419 To -15417: pinyin = "men"

Case -15416 To -15409: pinyin = "meng"

Case -15408 To -15395: pinyin = "mi"

Case -15394 To -15386: pinyin = "mian"

Case -15385 To -15378: pinyin = "miao"

Case -15377 To -15376: pinyin = "mie"

Case -15375 To -15370: pinyin = "min"

Case -15369 To -15364: pinyin = "ming"

Case -15363 To -15363: pinyin = "miu"

Case -15362 To -15184: pinyin = "mo"

Case -15183 To -15181: pinyin = "mou"

Case -15180 To -15166: pinyin = "mu"

Case -15165 To -15159: pinyin = "na"

Case -15158 To -15154: pinyin = "nai"

Case -15153 To -15151: pinyin = "nan"

Case -15150 To -15150: pinyin = "nang"

Case -15149 To -15145: pinyin = "nao"

Case -15144 To -15144: pinyin = "ne"

Case -15143 To -15142: pinyin = "nei"

Case -15141 To -15141: pinyin = "nen"

Case -15140 To -15140: pinyin = "neng"

Case -15139 To -15129: pinyin = "ni"

Case -15128 To -15122: pinyin = "nian"

Case -15121 To -15120: pinyin = "niang"

Case -15119 To -15118: pinyin = "niao"

Case -15117 To -15111: pinyin = "nie"

Case -15110 To -15110: pinyin = "nin"

Case -15109 To -14942: pinyin = "ning"

Case -14941 To -14938: pinyin = "niu"

Case -14937 To -14934: pinyin = "nong"

Case -14933 To -14931: pinyin = "nu"

Case -14930 To -14930: pinyin = "nv"

Case -14929 To -14929: pinyin = "nuan"

Case -14928 To -14927: pinyin = "nue"

Case -14926 To -14923: pinyin = "nuo"

Case -14922 To -14922: pinyin = "o"

Case -14921 To -14915: pinyin = "ou"

Case -14914 To -14909: pinyin = "pa"

Case -14908 To -14903: pinyin = "pai"

Case -14902 To -14895: pinyin = "pan"

Case -14894 To -14890: pinyin = "pang"

Case -14889 To -14883: pinyin = "pao"

Case -14882 To -14874: pinyin = "pei"

Case -14873 To -14872: pinyin = "pen"

Case -14871 To -14858: pinyin = "peng"

Case -14857 To -14679: pinyin = "pi"

Case -14678 To -14675: pinyin = "pian"

Case -14674 To -14671: pinyin = "piao"

Case -14670 To -14669: pinyin = "pie"

Case -14668 To -14664: pinyin = "pin"

Case -14663 To -14655: pinyin = "ping"

Case -14654 To -14646: pinyin = "po"

Case -14645 To -14631: pinyin = "pu"

Case -14630 To -14595: pinyin = "qi"

Case -14594 To -14430: pinyin = "qia"

Case -14429 To -14408: pinyin = "qian"

Case -14407 To -14400: pinyin = "qiang"

Case -14399 To -14385: pinyin = "qiao"

Case -14384 To -14380: pinyin = "qie"

Case -14379 To -14369: pinyin = "qin"

Case -14368 To -14356: pinyin = "qing"

Case -14355 To -14354: pinyin = "qiong"

Case -14353 To -14346: pinyin = "qiu"

Case -14345 To -14171, -2364: pinyin = "qu"

Case -14170 To -14160: pinyin = "quan"

Case -14159 To -14152: pinyin = "que"

Case -14151 To -14150: pinyin = "qun"

Case -14149 To -14146: pinyin = "ran"

Case -14145 To -14141: pinyin = "rang"

Case -14140 To -14138: pinyin = "rao"

Case -14137 To -14136: pinyin = "re"

Case -14135 To -14126: pinyin = "ren"

Case -14125 To -14124: pinyin = "reng"

Case -14123 To -14123: pinyin = "ri"

Case -14122 To -14113: pinyin = "rong"

Case -14112 To -14110: pinyin = "rou"

Case -14109 To -14100: pinyin = "ru"

Case -14099 To -14098: pinyin = "ruan"

Case -14097 To -14095: pinyin = "rui"

Case -14094 To -14093: pinyin = "run"

Case -14092 To -14091: pinyin = "ruo"

Case -14090 To -14088: pinyin = "sa"

Case -14087 To -14084: pinyin = "sai"

Case -14083 To -13918: pinyin = "san"

Case -13917 To -13915: pinyin = "sang"

Case -13914 To -13911: pinyin = "sao"

Case -13910 To -13908: pinyin = "se"

Case -13907 To -13907: pinyin = "sen"

Case -13906 To -13906: pinyin = "seng"

Case -13905 To -13897: pinyin = "sha"

Case -13896 To -13895: pinyin = "shai"

Case -13894 To -13879: pinyin = "shan"

Case -13878 To -13871: pinyin = "shang"

Case -13870 To -13860: pinyin = "shao"

Case -13859 To -13848: pinyin = "she"

Case -13847 To -13832, -9528: pinyin = "shen"

Case -13831 To -13659: pinyin = "sheng"

Case -13658 To -13612: pinyin = "shi"

Case -13611 To -13602: pinyin = "shou"

Case -13601 To -13407: pinyin = "shu"

Case -13406 To -13405: pinyin = "shua"

Case -13404 To -13401: pinyin = "shuai"

Case -13400 To -13399: pinyin = "shuan"

Case -13398 To -13396: pinyin = "shuang"

Case -13395 To -13392: pinyin = "shui"

Case -13391 To -13388: pinyin = "shun"

Case -13387 To -13384: pinyin = "shuo"

Case -13383 To -13368: pinyin = "si"

Case -13367 To -13360: pinyin = "song"

Case -13359 To -13357: pinyin = "sou"

Case -13356 To -13344: pinyin = "su"

Case -13343 To -13341: pinyin = "suan"

Case -13340 To -13330: pinyin = "sui"

Case -13329 To -13327: pinyin = "sun"

Case -13326 To -13319: pinyin = "suo"

Case -13318 To -13148: pinyin = "ta"

Case -13147 To -13139: pinyin = "tai"

Case -13138 To -13121: pinyin = "tan"

Case -13120 To -13108: pinyin = "tang"

Case -13107 To -13097: pinyin = "tao"

Case -13096 To -13096: pinyin = "te"

Case -13095 To -13092: pinyin = "teng"

Case -13091 To -13077: pinyin = "ti"

Case -13076 To -13069: pinyin = "tian"

Case -13068 To -13064: pinyin = "tiao"

Case -13063 To -13061: pinyin = "tie"

Case -13060 To -12889, -6461: pinyin = "ting"

Case -12888 To -12876: pinyin = "tong"

Case -12875 To -12872: pinyin = "tou"

Case -12871 To -12861: pinyin = "tu"

Case -12860 To -12859: pinyin = "tuan"

Case -12858 To -12853: pinyin = "tui"

Case -12852 To -12850: pinyin = "tun"

Case -12849 To -12839: pinyin = "tuo"

Case -12838 To -12832: pinyin = "wa"

Case -12831 To -12830: pinyin = "wai"

Case -12829 To -12813: pinyin = "wan"

Case -12812 To -12803: pinyin = "wang"

Case -12802 To -12608: pinyin = "wei"

Case -12607 To -12598: pinyin = "wen"

Case -12597 To -12595: pinyin = "weng"

Case -12594 To -12586: pinyin = "wo"

Case -12585 To -12557: pinyin = "wu"

Case -12556 To -12360: pinyin = "xi"

Case -12359 To -12347: pinyin = "xia"

Case -12346 To -12321: pinyin = "xian"

Case -12320 To -12301: pinyin = "xiang"

Case -12300 To -12121: pinyin = "xiao"

Case -12120 To -12100: pinyin = "xie"

Case -12099 To -12090: pinyin = "xin"

Case -12089 To -12075: pinyin = "xing"

Case -12074 To -12068: pinyin = "xiong"

Case -12067 To -12059: pinyin = "xiu"

Case -12058 To -12040: pinyin = "xu"

Case -12039 To -11868: pinyin = "xuan"

Case -11867 To -11862, -6682: pinyin = "xue"

Case -11861 To -11848: pinyin = "xun"

Case -11847 To -11832: pinyin = "ya"

Case -11831 To -11799, -5428, -9293: pinyin = "yan"

Case -11798 To -11782: pinyin = "yang"

Case -11781 To -11605: pinyin = "yao"

Case -11604 To -11590: pinyin = "ye"

Case -11589 To -11537, -4929: pinyin = "yi"

Case -11536 To -11359: pinyin = "yin"

Case -11358 To -11341, -6946: pinyin = "ying"

Case -11340 To -11340: pinyin = "yo"

Case -11339 To -11325: pinyin = "yong"

Case -11324 To -11304: pinyin = "you"

Case -11303 To -11098, -4902: pinyin = "yu"

Case -11097 To -11078, -6462: pinyin = "yuan"

Case -11077 To -11068: pinyin = "yue"

Case -11067 To -11056: pinyin = "yun"

Case -11055 To -11053: pinyin = "za"

Case -11052 To -11046: pinyin = "zai"

Case -11045 To -11042: pinyin = "zan"

Case -11041 To -11039: pinyin = "zang"

Case -11038 To -11025: pinyin = "zao"

Case -11024 To -11021: pinyin = "ze"

Case -11020 To -11020: pinyin = "zei"

Case -11019 To -11019: pinyin = "zen"

Case -11018 To -11015: pinyin = "zeng"

Case -11014 To -10839: pinyin = "zha"

Case -10838 To -10833: pinyin = "zhai"

Case -10832 To -10816: pinyin = "zhan"

Case -10815 To -10801, -5968: pinyin = "zhang"

Case -10800 To -10791, -4408: pinyin = "zhao"

Case -10790 To -10781: pinyin = "zhe"

Case -10780 To -10765: pinyin = "zhen"

Case -10764 To -10588: pinyin = "zheng"

Case -10587 To -10545: pinyin = "zhi"

Case -10544 To -10534: pinyin = "zhong"

Case -10533 To -10520: pinyin = "zhou"

Case -10519 To -10332: pinyin = "zhu"

Case -10331 To -10330: pinyin = "zhua"

Case -10329 To -10329: pinyin = "zhuai"

Case -10328 To -10323: pinyin = "zhuan"

Case -10322 To -10316: pinyin = "zhuang"

Case -10315 To -10310: pinyin = "zhui"

Case -10309 To -10308: pinyin = "zhun"

Case -10307 To -10297: pinyin = "zhuo"

Case -10296 To -10282: pinyin = "zi"

Case -10281 To -10275: pinyin = "zong"

Case -10274 To -10271: pinyin = "zou"

Case -10270 To -10263: pinyin = "zu"

Case -10262 To -10261: pinyin = "zuan"

Case -10260 To -10257: pinyin = "zui"

Case -10256 To -10255: pinyin = "zun"

Case -10254 To -10254: pinyin = "zuo"

Case Else: pinyin = p

End Select

End Function

Function getpy(str)

For i = 1 To Len(str)
    If Len(str) <= 3 Then
        If i = 1 Then
            getpy = getpy & pinyin(Mid(str, i, 1)) & " "
        Else
            getpy = getpy & pinyin(Mid(str, i, 1)) & ""
        End If
    ElseIf Len(str) > 3 Then
        If i = 2 Then
            getpy = getpy & pinyin(Mid(str, i, 1)) & " "
        Else
            getpy = getpy & pinyin(Mid(str, i, 1)) & ""
        End If
    End If

Next i

End Function
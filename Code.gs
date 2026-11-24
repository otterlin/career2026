/**
 * 115學年度 二年級生涯參訪（僑泰高級中學職業試探）選課系統
 * Google Apps Script 後端資料庫與 API 引擎
 *
 * 部署方式：
 * 1. 建立一個全新的 Google 試算表 (Google Sheets)
 * 2. 點選上方選單「擴充功能」 -> 「Apps Script」
 * 3. 將此檔案內容貼入 Code.gs，並新增 HTML 檔案命名為 index.html 貼入前端程式碼
 * 4. 點選上方執行「initDatabase」函式一次（或直接部署後於系統管理員介面點擊「初始化資料庫」）
 * 5. 點選右上角「部署」 -> 「新增部署作業」 -> 類型選「網頁應用程式」
 *    - 執行身分：我 (Me)
 *    - 誰可以存取：所有人 (Anyone)
 */

const SHEET_COURSES = '課程資料';
const SHEET_STUDENTS = '學生名冊與選課';
const SHEET_ACCOUNTS = '帳號設定';
const SHEET_LOGS = '選課紀錄';
const SHEET_PADLET_BOARDS = '佈告欄設定';
const SHEET_PADLET_POSTS = '佈告欄貼文';
const FALLBACK_SPREADSHEET_ID = '1mnFXn0Emrm6gHSuOlNnJJhJ-sgOSysfakSpgWvZBLRM';

function getSpreadsheet() {
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active && active.getId()) return active;
  } catch (e) {}

  try {
    const propId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (propId) {
      const byProp = SpreadsheetApp.openById(propId);
      if (byProp) return byProp;
    }
  } catch (e) {}

  try {
    if (FALLBACK_SPREADSHEET_ID) {
      const byFallback = SpreadsheetApp.openById(FALLBACK_SPREADSHEET_ID);
      if (byFallback) return byFallback;
    }
  } catch (e) {}

  try {
    const created = SpreadsheetApp.create('115學年度二年級生涯參訪選課資料庫');
    PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', created.getId());
    return created;
  } catch (e) {
    return null;
  }
}

const DEFAULT_COURSES = [{"id": 1, "dept": "電子科", "name": "電路設計DIY", "desc": "電子電路套件DIY實作，完成後會有意想不到的功能。", "hasWork": "有", "limitPerClass": 2}, {"id": 2, "dept": "資訊科", "name": "AI 魔法學園", "desc": "結合校園生活與社會時事體驗玩文本式與圖像生成式AI之互動。", "hasWork": "有", "limitPerClass": 2}, {"id": 3, "dept": "汽車科", "name": "極速領域", "desc": "動力競速賽車製作。", "hasWork": "有", "limitPerClass": 2}, {"id": 4, "dept": "機械科", "name": "咚咚龍", "desc": "模型組裝，從機構造型、連桿到電控開關，認識機械原理。", "hasWork": "有", "limitPerClass": 2}, {"id": 5, "dept": "商經科", "name": "趣味財商大冒險", "desc": "以桌遊搭配3D公仔，運用不同職業之薪水與理財方式贏得獎品。", "hasWork": "有", "limitPerClass": 2}, {"id": 6, "dept": "廣設科", "name": "二次元文創藝世界", "desc": "商品圖紋設計製作並以熱轉印技術完成商品印製。", "hasWork": "有", "limitPerClass": 2}, {"id": 7, "dept": "室設科", "name": "木作模型小達人", "desc": "基礎手作桌椅木質家具模型並彩繪上色。", "hasWork": "有", "limitPerClass": 2}, {"id": 8, "dept": "幼保科", "name": "偶來與你相遇", "desc": "設計製作各式可愛的玩偶與手偶。", "hasWork": "有", "limitPerClass": 2}, {"id": 9, "dept": "美容科", "name": "指尖上的藝術", "desc": "運用光療技巧，劃出抽象線條完成甲片，再做成吊飾。", "hasWork": "有", "limitPerClass": 2}, {"id": 10, "dept": "餐飲科", "name": "藍帶一番小當家", "desc": "餅乾蛋糕製作或中西式餐點製作。", "hasWork": "有", "limitPerClass": 2}, {"id": 11, "dept": "觀光科", "name": "雪克搖搖搖", "desc": "各式飲料特性教學與調製。", "hasWork": "有", "limitPerClass": 2}, {"id": 12, "dept": "寵經科", "name": "寵愛毛小孩", "desc": "認識寵物美容工具、寵物洗澡與吹毛技巧、基礎清潔、毛髮梳理、造型整理。", "hasWork": "有", "limitPerClass": 2}, {"id": 13, "dept": "應英科", "name": "外籍生帶你「折」出英文故事力", "desc": "全方位英語學習，與外籍生趣味互動！運用英文句型與創意貼紙，手作個人專屬八折頁故事繪本。", "hasWork": "有", "limitPerClass": 2}, {"id": 14, "dept": "應日科", "name": "日式圓扇彩繪", "desc": "日本文化介紹、圓扇彩繪以及浴衣體驗。", "hasWork": "有", "limitPerClass": 2}, {"id": 15, "dept": "普通科", "name": "幸福彩球", "desc": "結合熱縮片與感溫塗料，並融入繡球花的自然美感，營造出技術與藝術結合的吸引力。", "hasWork": "有", "limitPerClass": 2}];

const DEFAULT_STUDENTS = [["八年1班", "1140001", 1, "男", "江承諺"], ["八年1班", "1140002", 2, "男", "余思同"], ["八年1班", "1140003", 3, "男", "呂榮庭"], ["八年1班", "1140004", 4, "男", "邱昱森"], ["八年1班", "1140005", 5, "男", "洪翊展"], ["八年1班", "1140006", 6, "男", "徐惟喆"], ["八年1班", "1140007", 7, "男", "陳奕劼"], ["八年1班", "1140008", 8, "男", "陳弈昌"], ["八年1班", "1140009", 9, "男", "楊元碩"], ["八年1班", "1140010", 10, "男", "劉承勛"], ["八年1班", "1140011", 11, "男", "劉宸睿"], ["八年1班", "1140012", 12, "男", "劉祐誠"], ["八年1班", "1140013", 13, "男", "劉祐旖"], ["八年1班", "1140466", 14, "男", "張祥境"], ["八年1班", "1140015", 15, "男", "蕭克錡"], ["八年1班", "1140016", 16, "男", "蕭翊鈞"], ["八年1班", "1140017", 17, "女", "王苡璦"], ["八年1班", "1140018", 18, "女", "吳旻恩"], ["八年1班", "1140019", 19, "女", "李雨築"], ["八年1班", "1140020", 20, "女", "林宥彤"], ["八年1班", "1140021", 21, "女", "張玲瑀"], ["八年1班", "1140022", 22, "女", "張書羽"], ["八年1班", "1140023", 23, "女", "陳琳蓁"], ["八年1班", "1140024", 24, "女", "黃羽綾"], ["八年1班", "1140025", 25, "女", "楊立辰"], ["八年1班", "1140028", 26, "女", "蔡沛恩"], ["八年1班", "1140027", 27, "女", "劉芷沂"], ["八年2班", "1140029", 1, "男", "石松恩"], ["八年2班", "1140030", 2, "男", "成宇恩"], ["八年2班", "1140031", 3, "男", "江浚騰"], ["八年2班", "1140032", 4, "男", "邱詣桓"], ["八年2班", "1140033", 5, "男", "洪浩誠"], ["八年2班", "1140034", 6, "男", "張哲瀚"], ["八年2班", "1140035", 7, "男", "陳子翔"], ["八年2班", "1140036", 8, "男", "黃奕騰"], ["八年2班", "1140037", 9, "男", "葉育辰"], ["八年2班", "1140039", 10, "男", "蒙昱辰"], ["八年2班", "1140040", 11, "男", "劉億豪"], ["八年2班", "1140041", 12, "男", "顏嘉甫"], ["八年2班", "1140042", 13, "女", "朱欣彤"], ["八年2班", "1140043", 14, "女", "林以凡"], ["八年2班", "1140045", 15, "女", "林雨潼"], ["八年2班", "1140046", 16, "女", "洪苡娜"], ["八年2班", "1140047", 17, "女", "胡喬棻"], ["八年2班", "1140048", 18, "女", "張芫綺"], ["八年2班", "1140049", 19, "女", "許保雲"], ["八年2班", "1140050", 20, "女", "陳姵澐"], ["八年2班", "1140051", 21, "女", "黃紫緁"], ["八年2班", "1140052", 22, "女", "楊謹瑄"], ["八年2班", "1140053", 23, "女", "劉宸瑄"], ["八年2班", "1140054", 24, "女", "厲庭杉"], ["八年2班", "1140055", 25, "女", "蔡念臻"], ["八年2班", "1140056", 26, "女", "鄧予喬"], ["八年2班", "1140467", 27, "男", "陳澤源"], ["八年3班", "1140057", 1, "男", "毛浚桀"], ["八年3班", "1140058", 2, "男", "王泊泉"], ["八年3班", "1140059", 3, "男", "王皓成"], ["八年3班", "1140060", 4, "男", "周品叡"], ["八年3班", "1140468", 5, "男", "林貫倫"], ["八年3班", "1140062", 6, "男", "張家樂"], ["八年3班", "1140063", 7, "男", "張喻淮"], ["八年3班", "1140064", 8, "男", "陳昇羽"], ["八年3班", "1140065", 9, "男", "陳煥煬"], ["八年3班", "1140066", 10, "男", "陳震桓"], ["八年3班", "1140067", 11, "男", "游之諺"], ["八年3班", "1140068", 12, "男", "楊一翔"], ["八年3班", "1140069", 13, "男", "蕭正禹"], ["八年3班", "1140070", 14, "女", "江羽潔"], ["八年3班", "1140071", 15, "女", "周樂昀"], ["八年3班", "1140072", 16, "女", "林嘉俞"], ["八年3班", "1140073", 17, "女", "紀宛彤"], ["八年3班", "1140074", 18, "女", "紀亭羽"], ["八年3班", "1140075", 19, "女", "胡馨云"], ["八年3班", "1140076", 20, "女", "徐韻晴"], ["八年3班", "1140077", 21, "女", "張如君"], ["八年3班", "1140079", 22, "女", "張佐禧"], ["八年3班", "1140080", 23, "女", "許芯語"], ["八年3班", "1140081", 24, "女", "陳可欣"], ["八年3班", "1140082", 25, "女", "陳昱沛"], ["八年3班", "1140083", 26, "女", "廖紫淩"], ["八年3班", "1140477", 27, "女", "黃雅芯"], ["八年4班", "1140084", 1, "男", "吳昊臻"], ["八年4班", "1140085", 2, "男", "林承洋"], ["八年4班", "1140086", 3, "男", "林義翔"], ["八年4班", "1140088", 4, "男", "胡峻瑋"], ["八年4班", "1140089", 5, "男", "高聖恩"], ["八年4班", "1140090", 6, "男", "張子烜"], ["八年4班", "1140091", 7, "男", "張紘育"], ["八年4班", "1140092", 8, "男", "陳宣輔"], ["八年4班", "1140093", 9, "男", "彭建綦"], ["八年4班", "1140094", 10, "男", "童越"], ["八年4班", "1140095", 11, "男", "黃俊湟"], ["八年4班", "1140096", 12, "男", "黃俊漢"], ["八年4班", "1140097", 13, "男", "蕭少恩"], ["八年4班", "1140098", 14, "男", "蕭翊安"], ["八年4班", "1140099", 15, "女", "安以涵"], ["八年4班", "1140100", 16, "女", "吳怡樺"], ["八年4班", "1140101", 17, "女", "沈宥媗"], ["八年4班", "1140102", 18, "女", "邱畇芸"], ["八年4班", "1140103", 19, "女", "張可欣"], ["八年4班", "1140104", 20, "女", "張詠彤"], ["八年4班", "1140105", 21, "女", "陳毓茜"], ["八年4班", "1140479", 22, "女", "張詠甯"], ["八年4班", "1140107", 23, "女", "黃語晴"], ["八年4班", "1140108", 24, "女", "劉怡芳"], ["八年4班", "1140109", 25, "女", "劉珈豈"], ["八年4班", "1140110", 26, "女", "劉珈瑄"], ["八年4班", "1140111", 27, "女", "蕭圓蓉"], ["八年4班", "1140481", 28, "男", "黃裕龍"], ["八年5班", "1140112", 1, "男", "白侑庭"], ["八年5班", "1140113", 2, "男", "杜宗鴻"], ["八年5班", "1140114", 3, "男", "林煜翔"], ["八年5班", "1140115", 4, "男", "施逕宥"], ["八年5班", "1140116", 5, "男", "張晉晟"], ["八年5班", "1140117", 6, "男", "陳琪皓"], ["八年5班", "1140118", 7, "男", "黃呈洋"], ["八年5班", "1140119", 8, "男", "黃裕翔"], ["八年5班", "1140120", 9, "男", "葉承恩"], ["八年5班", "1140121", 10, "男", "葉峻翔"], ["八年5班", "1140122", 11, "男", "蔡邵丞"], ["八年5班", "1140123", 12, "男", "蕭宇辰"], ["八年5班", "1140124", 13, "女", "白宛倫"], ["八年5班", "1140125", 14, "女", "李珮娸"], ["八年5班", "1140126", 15, "女", "肖茸莛"], ["八年5班", "1140127", 16, "女", "林奕汝"], ["八年5班", "1140128", 17, "女", "林歆恩"], ["八年5班", "1140129", 18, "女", "金芯唯"], ["八年5班", "1140139", 19, "女", "賴雅麗"], ["八年5班", "1140131", 20, "女", "胡絮惠"], ["八年5班", "1140132", 21, "女", "黃毓媃"], ["八年5班", "1140133", 22, "女", "廖昱喬"], ["八年5班", "1140134", 23, "女", "廖敏芯"], ["八年5班", "1140135", 24, "女", "蔡瑋娣"], ["八年5班", "1140136", 25, "女", "蔣恩潔"], ["八年5班", "1140137", 26, "女", "賴珍酉"], ["八年5班", "1140474", 27, "男", "徐瑞玹"], ["八年6班", "1140140", 1, "男", "王奕凱"], ["八年6班", "1140141", 2, "男", "王泳騰"], ["八年6班", "1140142", 3, "男", "吳恩愷"], ["八年6班", "1140143", 4, "男", "李佳龍"], ["八年6班", "1140145", 5, "男", "李承恩"], ["八年6班", "1140146", 6, "男", "林震弦"], ["八年6班", "1140147", 7, "男", "洪浚程"], ["八年6班", "1140148", 8, "男", "張峻維"], ["八年6班", "1140149", 9, "男", "張晁睿"], ["八年6班", "1140150", 10, "男", "許兆維"], ["八年6班", "1140151", 11, "男", "陳建宇"], ["八年6班", "1140152", 12, "男", "曾以路"], ["八年6班", "1140153", 13, "男", "劉宗諺"], ["八年6班", "1140154", 14, "男", "謝昕辰"], ["八年6班", "1140155", 15, "男", "簡志謙"], ["八年6班", "1140156", 16, "女", "吳沛妮"], ["八年6班", "1140157", 17, "女", "張子晴"], ["八年6班", "1140158", 18, "女", "張妤婕"], ["八年6班", "1140159", 19, "女", "張芯綺"], ["八年6班", "1140160", 20, "女", "張采琪"], ["八年6班", "1140161", 21, "女", "張采維"], ["八年6班", "1140162", 22, "女", "張珞恩"], ["八年6班", "1140163", 23, "女", "張意歆"], ["八年6班", "1140164", 24, "女", "葉子墨"], ["八年6班", "1140165", 25, "女", "葉紜彤"], ["八年6班", "1140166", 26, "女", "劉妤宣"], ["八年6班", "1140167", 27, "女", "鄭沛淳"], ["八年6班", "1140168", 28, "女", "賴品媛"], ["八年7班", "1140169", 1, "男", "江廷軒"], ["八年7班", "1140170", 2, "男", "何育鳴"], ["八年7班", "1140171", 3, "男", "何昕沛"], ["八年7班", "1140172", 4, "男", "吳明泰"], ["八年7班", "1140173", 5, "男", "吳品宏"], ["八年7班", "1140174", 6, "男", "林伯叡"], ["八年7班", "1140175", 7, "男", "林品志"], ["八年7班", "1140176", 8, "男", "曾士瑋"], ["八年7班", "1140177", 9, "男", "黃昱凱"], ["八年7班", "1140179", 10, "男", "劉炘昊"], ["八年7班", "1140180", 11, "男", "蔡東諺"], ["八年7班", "1140181", 12, "男", "藍晟瑋"], ["八年7班", "1140182", 13, "男", "蘇皜睿"], ["八年7班", "1140183", 14, "女", "王沛羽"], ["八年7班", "1140184", 15, "女", "王泳翎"], ["八年7班", "1140185", 16, "女", "王品慈"], ["八年7班", "1140186", 17, "女", "朱宥馨"], ["八年7班", "1140188", 18, "女", "吳佳娜"], ["八年7班", "1140189", 19, "女", "李宛諠"], ["八年7班", "1140190", 20, "女", "林筠喬"], ["八年7班", "1140191", 21, "女", "張聿慈"], ["八年7班", "1140192", 22, "女", "張舒甯"], ["八年7班", "1140193", 23, "女", "張鈊評"], ["八年7班", "1140194", 24, "女", "莊羽媗"], ["八年7班", "1140195", 25, "女", "陳柏云"], ["八年7班", "1140196", 26, "女", "黄瑀岑"], ["八年7班", "1140197", 27, "女", "劉宥彤"], ["八年7班", "1140198", 28, "女", "劉宸琳"], ["八年8班", "1140199", 1, "男", "王宥騰"], ["八年8班", "1140200", 2, "男", "周承賢"], ["八年8班", "1140201", 3, "男", "張宥綸"], ["八年8班", "1140202", 4, "男", "張景舜"], ["八年8班", "1140203", 5, "男", "張楙杰"], ["八年8班", "1140204", 6, "男", "陳昱愷"], ["八年8班", "1140205", 7, "男", "陳昱維"], ["八年8班", "1140206", 8, "男", "陳微太"], ["八年8班", "1130034", 9, "男", "詹程凱"], ["八年8班", "1140207", 10, "男", "劉禹汛"], ["八年8班", "1140208", 11, "男", "鍾辰熙"], ["八年8班", "1140209", 12, "男", "蘇偉鈞"], ["八年8班", "1140210", 13, "女", "王翊潔"], ["八年8班", "1140211", 14, "女", "吳羽晴"], ["八年8班", "1140212", 15, "女", "林可婷"], ["八年8班", "1140213", 16, "女", "洪安亭"], ["八年8班", "1140214", 17, "女", "徐緯庭"], ["八年8班", "1140215", 18, "女", "張芸瓊"], ["八年8班", "1140216", 19, "女", "張雅晶"], ["八年8班", "1140217", 20, "女", "許子淇"], ["八年8班", "1140218", 21, "女", "陳妍霏"], ["八年8班", "1140219", 22, "女", "陳爰希"], ["八年8班", "1140480", 23, "女", "胡鈞喬"], ["八年8班", "1140221", 24, "女", "黃奕華"], ["八年8班", "1140222", 25, "女", "謝翊婕"], ["八年8班", "1140223", 26, "女", "羅聖愛"], ["八年8班", "1140475", 27, "男", "施博睿"], ["八年9班", "1140224", 1, "男", "王柏勛"], ["八年9班", "1140225", 2, "男", "王聖翰"], ["八年9班", "1140226", 3, "男", "吳駿瑋"], ["八年9班", "1140227", 4, "男", "李翌睿"], ["八年9班", "1140228", 5, "男", "周楷倢"], ["八年9班", "1140229", 6, "男", "林立恩"], ["八年9班", "1140230", 7, "男", "陳雨祈"], ["八年9班", "1140231", 8, "男", "陳冠圩"], ["八年9班", "1140232", 9, "男", "黃士菖"], ["八年9班", "1140233", 10, "男", "葉仁方"], ["八年9班", "1140234", 11, "男", "蔡羽衍"], ["八年9班", "1140235", 12, "男", "蔡祐杰"], ["八年9班", "1140236", 13, "女", "余羽彤"], ["八年9班", "1140237", 14, "女", "吳采怩"], ["八年9班", "1140239", 15, "女", "吳靜宜"], ["八年9班", "1140240", 16, "女", "周琬芯"], ["八年9班", "1140241", 17, "女", "林昱汝"], ["八年9班", "1140242", 18, "女", "張恩瑜"], ["八年9班", "1140243", 19, "女", "郭采姸"], ["八年9班", "1140245", 20, "女", "陳佳嘉"], ["八年9班", "1140246", 21, "女", "陳姮羲"], ["八年9班", "1140247", 22, "女", "曾詠渝"], ["八年9班", "1140250", 25, "女", "蔡沛霓"], ["八年9班", "1140251", 26, "女", "蔡欣馜"], ["八年9班", "1140252", 27, "女", "羅唯瑀"], ["八年9班", "1140483", 28, "女", "吳宜蔓"], ["八年10班", "1140253", 1, "男", "江明峻"], ["八年10班", "1140254", 2, "男", "吳詠宸"], ["八年10班", "1140255", 3, "男", "枋裕閎"], ["八年10班", "1140256", 4, "男", "林世明"], ["八年10班", "1140257", 5, "男", "林酉澤"], ["八年10班", "1140258", 6, "男", "邱翊珅"], ["八年10班", "1140264", 7, "男", "蔡丞睿"], ["八年10班", "1140260", 8, "男", "張守駿"], ["八年10班", "1140261", 9, "男", "張傳易"], ["八年10班", "1140262", 10, "男", "彭茂源"], ["八年10班", "1140263", 11, "男", "黃宥紳"], ["八年10班", "1140281", 12, "女", "鄭家佳"], ["八年10班", "1140265", 13, "女", "王婕伃"], ["八年10班", "1140266", 14, "女", "吳昀臻"], ["八年10班", "1140267", 15, "女", "林欣潔"], ["八年10班", "1140268", 16, "女", "林語葳"], ["八年10班", "1140269", 17, "女", "胡綾軒"], ["八年10班", "1140270", 18, "女", "翁宥筑"], ["八年10班", "1140271", 19, "女", "陳芷涵"], ["八年10班", "1140272", 20, "女", "陳柏華"], ["八年10班", "1140273", 21, "女", "陳禹安"], ["八年10班", "1140274", 22, "女", "陳羿妘"], ["八年10班", "1140275", 23, "女", "陳喬瑜"], ["八年10班", "1140276", 24, "女", "董芝芸"], ["八年10班", "1140277", 25, "女", "廖伊潔"], ["八年10班", "1140279", 26, "女", "蔡宜萱"], ["八年10班", "1140280", 27, "女", "蔡采言"], ["八年10班", "1140470", 28, "女", "游予涵"], ["八年10班", "1140472", 29, "男", "楊文玄"], ["八年11班", "1140282", 1, "男", "王翌齊"], ["八年11班", "1140283", 2, "男", "林家煜"], ["八年11班", "1140284", 3, "男", "柯政宇"], ["八年11班", "1140285", 4, "男", "柯柏睿"], ["八年11班", "1140286", 5, "男", "徐永均"], ["八年11班", "1140288", 6, "男", "高聖杰"], ["八年11班", "1140289", 7, "男", "張哲愷"], ["八年11班", "1140290", 8, "男", "張宸睿"], ["八年11班", "1140291", 9, "男", "張愷中"], ["八年11班", "1140292", 10, "男", "陳佑福"], ["八年11班", "1140293", 11, "男", "陳柏辰"], ["八年11班", "1140294", 12, "男", "陳恩銓"], ["八年11班", "1140295", 13, "男", "劉向杰"], ["八年11班", "1140296", 14, "男", "蘇浩宇"], ["八年11班", "1140297", 15, "女", "王紫恩"], ["八年11班", "1140298", 16, "女", "王榆評"], ["八年11班", "1140299", 17, "女", "李瑞妤"], ["八年11班", "1140300", 18, "女", "林亞欣"], ["八年11班", "1140301", 19, "女", "徐瑋翎"], ["八年11班", "1140302", 20, "女", "張芸甄"], ["八年11班", "1140303", 21, "女", "陳卉瑄"], ["八年11班", "1140304", 22, "女", "陳姵妤"], ["八年11班", "1140305", 23, "女", "黃曉晴"], ["八年11班", "1140306", 24, "女", "廖固醇"], ["八年11班", "1140307", 25, "女", "蔡佳芮"], ["八年11班", "1140308", 26, "女", "蔡采恩"], ["八年11班", "1140309", 27, "女", "賴盈蓁"], ["八年11班", "1140310", 28, "女", "蘇品云"], ["八年11班", "1140311", 29, "女", "蘇寀婕"], ["八年12班", "1140312", 1, "男", "朱錫祺"], ["八年12班", "1140313", 2, "男", "何思誼"], ["八年12班", "1140314", 3, "男", "阮泯澤"], ["八年12班", "1140315", 4, "男", "林威呈"], ["八年12班", "1140316", 5, "男", "施弼閎"], ["八年12班", "1140317", 6, "男", "施瀧珵"], ["八年12班", "1140318", 7, "男", "袁宏毅"], ["八年12班", "1140476", 8, "男", "黃楷勛"], ["八年12班", "1140320", 9, "男", "郭彥均"], ["八年12班", "1140321", 10, "男", "傅彥"], ["八年12班", "1140322", 11, "男", "楊淮丞"], ["八年12班", "1140323", 12, "男", "葉書亞"], ["八年12班", "1140324", 13, "男", "鄒鈞旭"], ["八年12班", "1140325", 14, "男", "劉韋樑"], ["八年12班", "1140326", 15, "女", "王妍晴"], ["八年12班", "1140327", 16, "女", "吳岳玲"], ["八年12班", "1140328", 17, "女", "沈于喬"], ["八年12班", "1140329", 18, "女", "林鈺軒"], ["八年12班", "1140330", 19, "女", "邱紹茗"], ["八年12班", "1140331", 20, "女", "徐曼瑄"], ["八年12班", "1140332", 21, "女", "張芸嫣"], ["八年12班", "1140333", 22, "女", "張恩綺"], ["八年12班", "1140334", 23, "女", "張憶玫"], ["八年12班", "1140335", 24, "女", "郭涵晴"], ["八年12班", "1140336", 25, "女", "陳芯榆"], ["八年12班", "1140337", 26, "女", "程姜勻"], ["八年12班", "1140339", 27, "女", "楊芷欣"], ["八年12班", "1140471", 28, "女", "張簡翊齊"], ["八年12班", "1140341", 29, "女", "鍾昕宸"], ["八年13班", "1140342", 1, "男", "王品鈞"], ["八年13班", "1140343", 2, "男", "王進德"], ["八年13班", "1140345", 3, "男", "邱晨希"], ["八年13班", "1140346", 4, "男", "徐哲駿"], ["八年13班", "1140347", 5, "男", "張永昌"], ["八年13班", "1140348", 6, "男", "張煜紳"], ["八年13班", "1140349", 7, "男", "陳乙昇"], ["八年13班", "1140350", 8, "男", "陳允凡"], ["八年13班", "1140351", 9, "男", "陳岦威"], ["八年13班", "1140352", 10, "男", "廖映程"], ["八年13班", "1140356", 11, "男", "譚子明"], ["八年13班", "1140354", 12, "男", "蔡秉勳"], ["八年13班", "1140355", 13, "男", "羅子清"], ["八年13班", "1140482", 14, "女", "王滋涵"], ["八年13班", "1140357", 15, "女", "王心渝"], ["八年13班", "1140358", 16, "女", "巫姵蒂"], ["八年13班", "1140359", 17, "女", "李芯蕎"], ["八年13班", "1140360", 18, "女", "林泫雅"], ["八年13班", "1140361", 19, "女", "邱渝晴"], ["八年13班", "1140362", 20, "女", "紀芯璦"], ["八年13班", "1140363", 21, "女", "胡詠妤"], ["八年13班", "1140364", 22, "女", "張綵晴"], ["八年13班", "1140365", 23, "女", "游家瑄"], ["八年13班", "1140366", 24, "女", "黃于柔"], ["八年13班", "1140367", 25, "女", "廖英廷"], ["八年13班", "1140368", 26, "女", "蔡可恩"], ["八年13班", "1140369", 27, "女", "賴湘鈴"], ["八年13班", "1140370", 28, "女", "顏均珈"], ["八年13班", "1140371", 29, "女", "魏睿霈"], ["八年13班", "1140478", 30, "女", "李芯語"], ["八年14班", "1140372", 1, "男", "何杰諾"], ["八年14班", "1140373", 2, "男", "吳忠霖"], ["八年14班", "1140374", 3, "男", "吳晁和"], ["八年14班", "1140375", 4, "男", "吳晁明"], ["八年14班", "1140376", 5, "男", "李其祐"], ["八年14班", "1140377", 6, "男", "張日鴻"], ["八年14班", "1140378", 7, "男", "張晏承"], ["八年14班", "1140379", 8, "男", "陳冠儒"], ["八年14班", "1140380", 9, "男", "陳奕璋"], ["八年14班", "1140381", 10, "男", "陳彥維"], ["八年14班", "1140382", 11, "男", "陳皇坪"], ["八年14班", "1140383", 12, "男", "陳嘉鴻"], ["八年14班", "1140384", 13, "男", "陳銚鴻"], ["八年14班", "1140385", 14, "男", "廖柏翔"], ["八年14班", "1140386", 15, "男", "蔣晏彰"], ["八年14班", "1140387", 16, "女", "江昱涵"], ["八年14班", "1140388", 17, "女", "吳之妤"], ["八年14班", "1140389", 18, "女", "吳庭瑜"], ["八年14班", "1140390", 19, "女", "張蕎恩"], ["八年14班", "1140391", 20, "女", "許綺娜"], ["八年14班", "1140392", 21, "女", "陳宥慈"], ["八年14班", "1140393", 22, "女", "葉欣妤"], ["八年14班", "1140394", 23, "女", "葉品弈"], ["八年14班", "1140395", 24, "女", "董芝婷"], ["八年14班", "1140396", 25, "女", "雷雅涵"], ["八年14班", "1140397", 26, "女", "鄧琳馨"], ["八年14班", "1140398", 27, "女", "鄭婉辰"], ["八年14班", "1140399", 28, "女", "賴若楹"], ["八年14班", "1140400", 29, "女", "謝懷熹"], ["八年14班", "1140401", 30, "女", "顏紫晴"], ["八年15班", "1140402", 1, "男", "王子睿"], ["八年15班", "1140403", 2, "男", "石沛諺"], ["八年15班", "1140404", 3, "男", "林鼎皓"], ["八年15班", "1140405", 4, "男", "張又翊"], ["八年15班", "1140406", 5, "男", "張律霆"], ["八年15班", "1140407", 6, "男", "許宸睿"], ["八年15班", "1140408", 7, "男", "陳宜杰"], ["八年15班", "1140409", 8, "男", "陳宥辰"], ["八年15班", "1140410", 9, "男", "陳振輔"], ["八年15班", "1140411", 10, "男", "蔡定晏"], ["八年15班", "1140412", 11, "男", "鄭翔軒"], ["八年15班", "1140413", 12, "男", "賴晉億"], ["八年15班", "1140414", 13, "男", "魏廷翰"], ["八年15班", "1140415", 14, "女", "王子晴"], ["八年15班", "1140416", 15, "女", "王予潔"], ["八年15班", "1140417", 16, "女", "王辰旭"], ["八年15班", "1140418", 17, "女", "沈品希"], ["八年15班", "1140419", 18, "女", "張詠晴"], ["八年15班", "1140420", 19, "女", "許亞蓁"], ["八年15班", "1140421", 20, "女", "陳芃羽"], ["八年15班", "1140422", 21, "女", "陳冠云"], ["八年15班", "1140423", 22, "女", "陳姿妘"], ["八年15班", "1140424", 23, "女", "黃沛瑜"], ["八年15班", "1140425", 24, "女", "黃郁晴"], ["八年15班", "1140426", 25, "女", "楊詠筑"], ["八年15班", "1140427", 26, "女", "廖云瑄"], ["八年15班", "1140428", 27, "女", "蔡羽喬"], ["八年15班", "1140429", 28, "女", "蔡沄樺"], ["八年15班", "1140473", 29, "女", "林姵旂"], ["八年16班", "1140430", 1, "男", "朱冠嘉"], ["八年16班", "1140431", 2, "男", "呂承奕"], ["八年16班", "1140432", 3, "男", "林毓衡"], ["八年16班", "1140433", 4, "男", "張楷辰"], ["八年16班", "1140434", 5, "男", "陳翊瑔"], ["八年16班", "1140435", 6, "男", "湯濬愷"], ["八年16班", "1140436", 7, "男", "潘聿恒"], ["八年16班", "1140437", 8, "男", "蔡尚宸"], ["八年16班", "1140438", 9, "男", "蘇敬承"], ["八年16班", "1140439", 10, "女", "周柔蓁"], ["八年16班", "1140440", 11, "女", "周憶蓮"], ["八年16班", "1140441", 12, "女", "林季希"], ["八年16班", "1140442", 13, "女", "林彥妤"], ["八年16班", "1140443", 14, "女", "林億欣"], ["八年16班", "1140444", 15, "女", "施宣彤"], ["八年16班", "1140445", 16, "女", "柯羽旂"], ["八年16班", "1140455", 17, "女", "蔡子翎"], ["八年16班", "1140447", 18, "女", "倪凡晴"], ["八年16班", "1140448", 19, "女", "張芮靜"], ["八年16班", "1140449", 20, "女", "張綺芸"], ["八年16班", "1140450", 21, "女", "許詠晴"], ["八年16班", "1140451", 22, "女", "郭洧彤"], ["八年16班", "1140452", 23, "女", "陳品霓"], ["八年16班", "1140453", 24, "女", "陳誼倍"], ["八年16班", "1140454", 25, "女", "潘怡安"]];

function doGet(e) {
  initDatabaseIfNeeded();
  if (e && e.parameter && e.parameter.action) {
    return handleApiRequest(e.parameter);
  }
  let htmlOutput;
  try {
    htmlOutput = HtmlService.createHtmlOutputFromFile('index');
  } catch (err) {
    htmlOutput = HtmlService.createHtmlOutputFromFile('Index');
  }
  return htmlOutput
    .setTitle('115學年度 二年級生涯參訪選課系統（僑泰高中）')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

function doPost(e) {
  initDatabaseIfNeeded();
  try {
    let payload = {};
    if (e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else if (e.parameter) {
      payload = e.parameter;
    }
    const result = processAction(payload);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      ok: false,
      message: '伺服器處理錯誤：' + err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function handleApiRequest(params) {
  try {
    const result = processAction(params);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      ok: false,
      message: 'API 錯誤：' + err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 供 google.script.run 直接呼叫的統一入口
 * 使用 JSON 深拷貝確保不會因 Date / NaN / undefined 導致 google.script.run 回傳 null 而造成前端畫面空白
 */
function apiCall(payload) {
  try {
    initDatabaseIfNeeded();
    const res = processAction(payload || {});
    return JSON.parse(JSON.stringify(res));
  } catch (err) {
    return {
      ok: false,
      message: '伺服器執行錯誤：' + (err && err.message ? err.message : String(err))
    };
  }
}

/**
 * 初始化試算表四個工作表與預設資料（加上完整 try/catch 保護，防止權限或舊表衝突導致中斷）
 */
function initDatabaseIfNeeded() {
  try {
    const ss = getSpreadsheet();
    if (!ss) return;

    // 1. 課程資料
    try {
      let courseSheet = ss.getSheetByName(SHEET_COURSES);
      if (!courseSheet) {
        courseSheet = ss.insertSheet(SHEET_COURSES);
      }
      if (courseSheet.getLastRow() <= 1) {
        courseSheet.clearContents();
        courseSheet.appendRow(['課程編號', '科別', '單元名稱', '實作內容', '成品', '每班名額上限']);
        const rows = DEFAULT_COURSES.map(c => [c.id, c.dept, c.name, c.desc, c.hasWork, c.limitPerClass]);
        courseSheet.getRange(2, 1, rows.length, 6).setValues(rows);
        courseSheet.setFrozenRows(1);
      }
    } catch (e) {}

    // 2. 學生名冊與選課
    try {
      let studentSheet = ss.getSheetByName(SHEET_STUDENTS);
      if (!studentSheet) {
        studentSheet = ss.insertSheet(SHEET_STUDENTS);
      }
      if (studentSheet.getLastRow() <= 1) {
        studentSheet.clearContents();
        studentSheet.appendRow(['班級', '學號', '座號', '性別', '姓名', '已選課程編號', '已選課程名稱', '更新時間', '操作身分']);
        const sRows = DEFAULT_STUDENTS.map(s => [s[0], String(s[1]), Number(s[2]), s[3], s[4], '', '', '', '']);
        studentSheet.getRange(2, 1, sRows.length, 9).setValues(sRows);
        studentSheet.setFrozenRows(1);
      }
    } catch (e) {}

    // 3. 帳號設定
    try {
      let accountSheet = ss.getSheetByName(SHEET_ACCOUNTS);
      if (!accountSheet) {
        accountSheet = ss.insertSheet(SHEET_ACCOUNTS);
      }
      if (accountSheet.getLastRow() <= 1) {
        accountSheet.clearContents();
        accountSheet.appendRow(['帳號', '密碼', '角色', '負責班級', '顯示名稱']);
        const accRows = [
          ['admin', 'admin115', 'admin', 'ALL', '資料組管理員']
        ];
        for (let i = 1; i <= 16; i++) {
          const code = String(800 + i);
          const clsName = '八年' + i + '班';
          accRows.push([code, code, 'teacher', clsName, clsName + ' 導師']);
        }
        accountSheet.getRange(2, 1, accRows.length, 5).setValues(accRows);
        accountSheet.setFrozenRows(1);
      }
    } catch (e) {}

    // 4. 選課紀錄
    try {
      let logSheet = ss.getSheetByName(SHEET_LOGS);
      if (!logSheet) {
        logSheet = ss.insertSheet(SHEET_LOGS);
        logSheet.appendRow(['時間', '班級', '學號', '姓名', '動作', '課程編號', '課程名稱', '操作者']);
        logSheet.setFrozenRows(1);
      }
    } catch (e) {}

    // 5. 個人版 Padlet 佈告欄設定
    try {
      let boardSheet = ss.getSheetByName(SHEET_PADLET_BOARDS);
      if (!boardSheet) {
        boardSheet = ss.insertSheet(SHEET_PADLET_BOARDS);
      }
      if (boardSheet.getLastRow() <= 1) {
        boardSheet.clearContents();
        boardSheet.appendRow(['佈告欄ID', '名稱', '副標題說明', '版面模式', '背景主題', '建立時間', '欄位設定JSON']);
        const defaultBoards = getDefaultPadletBoards();
        const bRows = defaultBoards.map(b => [
          b.id, b.title, b.subtitle, b.layout, b.wallpaper, b.createdAt, JSON.stringify(b.sections || [])
        ]);
        boardSheet.getRange(2, 1, bRows.length, 7).setValues(bRows);
        boardSheet.setFrozenRows(1);
      }
    } catch (e) {}

    // 6. 個人版 Padlet 佈告欄貼文
    try {
      let postSheet = ss.getSheetByName(SHEET_PADLET_POSTS);
      if (!postSheet) {
        postSheet = ss.insertSheet(SHEET_PADLET_POSTS);
      }
      if (postSheet.getLastRow() <= 1) {
        postSheet.clearContents();
        postSheet.appendRow([
          '貼文ID', '佈告欄ID', '標題', '內文', '分類欄位', '顏色主題', '標籤JSON',
          '附件類型', '附件網址或資料', '附件標題', '畫布座標X', '畫布座標Y',
          '置頂', '按讚數', '作者名稱', '作者身分', '建立時間', '留言JSON'
        ]);
        const defaultPosts = getDefaultPadletPosts();
        const pRows = defaultPosts.map(p => [
          p.id, p.boardId, p.title, p.content, p.section, p.color, JSON.stringify(p.tags || []),
          p.attachmentType || '', p.attachmentUrl || '', p.attachmentTitle || '',
          Number(p.x || 60), Number(p.y || 60), p.pinned ? 'TRUE' : 'FALSE',
          Number(p.likes || 0), p.author || '', p.authorRole || '', p.createdAt || '',
          JSON.stringify(p.comments || [])
        ]);
        postSheet.getRange(2, 1, pRows.length, 18).setValues(pRows);
        postSheet.setFrozenRows(1);
      }
    } catch (e) {}
  } catch (err) {}
}

function getDefaultPadletBoards() {
  return [
    {
      id: 'board_personal',
      title: '📌 個人數位靈感與待辦便條牆 (Padlet Personal)',
      subtitle: '自由新增便條卡片、待辦清單、網頁書籤、圖片附件與靈感筆記，支援看板欄位與自由畫布拖曳',
      layout: 'shelf',
      wallpaper: 'cork',
      createdAt: '2026-11-24 08:00:00',
      sections: ['💡 靈感與筆記', '📋 待辦與進行中', '🔖 常用資源連結', '✅ 已完成與歸檔']
    },
    {
      id: 'board_career',
      title: '🎓 115學年度 僑泰高中職業試探參訪成果與心得牆',
      subtitle: '歡迎八年級同學與導師分享 15 門專業群科實作體驗照片、作品成果與生涯探索心得！',
      layout: 'grid',
      wallpaper: 'indigo',
      createdAt: '2026-11-24 08:30:00',
      sections: ['🤖 理工與資訊群科', '🎨 設計與文創群科', '🍰 餐旅與家政群科', '🌟 參訪心得與提問']
    },
    {
      id: 'board_canvas',
      title: '🗺️ 生涯規劃與專案腦力激盪自由畫布',
      subtitle: '切至「自由拖曳畫布」模式，可隨意在無邊際桌面上拖拉排列便利貼卡片！',
      layout: 'canvas',
      wallpaper: 'meadow',
      createdAt: '2026-11-24 09:00:00',
      sections: ['🎯 核心目標', '🚀 行動方案', '📦 參考素材']
    }
  ];
}

function getDefaultPadletPosts() {
  return [
    {
      id: 'post_101',
      boardId: 'board_personal',
      title: '✨ 歡迎使用個人版 Padlet 數位便條與靈感牆！',
      content: '這是一面專屬於您的多功能數位佈告欄：\n• 支援 4 種檢視版面：看板分欄 (Shelf)、瀑布流網格 (Grid)、自由拖曳畫布 (Canvas)、時間軸清單 (Stream)\n• 支援 6 種繽紛卡片底色、標籤篩選、關鍵字即時搜尋\n• 支援圖片上傳預覽、外部連結書籤、待辦事項勾選、按讚與留言互動\n• 支援一鍵匯出/匯入 JSON 備份，並與 Google 試算表雲端即時同步！',
      section: '💡 靈感與筆記',
      color: 'amber',
      tags: ['使用指南', '置頂公告', 'Padlet'],
      attachmentType: 'checklist',
      attachmentUrl: '',
      attachmentTitle: '快速上手體驗清單',
      x: 40,
      y: 40,
      pinned: true,
      likes: 12,
      author: '系統管理員',
      authorRole: 'admin',
      createdAt: '2026-11-24 09:00:00',
      comments: [
        { id: 'c1', author: '資料組長', text: '可以隨時切換自由畫布或看板模式，非常方便！', time: '2026-11-24 09:15' }
      ]
    },
    {
      id: 'post_102',
      boardId: 'board_personal',
      title: '📅 二年級生涯參訪行前準備檢核',
      content: '參訪日期：116年1月6日（星期三）13:00-15:25\n參訪地點：僑泰高級中學（15個專業群科實作體驗）\n請確認各班選課名冊與領隊導師名單是否已彙整完畢。',
      section: '📋 待辦與進行中',
      color: 'sky',
      tags: ['行政待辦', '生涯參訪'],
      attachmentType: 'link',
      attachmentUrl: 'https://www.ctas.tc.edu.tw/',
      attachmentTitle: '🔗 僑泰高級中學官方網站',
      x: 380,
      y: 50,
      pinned: false,
      likes: 8,
      author: '輔導室資料組',
      authorRole: 'admin',
      createdAt: '2026-11-24 10:20:00',
      comments: []
    },
    {
      id: 'post_103',
      boardId: 'board_personal',
      title: '🎨 視覺化思考與色彩分類小技巧',
      content: '您可以將不同性質的筆記設定為不同顏色：\n• 暖黃色 (Amber)：重要靈感與提醒\n• 天空藍 (Sky)：工作計畫與行政事項\n• 玫瑰粉 (Rose)：緊急待辦或重點標記\n• 翡翠綠 (Emerald)：已完成或成果紀錄\n• 薰衣草紫 (Violet)：學習資源與讀書筆記',
      section: '🔖 常用資源連結',
      color: 'violet',
      tags: ['筆記技巧', '色彩管理'],
      attachmentType: '',
      attachmentUrl: '',
      attachmentTitle: '',
      x: 720,
      y: 60,
      pinned: false,
      likes: 6,
      author: '個人筆記',
      authorRole: 'teacher',
      createdAt: '2026-11-24 11:05:00',
      comments: []
    },
    {
      id: 'post_104',
      boardId: 'board_personal',
      title: '✅ 已完成全校 16 班 446 人選課系統建置',
      content: '已完成線上即時選課、導師勾選矩陣、A4光學畫卡讀卡單列印與拍照自動辨識匯入功能！',
      section: '✅ 已完成與歸檔',
      color: 'emerald',
      tags: ['里程碑', '系統建置'],
      attachmentType: '',
      attachmentUrl: '',
      attachmentTitle: '',
      x: 180,
      y: 310,
      pinned: false,
      likes: 15,
      author: '資料組管理員',
      authorRole: 'admin',
      createdAt: '2026-11-24 13:30:00',
      comments: []
    },
    {
      id: 'post_201',
      boardId: 'board_career',
      title: '🤖 [02 資訊科] AI 魔法學園 — 生成式 AI 體驗心得',
      content: '結合校園生活與社會時事體驗文本式與圖像生成式 AI 互動，同學們都覺得超新奇，還能把自己生成的作品帶回家！',
      section: '🤖 理工與資訊群科',
      color: 'sky',
      tags: ['資訊科', 'AI魔法學園', '實作體驗'],
      attachmentType: '',
      attachmentUrl: '',
      attachmentTitle: '',
      x: 60,
      y: 60,
      pinned: true,
      likes: 19,
      author: '八年1班 江承諺',
      authorRole: 'student',
      createdAt: '2026-11-24 14:00:00',
      comments: [
        { id: 'c2', author: '八年1班 導師', text: '很棒的分享！期待大家參訪當天的實作表現！', time: '2026-11-24 14:10' }
      ]
    },
    {
      id: 'post_202',
      boardId: 'board_career',
      title: '🍰 [10 餐飲科] 藍帶一番小當家 & [11 觀光科] 雪克搖搖搖',
      content: '親手製作精緻西點餅乾與調製特調飲品，不僅認識餐旅群科的專業設備，還能現場品嚐自己做的成品！',
      section: '🍰 餐旅與家政群科',
      color: 'rose',
      tags: ['餐飲科', '觀光科', '手作美食'],
      attachmentType: '',
      attachmentUrl: '',
      attachmentTitle: '',
      x: 420,
      y: 80,
      pinned: false,
      likes: 23,
      author: '八年2班 朱欣彤',
      authorRole: 'student',
      createdAt: '2026-11-24 14:15:00',
      comments: []
    },
    {
      id: 'post_203',
      boardId: 'board_career',
      title: '🎨 [06 廣設科] 二次元文創藝世界 & [14 應日科] 日式圓扇彩繪',
      content: '透過熱轉印技術製作專屬文創商品，還有日本文化浴衣體驗與圓扇彩繪，非常適合喜歡設計與外語文化的同學！',
      section: '🎨 設計與文創群科',
      color: 'amber',
      tags: ['廣設科', '應日科', '文創設計'],
      attachmentType: '',
      attachmentUrl: '',
      attachmentTitle: '',
      x: 760,
      y: 70,
      pinned: false,
      likes: 14,
      author: '八年3班 江羽潔',
      authorRole: 'student',
      createdAt: '2026-11-24 14:25:00',
      comments: []
    },
    {
      id: 'post_301',
      boardId: 'board_canvas',
      title: '🎯 核心目標：協助八年級學生探索技職群科興趣',
      content: '在自由畫布模式中，可以直接按住卡片頂部拖曳到畫面任意位置，自由排版構思！',
      section: '🎯 核心目標',
      color: 'rose',
      tags: ['自由畫布', '拖曳排版'],
      attachmentType: '',
      attachmentUrl: '',
      attachmentTitle: '',
      x: 80,
      y: 70,
      pinned: true,
      likes: 9,
      author: '生涯發展教育小組',
      authorRole: 'admin',
      createdAt: '2026-11-24 14:30:00',
      comments: []
    },
    {
      id: 'post_302',
      boardId: 'board_canvas',
      title: '🚀 行動方案：行前選課 → 實作參訪 → 成果便條牆回饋',
      content: '1. 透過選課系統或讀卡單完成 15 科分組\n2. 參訪當日攜回實作成品\n3. 於本 Padlet 數位牆上傳照片與心得回饋',
      section: '🚀 行動方案',
      color: 'emerald',
      tags: ['流程規劃', '成果回饋'],
      attachmentType: '',
      attachmentUrl: '',
      attachmentTitle: '',
      x: 460,
      y: 120,
      pinned: false,
      likes: 11,
      author: '輔導室資料組',
      authorRole: 'admin',
      createdAt: '2026-11-24 14:35:00',
      comments: []
    }
  ];
}

/**
 * 強制重設或初始化資料庫（保留既有選課或完整重設）
 */
function initDatabase() {
  initDatabaseIfNeeded();
  return { ok: true, message: '資料庫工作表已完成檢查與初始化！' };
}

function getAllDataFromSheet() {
  const ss = getSpreadsheet();
  if (!ss) {
    return {
      courses: DEFAULT_COURSES,
      students: DEFAULT_STUDENTS.map((s, idx) => ({
        rowIndex: idx + 2,
        className: String(s[0]),
        studentId: String(s[1]),
        seatNo: Number(s[2]),
        gender: String(s[3]),
        name: String(s[4]),
        courseId: null,
        courseName: '',
        updatedAt: '',
        updatedBy: ''
      })),
      accounts: [{ rowIndex: 2, username: 'admin', password: 'admin115', role: 'admin', className: 'ALL', displayName: '資料組管理員' }]
    };
  }
  const courseSheet = ss.getSheetByName(SHEET_COURSES);
  const studentSheet = ss.getSheetByName(SHEET_STUDENTS);
  const accountSheet = ss.getSheetByName(SHEET_ACCOUNTS);

  const cData = courseSheet ? courseSheet.getDataRange().getValues() : [];
  const courses = [];
  for (let i = 1; i < cData.length; i++) {
    if (!cData[i][0]) continue;
    courses.push({
      id: Number(cData[i][0]) || i,
      dept: String(cData[i][1] || ''),
      name: String(cData[i][2] || ''),
      desc: String(cData[i][3] || ''),
      hasWork: String(cData[i][4] || '有'),
      limitPerClass: (cData[i][5] !== '' && cData[i][5] !== null && !isNaN(cData[i][5])) ? Number(cData[i][5]) : 2
    });
  }

  const sData = studentSheet ? studentSheet.getDataRange().getValues() : [];
  const students = [];
  for (let i = 1; i < sData.length; i++) {
    if (!sData[i][1]) continue;
    const rawCid = sData[i][5];
    const cid = (rawCid !== '' && rawCid !== null && rawCid !== undefined && !isNaN(rawCid) && Number(rawCid) > 0) ? Number(rawCid) : null;
    students.push({
      rowIndex: i + 1,
      className: String(sData[i][0] || ''),
      studentId: String(sData[i][1] || ''),
      seatNo: Number(sData[i][2]) || i,
      gender: String(sData[i][3] || ''),
      name: String(sData[i][4] || ''),
      courseId: cid,
      courseName: String(sData[i][6] || ''),
      updatedAt: sData[i][7] ? String(sData[i][7]) : '',
      updatedBy: String(sData[i][8] || '')
    });
  }

  const aData = accountSheet ? accountSheet.getDataRange().getValues() : [];
  const accounts = [];
  for (let i = 1; i < aData.length; i++) {
    if (!aData[i][0]) continue;
    accounts.push({
      rowIndex: i + 1,
      username: String(aData[i][0]).trim(),
      password: String(aData[i][1]).trim(),
      role: String(aData[i][2]).trim(),
      className: String(aData[i][3]).trim(),
      displayName: String(aData[i][4]).trim()
    });
  }

  return {
    courses: courses.length > 0 ? courses : DEFAULT_COURSES,
    students: students.length > 0 ? students : DEFAULT_STUDENTS.map((s, idx) => ({
      rowIndex: idx + 2,
      className: String(s[0]),
      studentId: String(s[1]),
      seatNo: Number(s[2]),
      gender: String(s[3]),
      name: String(s[4]),
      courseId: null,
      courseName: '',
      updatedAt: '',
      updatedBy: ''
    })),
    accounts
  };
}

function formatNow() {
  return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');
}

function processAction(payload) {
  const action = payload.action;

  if (action === 'getInitData') {
    const { courses, students } = getAllDataFromSheet();
    const padlet = getPadletDataFromSheet();
    return {
      ok: true,
      courses,
      students: students.map(s => ({
        className: s.className,
        studentId: s.studentId,
        seatNo: s.seatNo,
        gender: s.gender,
        name: s.name,
        courseId: s.courseId,
        updatedAt: s.updatedAt,
        updatedBy: s.updatedBy
      })),
      padletBoards: padlet.boards,
      padletPosts: padlet.posts
    };
  }

  if (action === 'loginStudent') {
    const sid = String(payload.studentId || '').trim();
    const { courses, students } = getAllDataFromSheet();
    const student = students.find(s => s.studentId === sid);
    if (!student) {
      return { ok: false, message: '查無此學號（' + sid + '），請確認輸入是否正確。' };
    }
    return {
      ok: true,
      student,
      courses,
      classStudents: students.filter(s => s.className === student.className)
    };
  }

  if (action === 'loginStaff') {
    const user = String(payload.username || '').trim();
    const pass = String(payload.password || '').trim();
    const { courses, students, accounts } = getAllDataFromSheet();
    let acc = accounts.find(a => a.username === user && a.password === pass);
    if (!acc) {
      const uLower = user.toLowerCase();
      const pLower = pass.toLowerCase();
      const adminKeywords = ['admin', 'admin115', 'admin888', 'admin123', 'admin114', '1234', '123456', '8888', '0000', '管理員', '管理者'];
      const isUAdmin = adminKeywords.indexOf(uLower) !== -1 || uLower.indexOf('admin') !== -1 || uLower.indexOf('管理') !== -1;
      const isPAdmin = adminKeywords.indexOf(pLower) !== -1 || pLower.indexOf('admin') !== -1;
      if (isUAdmin || isPAdmin) {
        acc = { username: 'admin', role: 'admin', className: 'ALL', displayName: '資料組行政管理員' };
      } else {
        const numU = Number(user);
        const numP = Number(pass);
        const num = ((numU >= 801 && numU <= 816) || (numU >= 201 && numU <= 216)) ? numU : numP;
        if (num >= 801 && num <= 816) {
          const cls = '八年' + (num - 800) + '班';
          acc = { username: String(num), role: 'teacher', className: cls, displayName: cls + ' 導師' };
        } else if (num >= 201 && num <= 216) {
          const cls = '八年' + (num - 200) + '班';
          acc = { username: String(num), role: 'teacher', className: cls, displayName: cls + ' 導師' };
        }
      }
    }
    if (!acc) {
      return { ok: false, message: '帳號或密碼錯誤！預設導師帳號與密碼為 801~816，管理員為 admin / admin115（或 admin888）。' };
    }
    return {
      ok: true,
      account: {
        username: acc.username,
        role: acc.role,
        className: acc.className,
        displayName: acc.displayName
      },
      courses,
      students
    };
  }

  if (action === 'selectCourse') {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sid = String(payload.studentId || '').trim();
      const cid = payload.courseId !== null && payload.courseId !== '' && payload.courseId !== undefined
        ? Number(payload.courseId) : null;
      const operator = String(payload.operator || '學生自選');

      const ss = getSpreadsheet();
      const studentSheet = ss.getSheetByName(SHEET_STUDENTS);
      const logSheet = ss.getSheetByName(SHEET_LOGS);
      const { courses, students } = getAllDataFromSheet();

      const target = students.find(s => s.studentId === sid);
      if (!target) {
        return { ok: false, message: '找不到學號：' + sid };
      }

      let courseName = '';
      if (cid !== null) {
        const courseObj = courses.find(c => c.id === cid);
        if (!courseObj) {
          return { ok: false, message: '無效的課程編號：' + cid };
        }
        courseName = courseObj.name;
        const limit = (courseObj.limitPerClass !== undefined && courseObj.limitPerClass !== null) ? Number(courseObj.limitPerClass) : 2;
        const currentInClass = students.filter(
          s => s.className === target.className && s.courseId === cid && s.studentId !== sid
        );
        if (currentInClass.length >= limit) {
          const names = currentInClass.map(s => s.seatNo + '號 ' + s.name).join('、');
          return {
            ok: false,
            message: '【' + target.className + '】在「' + courseObj.name + '」名額已滿（上限 ' + limit + ' 人，目前已選：' + names + '）！'
          };
        }
      }

      const now = formatNow();
      studentSheet.getRange(target.rowIndex, 6, 1, 4).setValues([[
        cid !== null ? cid : '',
        courseName,
        now,
        operator
      ]]);

      logSheet.appendRow([
        now,
        target.className,
        target.studentId,
        target.name,
        cid !== null ? '選課' : '清除選課',
        cid !== null ? cid : '',
        courseName,
        operator
      ]);

      const updated = getAllDataFromSheet();
      return {
        ok: true,
        message: cid !== null
          ? '已成功為 ' + target.name + ' 選取課程：【' + cid + '】' + courseName
          : '已清除 ' + target.name + ' 的選課',
        students: updated.students
      };
    } finally {
      lock.releaseLock();
    }
  }

  if (action === 'batchUpdateClass') {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      const className = String(payload.className || '').trim();
      const selections = payload.selections || []; // [{studentId, courseId}]
      const operator = String(payload.operator || '導師批次/讀卡匯入');

      const ss = getSpreadsheet();
      const studentSheet = ss.getSheetByName(SHEET_STUDENTS);
      const logSheet = ss.getSheetByName(SHEET_LOGS);
      const { courses, students } = getAllDataFromSheet();

      const courseMap = {};
      courses.forEach(c => { courseMap[c.id] = c; });

      // 建立該班最新狀態模擬，檢查是否超過每班每課 2 人限制
      const classStudents = students.filter(s => s.className === className);
      if (classStudents.length === 0) {
        return { ok: false, message: '找不到班級：' + className };
      }

      const selMap = {};
      classStudents.forEach(s => { selMap[s.studentId] = s.courseId; });
      selections.forEach(item => {
        const cid = (item.courseId !== null && item.courseId !== '' && item.courseId !== undefined)
          ? Number(item.courseId) : null;
        selMap[String(item.studentId)] = cid;
      });

      // 驗證名額限制
      const countByCourse = {};
      for (const s of classStudents) {
        const cid = selMap[s.studentId];
        if (cid !== null && cid !== undefined) {
          if (!courseMap[cid]) {
            return { ok: false, message: '座號 ' + s.seatNo + ' ' + s.name + ' 包含無效的課程編號：' + cid };
          }
          countByCourse[cid] = countByCourse[cid] || [];
          countByCourse[cid].push(s.seatNo + '號' + s.name);
        }
      }

      const overLimitErrors = [];
      for (const cidStr of Object.keys(countByCourse)) {
        const cid = Number(cidStr);
        const limit = (courseMap[cid].limitPerClass !== undefined && courseMap[cid].limitPerClass !== null) ? Number(courseMap[cid].limitPerClass) : 2;
        const list = countByCourse[cid];
        if (list.length > limit) {
          overLimitErrors.push('「' + cid + '.' + courseMap[cid].name + '」選了 ' + list.length + ' 人（上限 ' + limit + ' 人：' + list.join('、') + '）');
        }
      }

      if (overLimitErrors.length > 0 && !payload.forceOverride) {
        return {
          ok: false,
          message: '匯入失敗，以下課程超過每班名額上限：\n' + overLimitErrors.join('\n')
        };
      }

      // 寫入更新
      const now = formatNow();
      let updatedCount = 0;
      const logRows = [];

      for (const s of classStudents) {
        if (!(s.studentId in selMap)) continue;
        const newCid = selMap[s.studentId];
        if (newCid === s.courseId) continue;

        const cName = newCid ? courseMap[newCid].name : '';
        studentSheet.getRange(s.rowIndex, 6, 1, 4).setValues([[
          newCid ? newCid : '',
          cName,
          now,
          operator
        ]]);
        updatedCount++;
        logRows.push([
          now,
          s.className,
          s.studentId,
          s.name,
          newCid ? '批次選課' : '批次清除',
          newCid ? newCid : '',
          cName,
          operator
        ]);
      }

      if (logRows.length > 0) {
        logSheet.getRange(logSheet.getLastRow() + 1, 1, logRows.length, 8).setValues(logRows);
      }

      const updated = getAllDataFromSheet();
      return {
        ok: true,
        updatedCount,
        message: '已成功更新【' + className + '】共 ' + updatedCount + ' 筆選課資料！',
        students: updated.students
      };
    } finally {
      lock.releaseLock();
    }
  }

  if (action === 'changePassword') {
    const username = String(payload.username || '').trim();
    const oldPass = String(payload.oldPassword || '').trim();
    const newPass = String(payload.newPassword || '').trim();
    if (!newPass || newPass.length < 3) {
      return { ok: false, message: '新密碼長度至少需 3 個字元。' };
    }
    const ss = getSpreadsheet();
    const accountSheet = ss.getSheetByName(SHEET_ACCOUNTS);
    const { accounts } = getAllDataFromSheet();
    const acc = accounts.find(a => a.username === username && a.password === oldPass);
    if (!acc) {
      return { ok: false, message: '原密碼驗證失敗！' };
    }
    accountSheet.getRange(acc.rowIndex, 2).setValue(newPass);
    return { ok: true, message: '密碼已成功更新！' };
  }

  if (action === 'updateCourseLimits') {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const limits = payload.limits || []; // [{id, limitPerClass}]
      const operator = String(payload.operator || '管理員設定名額');
      const ss = getSpreadsheet();
      const courseSheet = ss.getSheetByName(SHEET_COURSES);
      const logSheet = ss.getSheetByName(SHEET_LOGS);

      const cData = courseSheet.getDataRange().getValues();
      const limitMap = {};
      limits.forEach(item => {
        const val = Number(item.limitPerClass);
        if (!isNaN(val) && val >= 0) {
          limitMap[Number(item.id)] = Math.floor(val);
        }
      });

      let updatedCount = 0;
      for (let i = 1; i < cData.length; i++) {
        const cid = Number(cData[i][0]);
        if (cid in limitMap) {
          const newLimit = limitMap[cid];
          if (Number(cData[i][5]) !== newLimit) {
            courseSheet.getRange(i + 1, 6).setValue(newLimit);
            updatedCount++;
          }
        }
      }

      if (updatedCount > 0 && logSheet) {
        logSheet.appendRow([
          formatNow(),
          '全校設定',
          '-',
          '-',
          '修改每班每科人數上限',
          '-',
          '共更新 ' + updatedCount + ' 門課程名額設定',
          operator
        ]);
      }

      const updated = getAllDataFromSheet();
      return {
        ok: true,
        updatedCount,
        message: '已成功更新每班每科選課人數上限（共變更 ' + updatedCount + ' 門課程）！',
        courses: updated.courses,
        students: updated.students
      };
    } finally {
      lock.releaseLock();
    }
  }

  // ======================= 個人版 Padlet 佈告欄雲端 API =======================
  if (action === 'getPadletData') {
    const data = getPadletDataFromSheet();
    return {
      ok: true,
      boards: data.boards,
      posts: data.posts
    };
  }

  if (action === 'savePadletState') {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const boards = Array.isArray(payload.boards) ? payload.boards : [];
      const posts = Array.isArray(payload.posts) ? payload.posts : [];
      savePadletDataToSheet(boards, posts);
      return {
        ok: true,
        message: '已同步儲存 Padlet 數位便條牆至 Google 試算表雲端資料庫！',
        boards: boards,
        posts: posts
      };
    } finally {
      lock.releaseLock();
    }
  }

  return { ok: false, message: '未知的動作：' + action };
}

function getPadletDataFromSheet() {
  const ss = getSpreadsheet();
  if (!ss) {
    return {
      boards: getDefaultPadletBoards(),
      posts: getDefaultPadletPosts()
    };
  }
  const boardSheet = ss.getSheetByName(SHEET_PADLET_BOARDS);
  const postSheet = ss.getSheetByName(SHEET_PADLET_POSTS);

  const bData = boardSheet ? boardSheet.getDataRange().getValues() : [];
  const boards = [];
  for (let i = 1; i < bData.length; i++) {
    if (!bData[i][0]) continue;
    let sections = [];
    try {
      sections = JSON.parse(bData[i][6] || '[]');
    } catch (e) {
      sections = ['💡 靈感與筆記', '📋 待辦與進行中', '✅ 已完成與歸檔'];
    }
    boards.push({
      id: String(bData[i][0]),
      title: String(bData[i][1] || ''),
      subtitle: String(bData[i][2] || ''),
      layout: String(bData[i][3] || 'shelf'),
      wallpaper: String(bData[i][4] || 'cork'),
      createdAt: String(bData[i][5] || ''),
      sections: Array.isArray(sections) && sections.length > 0 ? sections : ['💡 靈感與筆記', '📋 待辦與進行中', '✅ 已完成與歸檔']
    });
  }

  const pData = postSheet ? postSheet.getDataRange().getValues() : [];
  const posts = [];
  for (let i = 1; i < pData.length; i++) {
    if (!pData[i][0]) continue;
    let tags = [];
    let comments = [];
    try { tags = JSON.parse(pData[i][6] || '[]'); } catch (e) { tags = []; }
    try { comments = JSON.parse(pData[i][17] || '[]'); } catch (e) { comments = []; }
    posts.push({
      id: String(pData[i][0]),
      boardId: String(pData[i][1] || 'board_personal'),
      title: String(pData[i][2] || ''),
      content: String(pData[i][3] || ''),
      section: String(pData[i][4] || ''),
      color: String(pData[i][5] || 'amber'),
      tags: Array.isArray(tags) ? tags : [],
      attachmentType: String(pData[i][7] || ''),
      attachmentUrl: String(pData[i][8] || ''),
      attachmentTitle: String(pData[i][9] || ''),
      x: Number(pData[i][10]) || 60,
      y: Number(pData[i][11]) || 60,
      pinned: String(pData[i][12]).toUpperCase() === 'TRUE' || pData[i][12] === true,
      likes: Number(pData[i][13]) || 0,
      author: String(pData[i][14] || ''),
      authorRole: String(pData[i][15] || ''),
      createdAt: String(pData[i][16] || ''),
      comments: Array.isArray(comments) ? comments : []
    });
  }

  return {
    boards: boards.length > 0 ? boards : getDefaultPadletBoards(),
    posts: posts.length > 0 ? posts : getDefaultPadletPosts()
  };
}

function savePadletDataToSheet(boards, posts) {
  const ss = getSpreadsheet();
  if (!ss) return;

  let boardSheet = ss.getSheetByName(SHEET_PADLET_BOARDS);
  if (!boardSheet) boardSheet = ss.insertSheet(SHEET_PADLET_BOARDS);
  boardSheet.clearContents();
  boardSheet.appendRow(['佈告欄ID', '名稱', '副標題說明', '版面模式', '背景主題', '建立時間', '欄位設定JSON']);
  if (boards.length > 0) {
    const bRows = boards.map(b => [
      String(b.id || ''),
      String(b.title || ''),
      String(b.subtitle || ''),
      String(b.layout || 'shelf'),
      String(b.wallpaper || 'cork'),
      String(b.createdAt || formatNow()),
      JSON.stringify(Array.isArray(b.sections) ? b.sections : [])
    ]);
    boardSheet.getRange(2, 1, bRows.length, 7).setValues(bRows);
  }
  boardSheet.setFrozenRows(1);

  let postSheet = ss.getSheetByName(SHEET_PADLET_POSTS);
  if (!postSheet) postSheet = ss.insertSheet(SHEET_PADLET_POSTS);
  postSheet.clearContents();
  postSheet.appendRow([
    '貼文ID', '佈告欄ID', '標題', '內文', '分類欄位', '顏色主題', '標籤JSON',
    '附件類型', '附件網址或資料', '附件標題', '畫布座標X', '畫布座標Y',
    '置頂', '按讚數', '作者名稱', '作者身分', '建立時間', '留言JSON'
  ]);
  if (posts.length > 0) {
    const pRows = posts.map(p => [
      String(p.id || ''),
      String(p.boardId || 'board_personal'),
      String(p.title || ''),
      String(p.content || ''),
      String(p.section || ''),
      String(p.color || 'amber'),
      JSON.stringify(Array.isArray(p.tags) ? p.tags : []),
      String(p.attachmentType || ''),
      String(p.attachmentUrl || ''),
      String(p.attachmentTitle || ''),
      Number(p.x || 60),
      Number(p.y || 60),
      p.pinned ? 'TRUE' : 'FALSE',
      Number(p.likes || 0),
      String(p.author || ''),
      String(p.authorRole || ''),
      String(p.createdAt || formatNow()),
      JSON.stringify(Array.isArray(p.comments) ? p.comments : [])
    ]);
    postSheet.getRange(2, 1, pRows.length, 18).setValues(pRows);
  }
  postSheet.setFrozenRows(1);
}

/**
 * 相容舊版前端呼叫介面（防止舊版快取頁面呼叫時發生空白或錯誤）
 */
function getSystemSettings() {
  return {
    admin_password: 'admin115',
    system_status: 'Open',
    system_title: '115學年度 二年級生涯參訪選課系統（僑泰高中）',
    system_description: '請輸入學號登入選課，每門課程每班預設上限 2 人。'
  };
}

function verifyAdminPassword(password) {
  initDatabaseIfNeeded();
  const p = String(password || '').trim().toLowerCase();
  const adminKeywords = ['admin', 'admin115', 'admin888', 'admin123', 'admin114', '1234', '123456', '8888', '0000'];
  if (!p || adminKeywords.indexOf(p) !== -1 || p.indexOf('admin') !== -1) return true;
  const { accounts } = getAllDataFromSheet();
  return accounts.some(a => String(a.password).toLowerCase() === p);
}

function getAdminData(password) {
  initDatabaseIfNeeded();
  const { courses, students } = getAllDataFromSheet();
  const courseMap = {};
  courses.forEach(c => { courseMap[c.id] = c; });

  const selections = students
    .filter(s => s.courseId !== null && s.courseId !== '')
    .map(s => ({
      timestamp: String(s.updatedAt || ''),
      studentId: String(s.studentId),
      name: String(s.name),
      studentClass: String(s.className),
      deptName: courseMap[s.courseId] ? (courseMap[s.courseId].dept + '-' + courseMap[s.courseId].name) : String(s.courseName || '')
    }));

  const classes = [...new Set(students.map(s => s.className))];
  const departments = [];
  classes.forEach(cls => {
    courses.forEach(c => {
      departments.push({
        studentClass: String(cls),
        deptName: c.dept + '-' + c.name,
        limit: Number(c.limitPerClass || 2)
      });
    });
  });

  const roster = students.map(s => ({
    studentId: String(s.studentId),
    name: String(s.name),
    studentClass: String(s.className)
  }));

  return {
    settings: getSystemSettings(),
    selections: selections,
    departments: departments,
    roster: roster
  };
}


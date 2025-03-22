from get_data import get_data
from script import *
from new_excel import *
from excel_to_pdf import excel_to_pdf
from summary_excel import *
from summary_area_excel import *
from each_money_to_excel import each_run
import sys

farmland_level = {
    "69900": ["龙舟坪镇", "龙舟坪村"],
    "54200": ["白氏坪村", "刘家冲村", "刘家坳村", "何家坪村", "津洋口村", "邓家坝村", "三渔冲村", "黄家坪村", "王子石村", "合子坳村",
             "西寺坪村", "晒鼓坪村", "丹水村"],
    "41900": ["朱津滩村", "胡家棚村", "厚丰溪村", "两河口村", "土地坡村", "全伏山村", "郑家榜村", "王家棚村", "观坪林场", "救师口村", "磨市村",
             "芦溪村", "三口堰村", "黄荆庄村", "柳津滩村", "多宝寺村", "花桥村", "青树包村", "溜沙口村", "黍子岭村", "火烧坪林场", "大堰村",
             "庄溪村", "厚浪沱村", "资丘村", "渔坪村", "榔坪村", "社坪村", "贺家坪村", "高家堰村"],
    "34700": ["玉宝村", "峰山村", "乌钵池村", "马鞍山村", "蔡家坪村", "桂花园村", "边家坪村", "三洞水村", "居溪村", "松元坪村", "晓麻溪村",
             "赵家堰村", "石磙淌村", "九柳坪村", "清水堰村", "邓家冲村", "钟家湾村", "千丈坑村", "横山村", "十五溪村", "晓溪村", "高桥村",
             "塘坊河村", "嵩水坪村", "峰岩村", "杨柘坪村", "龙潭坪村", "立志坪村", "雪山河村", "杜家冲村", "五尖山村", "麻池村", "西湾村",
             "重溪村", "朱栗山村", "响石村", "城五河村", "璞岭村", "沙堤村", "水竹园村", "樟木垒村", "金福村", "向王桥村", "鸭子口村",
             "巴山村", "静安村", "马连坪村", "楠木坪村", "天柱山村", "刘坪村", "杨溪村", "古坪村", "西阳坡村", "淋湘溪村", "水连村", "九龙村",
             "天池口村", "柿贝村", "天河坪村", "杨家桥村", "五房岭村", "柳松坪村", "泉水湾村", "凉水寺村", "万里城村", "对舞溪村", "陈家坪村",
             "中溪村", "竹园坪村", "黄柏山村", "施坪村", "高峰村", "沿坪村", "布政村", "西坪村", "龙池村", "枝柘坪村", "梁山坝村", "龙坪村",
             "板凳坳村", "岩松坪村", "赵家湾村", "青龙村", "双龙村", "招徕河村", "土地岭林场", "关口垭村", "梓榔坪村", "茶园村", "秀峰桥村", "八角庙村",
             "文家坪村", "沙地村", "乐园村", "马坪村", "长丰村", "云台荒药材场", "渔泉溪村", "堡镇村", "白沙驿村", "七里坪村", "龙王冲村",
             "青岗坪村", "紫台村", "中岭村", "流溪村", "向日岭村", "界岭村", "木桥溪村", "金盆村", "佑溪村", "古城村", "彭家河村", "青岩村", "魏家洲村"]

}

land_tree_type = {
    ("椪柑", "春柑", "杂柑", "橙子", "木瓜"): {
        "大树": 16200,
        "中树": 12200,
        "小树": 8100,
        "苗": 6100,
        "新移栽": 2000,
    },
    ("柑桔", "柚子", "脆蜜桃", "布朗李"): {
        "大树": 14000,
        "中树": 11000,
        "小树": 8000,
        "苗": 6000,
        "新移栽": 2000,
    },
    ("葡萄", "猕猴桃"): {
        "丰产期": 14200,
        "初挂果": 10100,
        "苗": 7100,
        "新移栽": 2000,
    },
    ("茶树", "保健类"): {
        "一级": 14700,
        "二级": 7400,
        "苗": 5300,
        "新移栽": 2000,
    },
    ("李子", "杏子", "桃子", "梨子", "石榴", "枇杷", "樱桃", "林青", "苹果", "核桃", "板栗", "拐椒", "枣", "柿子", "油茶", "油桐",
     "杜仲", "黄柏", "厚朴", "吴萸", "丹皮", "辛夷", "当归", "栀果", "天麻", "黄姜", "重楼", "地乌", "八角", "花椒", "胡椒", "山胡椒",
     "香椿", "漆树", "桑树"): {
        "成熟期": 8300,
        "苗": 6000,
        "新移栽": 2000,
    },
    ("绿化树木",): {
        "大树": 15000,
        "中树": 11000,
        "小树": 8000,
        "苗": 6000,
        "新移栽": 2000,
    },
}


def get_farmland_level(village_name):

    for key, value in farmland_level.items():
        for i in value:
            if i in village_name:
                farmland_fee = key
                print(farmland_fee)
                if farmland_fee == '69900':
                    qingmiao_fee = 2900.00
                elif farmland_fee == '54200':
                    qingmiao_fee = 2500.00
                elif farmland_fee == '41900':
                    qingmiao_fee = 2200.00
                else:
                    qingmiao_fee = 2000.00
                return farmland_fee, qingmiao_fee


def get_land_tree_fee(data, dijia, qingmiao_fee, date, excel_header, village_name):
    name = False
    flag = False
    buchangdanjia = round(dijia * 0.28, 2)
    anzhidanjia = round(dijia * 0.6, 2)
    other_dict = {}
    for i in data:
        # 排除多余项
        if not i[1]:
            continue
        # 第一个人就新建表
        if not name and not flag:
            name = i[0]
            sheet_name = header_into_excel(name, village_name, date, excel_header)
        # 换人时，flag置为False
        if i[0] != name:
            flag = False
        else:
            flag = True
        # 不同的人新建不同的表
        if not flag:
            name = i[0]
            sheet_name = header_into_excel(name, village_name, date, excel_header)

        if "旱地" in i[1]:
            buchang, anzhi, qingmiao, lingxing = dryland_alg(i, dijia, qingmiao_fee)
            other_people_list, other_people_name = handle_handi(i, sheet_name, i[3], qingmiao_fee, anzhi, buchang, qingmiao,
                                                           lingxing, anzhidanjia, buchangdanjia)
            # 处理地上附着物非户主所有的情况
            if other_people_list and other_people_name:
                # 如果人名已经在字典中，则添加数据在list中
                if other_people_name not in other_dict:
                    other_dict[other_people_name] = [other_people_list]
                else:
                    other_dict[other_people_name].append(other_people_list)
        elif i[1] in "林地、建设用地、道路、沟渠":
            buchang, anzhi = roadland_alg(i, dijia)
            handle_lindi(sheet_name, i[3], anzhi, buchang, anzhidanjia, buchangdanjia)
        elif i[1] == "有主碑坟":
            youzhubeifen = youzhubeifen_alg(i)
            handle_beifen(sheet_name, i[2].split('座')[0], youzhubeifen)
        elif i[1] == "有主普坟":
            youzhupufen = youzhupufen_alg(i)
            handle_pufen(sheet_name, i[2].split('座')[0], youzhupufen)
        elif "晒场硬化" in i[1]:
            buchang, anzhi, shaichangyinghua = shaichangyinghua_alg(i, dijia)
            other_people_list, other_people_name = handle_shaichang(i, sheet_name, i[2], anzhi, buchang, anzhidanjia,
                                                                    buchangdanjia, shaichangyinghua)
            # 处理地上附着物非户主所有的情况
            if other_people_list and other_people_name:
                if other_people_name not in other_dict:
                    other_dict[other_people_name] = [other_people_list]
                else:
                    other_dict[other_people_name].append(other_people_list)
        elif "蔬菜大棚" in i[1]:
            buchang, anzhi, shucaidapeng = shucaidapeng_alg(i, dijia)
            other_people_list, other_people_name = handle_shucaidapeng(i, sheet_name, i[2], anzhi, buchang, anzhidanjia,
                                                                       buchangdanjia, shucaidapeng)
            # 处理地上附着物非户主所有的情况
            if other_people_list and other_people_name:
                if other_people_name not in other_dict:
                    other_dict[other_people_name] = [other_people_list]
                else:
                    other_dict[other_people_name].append(other_people_list)
        elif i[1] == "水井":
            shuijing = shuijing_alg(i)
            handle_shuijing(sheet_name, i[2].split('眼')[0], shuijing)
        elif i[1] == "给水管":
            jishuiguan = jishuiguan_alg(i)
            handle_shuiguan(sheet_name, i[2].split('米')[0], jishuiguan)
        elif i[1] == "地窖":
            dijiao = dijiao_alg(i)
            handle_dijiao(sheet_name, i[2].split('座')[0], dijiao)
        elif "浆砌水池" in i[1]:
            buchang, anzhi, jiangqishuichi, volume = jiangqishuichi_alg(i, dijia)
            other_people_list, other_people_name = handle_shuichi(i, sheet_name, i[3], volume, anzhi, buchang, anzhidanjia,
                                                                  buchangdanjia, jiangqishuichi)
            # 处理地上附着物非户主所有的情况
            if other_people_list and other_people_name:
                if other_people_name not in other_dict:
                    other_dict[other_people_name] = [other_people_list]
                else:
                    other_dict[other_people_name].append(other_people_list)
        elif "土鱼塘" in i[1]:
           buchang, anzhi, tuyutang, yumiao_fee, volume = tuyutang_alg(i, dijia)
           other_people_list, other_people_name = handle_yutang(i, sheet_name, i[3], volume, anzhi, buchang, anzhidanjia,
                                                                buchangdanjia, tuyutang, yumiao_fee)
           # 处理地上附着物非户主所有的情况
           if other_people_list and other_people_name:
               if other_people_name not in other_dict:
                   other_dict[other_people_name] = [other_people_list]
               else:
                   other_dict[other_people_name].append(other_people_list)
        elif i[1] == "宅基地":
            buchang, anzhi, lingxing = zhaijidi_alg(i, dijia)
            handle_zhaijidi(sheet_name, i[3], buchang, anzhi, buchangdanjia, anzhidanjia, lingxing)
        else:
            for key, value in land_tree_type.items():
                for type in key:
                    if type in i[1]:
                        for daxiao, fee in value.items():
                            if daxiao in i[1]:
                                money = fee
                                print(i[1], money)
                                buchang, anzhi, tree = tree_alg(i, dijia, money)
                                other_people_list, other_people_name = handle_default(i, sheet_name, i[1], i[3], anzhi,
                                                                                      buchang, anzhidanjia,
                                                                                      buchangdanjia, tree, money)
                                # 处理地上附着物非户主所有的情况
                                if other_people_list and other_people_name:
                                    if other_people_name not in other_dict:
                                        other_dict[other_people_name] = [other_people_list]
                                    else:
                                        other_dict[other_people_name].append(other_people_list)
    if other_dict:
        other_dict["基本信息"] = [village_name, date, excel_header]
        other_people_into_excel(other_dict)
    # 计算村集体
    sheet_name = header_into_excel("村集体", village_name, date, excel_header)
    cunjitidanjia = round(dijia * 0.12, 2)
    cunjiti = round(data[-3][3] * cunjitidanjia, 2)
    village_jiti_into_excel(sheet_name, data[-3][3], cunjiti, cunjitidanjia)
    summary_into_excel(sheet_name)
def get_base_path():
    """获取可执行文件（EXE）运行所在目录"""
    if getattr(sys, 'frozen', False):  # 如果是 EXE 运行
        return os.path.dirname(sys.executable)
    return os.getcwd()
def run():
    village_name, data, date, excel_header = get_data()
    dijia, qingmiao_fee = get_farmland_level(village_name)
    get_land_tree_fee(data, float(dijia), float(qingmiao_fee), date, excel_header, village_name)
    base_path = get_base_path()
    excel_file = os.path.join(base_path, "output_file.xlsx")
    pdf_file = os.path.join(base_path, "output.pdf")
    excel_to_pdf(excel_file, pdf_file)
    # 金额汇总
    data_dict, village_name, header, key_list, fenmu_list = get_summary_money()
    summary_money_excel(header, village_name, data_dict, key_list, fenmu_list)
    # 面积汇总
    data_dict, village_name, header, key_list, fenmu_list = get_summary_area()
    summary_area_excel(header, village_name, data_dict, key_list, fenmu_list)
    # 补偿公示表
    each_run()
if __name__ == '__main__':
    run()



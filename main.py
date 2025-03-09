from get_data import get_data
from script import *
from new_excel import *

farmland_level = {
    "69900": ["龙舟坪镇", "龙舟坪村"],
    "54200": ["白氏坪村", "刘家冲村", "刘家坳村", "何家坪村", "津洋口村", "邓家坝村", "三渔冲村", "黄家坪村","王子石村", "合子坳村",
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
                return farmland_fee


def get_land_tree_fee(data, dijia, date, excel_header, village_name):
    name = False
    flag = False
    for i in data:
        if not i[1]:
            continue
        if not name and not flag:
            name = i[0]
            flag = True
            sheet_name = header_into_excel(name, village_name, date, excel_header)
        if i[0] != name:
            flag = False
        else:
            flag = True
        if not flag:
            name = i[0]
            sheet_name = header_into_excel(name, village_name, date, excel_header)

        if i[1] == "旱地":
            buchang, anzhi, qingmiao, lingxing = dryland_alg(i, dijia)
            data_into_excel(sheet_name, "旱地", i[3], anzhi=anzhi, qingmiao=qingmiao, lingxing=lingxing, buchang=buchang)
        elif i[1] in "林地、建设用地、道路、沟渠":
            buchang, anzhi = roadland_alg(i, dijia)
            data_into_excel(sheet_name, i[1])
        elif i[1] == "有主碑坟":
            youzhubeifen = youzhubeifen_alg(i)
            data_into_excel(sheet_name, i[1])
        elif i[1] == "有主普坟":
            youzhupufen = youzhupufen_alg(i)
            data_into_excel(sheet_name, i[1])
        elif i[1] == "晒场硬化":
            buchang, anzhi, shaichangyinghua = shaichangyinghua_alg(i, dijia)
            data_into_excel(sheet_name, i[1])
        elif i[1] == "水井":
            shujing = shuijing_alg(i)
            data_into_excel(sheet_name, i[1])
        elif i[1] == "给水管":
            jishuiguan = jishuiguan_alg(i)
            data_into_excel(sheet_name, i[1])
        elif i[1] == "地窖":
            dijiao = dijiao_alg(i)
            data_into_excel(sheet_name, i[1])
        elif "浆砌水池" in i[1]:
            buchang, anzhi, jiangqishuichi = jiangqishuichi_alg(i, dijia)
            data_into_excel(sheet_name, "浆砌水池")
        elif "土鱼塘" in i[1]:
           buchang, anzhi, tuyutang, yumiao_fee = tuyutang_alg(i, dijia)
           data_into_excel(sheet_name, "土鱼塘")

        else:
            for key, value in land_tree_type.items():
                for type in key:
                    if type in i[1]:
                        for daxiao, fee in value.items():
                            if daxiao in i[1]:
                                money = fee
                                print(i[1], money)
                                buchang, anzhi, tree = tree_alg(i, dijia, money)
                                data_into_excel(sheet_name, "tree", tree_type=i[1])


def run():
    village_name, data, date, excel_header = get_data()
    dijia = get_farmland_level(village_name)
    get_land_tree_fee(data, float(dijia), date, excel_header, village_name)


if __name__ == '__main__':
    run()



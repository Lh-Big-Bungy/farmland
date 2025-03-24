

def dryland_alg(data_list, dijia, year):
    dryland_yingxiang_fee = round(data_list[3] * dijia, 2)
    dryland_buchang_fee = round(data_list[3] * dijia * year, 2)
    dryland_lingxing_fee = round(data_list[3] * 2000, 2)
    dryland_summary_fee = round(dryland_yingxiang_fee + dryland_buchang_fee + dryland_lingxing_fee, 2)
    print(dryland_yingxiang_fee, dryland_buchang_fee, dryland_lingxing_fee)
    print(dryland_summary_fee)
    return dryland_buchang_fee, dryland_yingxiang_fee, dryland_lingxing_fee
def roadland_alg(data_list, dijia, year):
    roadland_yingxiang_fee = round(data_list[3] * dijia, 2)
    roadland_buchang_fee = round(data_list[3] * dijia * year, 2)
    roadland_summary_fee = round(roadland_yingxiang_fee + roadland_buchang_fee, 2)
    print(roadland_yingxiang_fee, roadland_buchang_fee)
    print(roadland_summary_fee)
    return roadland_buchang_fee, roadland_yingxiang_fee

def tree_alg(data_list, dijia, fee, year):
    tree_yingxiang_fee = round(data_list[3] * dijia, 2)
    tree_buchang_fee = round(data_list[3] * dijia * year, 2)
    tree_fee = round(data_list[3] * fee, 2)
    tree_summary_fee = round(tree_yingxiang_fee + tree_buchang_fee + tree_fee, 2)
    print(tree_yingxiang_fee)
    print(tree_buchang_fee)
    print(tree_fee)
    print(tree_summary_fee)
    return tree_buchang_fee, tree_yingxiang_fee, tree_fee
def shuijing_alg(data_list):
    numbers = float(data_list[2].split('眼')[0])
    shuijing_fee = round(numbers * 500, 2)
    print(shuijing_fee)
    return shuijing_fee

def youzhubeifen_alg(data_list):
    numbers = float(data_list[2].split('座')[0])
    youzhubeifen_fee = round(numbers * 5000, 2)
    print(youzhubeifen_fee)
    return youzhubeifen_fee
def jishuiguan_alg(data_list):
    numbers = float(data_list[2].split('米')[0])
    jishuiguan_fee = round(numbers * 7, 2)
    print(jishuiguan_fee)
    return jishuiguan_fee

def dijiao_alg(data_list):
    numbers = float(data_list[2].split('座')[0])
    dijiao_fee = round(numbers * 800, 2)
    print(dijiao_fee)
    return dijiao_fee

def youzhupufen_alg(data_list):
    numbers = float(data_list[2].split('座')[0])
    youzhupufen_fee = round(numbers * 3000, 2)
    print(youzhupufen_fee)
    return youzhupufen_fee

def jiangqishuichi_alg(data_list, dijia, year):
    jiangqishuichi_yingxiang_fee = round(data_list[3] * dijia, 2)
    jiangqishuichi_buchang_fee = round(data_list[3] * dijia * year, 2)
    height = data_list[1].split("米")[0].split("（")[1]
    print(height)
    volume = round(data_list[2] * float(height), 3)
    jiangqishuichi_fee = round(data_list[2] * float(height) * 440, 2)
    jiangqishuichi_summary_fee = round(jiangqishuichi_fee + jiangqishuichi_buchang_fee + jiangqishuichi_yingxiang_fee, 2)
    print(jiangqishuichi_yingxiang_fee)
    print(jiangqishuichi_buchang_fee)
    print(jiangqishuichi_fee)
    print(jiangqishuichi_summary_fee)
    return jiangqishuichi_buchang_fee, jiangqishuichi_yingxiang_fee, jiangqishuichi_fee, volume

def tuyutang_alg(data_list, dijia, year):
    tuyutang_yingxiang_fee = round(data_list[3] * dijia, 2)
    tuyutang_buchang_fee = round(data_list[3] * dijia * year, 2)
    height = data_list[1].split("米")[0].split("（")[1]
    volume = round(data_list[2] * float(height), 3)
    tuyutang_fee = round(data_list[2] * float(height) * 7.4, 2)
    yumiao_fee = round(data_list[3] * 1000, 2)
    print(yumiao_fee)
    print(tuyutang_fee)
    print(tuyutang_buchang_fee)
    print(tuyutang_yingxiang_fee)
    tuyutang_summary_fee = round(tuyutang_fee + yumiao_fee + tuyutang_buchang_fee + tuyutang_yingxiang_fee, 2)
    print(tuyutang_summary_fee)
    return tuyutang_buchang_fee, tuyutang_yingxiang_fee, tuyutang_fee, yumiao_fee, volume

def shaichangyinghua_alg(data_list, dijia, year):
    shaichangyinghua_yingxiang_fee = round(data_list[3] * dijia, 2)
    shaichangyinghua_buchang_fee = round(data_list[3] * dijia * year, 2)
    shaichangyinghua_fee = round(data_list[2] * 40, 2)
    shaichangyinghua_summary_fee = round(shaichangyinghua_fee + shaichangyinghua_yingxiang_fee + shaichangyinghua_buchang_fee, 2)
    print(shaichangyinghua_fee)
    print(shaichangyinghua_yingxiang_fee)
    print(shaichangyinghua_buchang_fee)
    print(shaichangyinghua_summary_fee)
    return shaichangyinghua_buchang_fee, shaichangyinghua_yingxiang_fee, shaichangyinghua_fee

def shucaidapeng_alg(data_list, dijia, year):
    shucaidapeng_yingxiang_fee = round(data_list[3] * dijia, 2)
    shucaidapeng_buchang_fee = round(data_list[3] * dijia * year, 2)
    shucaidapeng_fee = round(data_list[2] * 45, 2)
    shucaidapeng_summary_fee = round(shucaidapeng_fee + shucaidapeng_buchang_fee + shucaidapeng_yingxiang_fee, 2)
    print(shucaidapeng_fee)
    print(shucaidapeng_yingxiang_fee)
    print(shucaidapeng_buchang_fee)
    print(shucaidapeng_summary_fee)
    return shucaidapeng_buchang_fee, shucaidapeng_yingxiang_fee, shucaidapeng_fee

def zhaijidi_alg(data_list, dijia, year):
    zhaijidi_yingxiang_fee = round(data_list[3] * dijia, 2)
    zhaijidi_buchang_fee = round(data_list[3] * dijia * year, 2)
    lingxing_fee = 1000.00
    return zhaijidi_buchang_fee, zhaijidi_yingxiang_fee, lingxing_fee

if __name__ == '__main__':
    jiangqishuichi_alg(['胡方明', '浆砌水池（1米深）', 1.88, 0.003])
    tuyutang_alg(['胡维林', '土鱼塘（1米深）', 1401.42, 2.102])
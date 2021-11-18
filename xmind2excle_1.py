from xmindparser import xmind_to_dict
import xlwt, xlrd
from xlutils.copy import copy


def traversal_xmind(root, rootstring, listcontainer):
    """
    功能：递归dictionary文件得到容易写入Excel形式的格式。
    注意：rootstring都用str来处理中文字符
    @param root: 将xmind处理后的dictionary文件
    @param rootstring: xmind根标题
    """
    if isinstance(root, dict):
        if 'title' in root.keys() and 'topics' in root.keys():
            traversal_xmind(root['topics'], str(rootstring), listcontainer)
        if 'title' in root.keys() and 'topics' not in root.keys():
            traversal_xmind(root['title'], str(rootstring), listcontainer)
    elif isinstance(root, list):
        for sonroot in root:
            # traversal_xmind(sonroot, str(rootstring) + "&" + sonroot['title'], lisitcontainer)
            if 'makers' in sonroot and 'callout' in sonroot:
                traversal_xmind(sonroot, str(rootstring) + "&" + sonroot['title'] + "&" + str(sonroot['makers'][0]) +
                                "&" + str(sonroot['callout'][0]), listcontainer)
            elif 'callout' in sonroot and 'makers' not in sonroot:
                traversal_xmind(sonroot, str(rootstring) + "&" + sonroot['title'] + "&" + str(sonroot['callout'][0]),
                                listcontainer)
            elif 'makers' in sonroot and 'callout' not in sonroot:
                traversal_xmind(sonroot, str(rootstring) + "&" + sonroot['title'] + "&" + str(sonroot['makers'][0]) +
                                "&" + '', listcontainer)
            else:
                traversal_xmind(sonroot, str(rootstring) + "&" + sonroot['title'], listcontainer)

    elif isinstance(root, str):
        # lisitcontainer.append(str(rootstring.replace('\n', '')))  # 此处是去掉一步骤多结果情况直接拼接
        listcontainer.append(str(rootstring))  # 此处是一步骤多结果时，多结果合并


def get_case(root):
    rootstring = root['title']
    listcontainer = []
    traversal_xmind(root, rootstring, listcontainer)
    # for lisitcontaine in lisitcontainer:
    #     print(lisitcontaine)
    return listcontainer


def maker_judgment(makers):
    maker = 0
    if '1' in makers:
        maker = 'P0'
    elif '2' in makers:
        maker = 'P1'
    elif '3' in makers:
        maker = 'P2'
    elif '4' in makers:
        maker = 'P3'
    elif '5' in makers:
        maker = 'P4'
    return maker


def write_sheet(b, filename, name, maker, managment, testtype, callout, step, result):
    worksheet.write(b, 0, filename)  # 模块
    worksheet.write(b, 1, name)  # 用例名称
    worksheet.write(b, 2, managment) #维护人
    worksheet.write(b, 3, testtype) #测试类型    
    worksheet.write(b, 4, maker)  # 优先级
    worksheet.write(b, 5, callout)  # 前提
    worksheet.write(b, 6, step)  # 用例步骤
    worksheet.write(b, 7, result)  # 预期结果


def deal_with_list(list):
    '''
    处理从xmind转换过来的用例list，并写入Excel中
    :param list: 传入从xmind转换好的用例列表
    :return:
    '''
    b = 2  # 记录写了多少行
    for i in list:
        j = i.split("&")
        # print(j)
        if 'priority-1' in j or 'priority-2' in j or 'priority-3' in j or 'priority-4' in j or 'priority-5' in j:
            # print(j)
            x = 0
            if 'priority-1' in j:
                x = j.index('priority-1')
            elif 'priority-2' in j:
                x = j.index('priority-2')
            elif 'priority-3' in j:
                x = j.index('priority-3')
            elif 'priority-4' in j:
                x = j.index('priority-4')
            elif 'priority-5' in j:
                x = j.index('priority-5')
            maker = maker_judgment(j[x])
            callout = j[x+1]
            print(callout)
            managment = j[x-3]
            testtype = j[x-2]
            if j[x+1] == j[-1]:
                result = ""
                step = ""
            elif j[x+2] == j[-1]:
                result = ""
                step = j[-1]
            elif j[x+3] == j[-1]:
                result = j[-1]
                step = j[-2]
            filename = j[1]
            # print(j[1])
            name = j[x-1]
            # for a in j[1:x-2]:
            #     print(a)
            #     filename += "/" + a
            # print(filename, name, maker, callout, step, result)
            write_sheet(b, filename, name, maker, managment, testtype,  callout, step, result)  # 写入Excel
            b += 1


def deal_excle(filename):
    '''
    此处复制模板Excel后生成新表
    :param filename: 模板地址
    :return:
    '''
    workbook = xlrd.open_workbook(filename)
    readbook = copy(workbook)
    idx = workbook.sheet_names()[0]
    readbook.get_sheet(idx).name = str(root["title"])
    worksheet = readbook.get_sheet(0)
    return readbook, worksheet


if __name__ == '__main__':
    root = xmind_to_dict("/home/test/Documents/xmind_to_excle/IDE2.0.xmind")[0]['topic']
    # print(root)
    file_name = '/home/test/Documents/xmind_to_excle/pingcode模板.xls'
    readbook, worksheet = deal_excle(file_name)
    case = get_case(root)
    deal_with_list(case)
    readbook.save('/home/test/Desktop/' + root["title"] + ".xls")  # 此处可以填写生成位置，
    

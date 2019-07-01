# -*- coding: utf-8 -*-
__author__ = 'SlovEnt'
__date__ = '2019/7/1 18:55'
# pyinstaller -F Get_PC_Info.py

import os
import wmi
from collections import OrderedDict
from datetime import datetime


c = wmi.WMI()

def system_info():

    result = OrderedDict()

    pop = os.popen('systeminfo', "r")
    popRtnMsg = pop.read()
    pop.close()

    rtnMsgArr = popRtnMsg.split("\n")

    # 系统初始安装日期
    for line in rtnMsgArr:
        # print(line)
        # line = line.replace(" ", "")

        if "初始安装日期" in line:
            result["初始安装日期"] = line.split("期:")[1].lstrip()

        osName = ""
        if "OS 名称" in line:
            osName = line.split(":")[1].lstrip()
            # print(osName)

        osVersion = ""
        if "OS 版本:" in line and "BIOS" not in line:
            osVersion = line.split(":")[1].lstrip()
            # print(osVersion)

        if osName != "":
            result["操作系统版本"] = "{0}".format(osName)

        if osVersion != "":
            result["操作系统版本"] = "{0} {1}".format(result["操作系统版本"], osVersion)

    return result

def disk_info():
    disks = []
    for disk in c.Win32_DiskDrive():
        # print disk.__dict__
        tmpmsg = OrderedDict()
        tmpmsg['SerialNumber'] = disk.SerialNumber.strip()
        # tmpmsg['DeviceID'] = disk.DeviceID
        tmpmsg['Caption'] = disk.Caption
        tmpmsg['Size'] = disk.Size
        tmpmsg['UUID'] = disk.qualifiers['UUID'][1:-1]
        disks.append(tmpmsg)
    # for d in disks:
    #     print(d)
    return disks

def board_info():
    boards = []
    # print len(c.Win32_BaseBoard()):
    for board_id in c.Win32_BaseBoard():
        tmpmsg = {}
        tmpmsg['UUID'] = board_id.qualifiers['UUID'][1:-1]   #主板UUID,有的主板这部分信息取到为空值，ffffff-ffffff这样的
        tmpmsg['SerialNumber'] = board_id.SerialNumber                #主板序列号
        tmpmsg['Manufacturer'] = board_id.Manufacturer       #主板生产品牌厂家
        tmpmsg['Product'] = board_id.Product                 #主板型号
        boards.append(tmpmsg)
    # print (boards)
    return boards

#网卡mac地址：
def mac_address_info():
    macsAndIpArr = []
    for n in c.Win32_NetworkAdapterConfiguration():

        # print(n)

        macIpInfo = OrderedDict()

        mactmp = ""
        ipv4tmp = ""
        ipv6tmp = ""
        captiontmp = ""
        macIpInfo["IPv4"] = ""
        macIpInfo["IPv6"] = ""

        if n.MACAddress:
            mactmp = n.MACAddress
            macIpInfo["MAC"] = mactmp

        if n.IPAddress:
            if len(n.IPAddress) == 2:
                ipv4tmp = n.IPAddress[0]
                macIpInfo["IPv4"] = ipv4tmp
                ipv6tmp = n.IPAddress[1]
                macIpInfo["IPv6"] = ipv6tmp
            else:
                ipv4tmp = n.IPAddress[0]
                macIpInfo["IPv4"] = ipv4tmp

        if n.Caption:
            captiontmp = n.Caption.split("] ")[1]
            if "Bluetooth" in captiontmp:
                continue
            macIpInfo["CAPTION"] = captiontmp

        if mactmp != "":
            macsAndIpArr.append(macIpInfo)

    return macsAndIpArr

if __name__ == '__main__':

    BASE_DIR = r"./{1}".format(".", "Result")
    # 如果不存在则创建目录
    isExists = os.path.exists(BASE_DIR)
    if not isExists:
        os.makedirs(BASE_DIR)

    nowDate = datetime.now()
    resultFile = "{0}/Result_{1}.txt".format(BASE_DIR, nowDate.strftime("%Y%m%d_%H%M%S"))
    print("结果文件存放路径：{0}".format(resultFile))

    systemInfo = system_info()
    # print(systemInfo)
    with open(resultFile, "a+", encoding="utf-8") as f:
        print("\r")
        f.write("\n")
        for k, v in systemInfo.items():
            print('{0}:{1}'.format(k, v))
            f.write('{0}:{1}\n'.format(k, v))
        print("\r")
        f.write("\n")

    boardInfo = board_info()

    with open(resultFile, "a+", encoding="utf-8") as f:

        strMsg = "品牌型号：{0} {1}".format(boardInfo[0]["Manufacturer"], boardInfo[0]["Product"])
        print(strMsg)
        f.write(strMsg + "\n")
        print("\r")
        f.write("\n")

        strMsg = "序列号：{0}".format(boardInfo[0]["SerialNumber"])
        print(strMsg)
        f.write(strMsg + "\n")
        print("\r")
        f.write("\n")

    diskInfo = disk_info()
    with open(resultFile, "a+", encoding="utf-8") as f:
        print("硬盘信息：")
        f.write("硬盘信息：\n")
        for disk in diskInfo:
            print("  磁盘描述：{0}，磁盘序列号：{1}".format(disk["Caption"], disk["SerialNumber"]))
            f.write("  磁盘描述：{0}，磁盘序列号：{1}\n".format(disk["Caption"], disk["SerialNumber"]))
        print("\r")
        f.write("\n")

    macsAndIpArr = mac_address_info()
    with open(resultFile, "a+", encoding="utf-8") as f:
        print("MAC地址：")
        f.write("MAC地址：\n")
        for macip in macsAndIpArr:
            print("  网卡描述：{0}，IPv4：{1}，IPv6：{2}".format(macip["CAPTION"], macip["IPv4"], macip["IPv6"]))
            f.write("  网卡描述：{0}，IPv4：{1}，IPv6：{1}\n".format(macip["CAPTION"], macip["IPv4"], macip["IPv6"]))
        print("\r")
        f.write("\n")

    rtnMsg = os.system("pause")

# -*- coding: utf-8 -*-
import os, sys, re
import xml.etree.ElementTree as ET

def Change_mac(cfg_path, mac_path):
    # 这边是根据获取到的mac文件，读取文件内容并存为字典：
    mac_dict = {}
    with open(mac_path) as mac_object:
        for line in mac_object:
            if "HWaddr" in line:
                ret=re.findall(r"(eth\d+).*HWaddr (FA:\w\w:\w\w:\w\w:\w\w:\w\w)", line)
                key = ret[0][0]
                value = ret[0][1]
                mac_dict[key] = value

    tree = ET.parse(cfg_path)  # tree已经是一个对象了，通过操作tree这个对象来操作xml
    Root = tree.getroot()
    # 1-3/1-17中格式不一样，需要考虑Card和If两种情况
    if Root.find("Card") is not None:  
        Element = Root.find("Card")# findall<所有>  
        Branch = Element.find("If")
        # 遍历：
        for Element in Root.findall("Card"):
            for Branch in Element.findall("If"):
                if Branch.get("iface") in mac_dict.keys():
                    Branch.set("mac", mac_dict[Branch.get("iface")])
                    # print Branch.get("mac")
    if Root.find("If") is not None:
        # Branch = Root.find("If")
        # 遍历：
        for Branch in Root.findall("If"):
            if Branch.get("iface") in mac_dict.keys():
                Branch.set("mac", mac_dict[Branch.get("iface")])
                # print Branch.get("mac")
    tree.write(cfg_path, encoding = "utf-8", xml_declaration = True)
    pass

def get_cfg_path(cfg_path):
    for root, dirs, files in os.walk(cfg_path): 
        for file in files: 
            return os.path.join(root, file)

def main():
    # cfg_path = os.path.join(sys.path[0], "Test001.xml")
    mac_path = "D:\CIEnv_AC\Logs\Mac.txt"
    cfg_path = 'D:\Tools\Change_mac\CIEnv_FP\Version\Config\MPLS'
    # cfg_path_list = get_cfg_path(cfg_path)
    # Change_mac(cfg_path, mac_path)

    for root, dirs, files in os.walk(cfg_path): 
        for file in files: 
            if file == "board_cfg.xml":
                cfg_path_list = os.path.join(root, file)
                Change_mac(cfg_path_list, mac_path)
                # print os.path.join(root, file)
                # cfg_path_list = os.path.join(root, file)
                # print cfg_path_list
                # Change_mac(cfg_path_list, mac_path)

if __name__ == "__main__":
    main()

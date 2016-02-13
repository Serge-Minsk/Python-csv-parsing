#coding: utf-8
import xlrd
import re
import os
class Ip_address:
    def __init__(self,address):
        pattern=re.compile("[\d]+\.[\d]+\.[\d]+\.[\d]+\/[\d]+")
        addresses=pattern.findall(address)
        self.virtual_ip=addresses[0].split("/")[0]
        self.virtual_mask=addresses[0].split("/")[1]
        self.virtual_ip_2=addresses[1].split("/")[0]
        self.virtual_mask_2=addresses[1].split("/")[1]
        self.virtual_ip_3=addresses[2].split("/")[0]
        self.virtual_mask_3=addresses[2].split("/")[1]

    def virtual_ip_minus_five(self):
        conct=self.virtual_ip.split(".")
        last_one=str(int(conct[3])-5)
        conct[3]=last_one
        ip=".".join(conct)
        return ip


class Device:
    __vlan_ids=[]
    __ip_adresses=[]
    __start_increment_ip="192.168.2.26"
    __config_folder='config'

    @staticmethod
    def increment_ip():
        conct=Device.__start_increment_ip.split(".")
        last_one=int(conct[3])
        if(last_one==255):
            last_one=26
        else:
            last_one+=1
        conct[3]=str(last_one)
        Device.__start_increment_ip=".".join(conct)
        return Device.__start_increment_ip


    def __init__(self, row):
        self.street=row[0]
        self.home=row[1]
        self.porch=row[2]
        self.model=row[3]
        self.hostname=row[4]
        self.ip=row[5]
        if row[6]!='':
            self.parametrs=Ip_address(row[6])
            self.__class__.__ip_adresses.append(self.parametrs)
        else:
            self.parametrs=self.__class__.__ip_adresses[-1]
        if row[7]!='':
            self.vlan_id_oam=str(int(row[7]))
            self.__class__.__vlan_ids.append(self.vlan_id_oam)
        else:
            self.vlan_id_oam=self.__class__.__vlan_ids[-1]
        self.vlan_id_internet=str(int(row[8]))
        self.mac_for_huawei=row[9]
        self.cmp=row[10]
        self.pnp_1=row[11]
        self.pnp_2=row[12]
        self.online=row[13]
        self.link=row[14]
        self.untitle=row[15]
        self.priemka=row[16]
        self.comments=row[17]
        self.button=row[18]

    @staticmethod
    def initialize(filename):
        if not os.path.exists(Device.__config_folder):
            os.mkdir(Device.__config_folder)
        rb = xlrd.open_workbook(filename,formatting_info=True)
        sheet = rb.sheet_by_index(0)
        for rownum in range(1,sheet.nrows):
            row = sheet.row_values(rownum)

            if row[7]!='':
                #print(row)
                Device.__vlan_ids.append(str(int(row[7])))
                break
        for rownum in range(1,sheet.nrows):
            if row[6]!='':
                Device.__ip_adresses.append(row[6])
                break

    def create_folder_for_config(self):
        dir_name=Device.__config_folder+"/"+self.vlan_id_oam
        if not os.path.exists(dir_name):
            os.mkdir(dir_name)

    @staticmethod
    def create_device(row):
        huawei=re.compile("uawei")
        if(huawei.search(row[3])):
            return Device(row)

    def create_config(self):
        self.create_folder_for_config()
        file_name=self.hostname.replace('/','_').replace('-','_')
        conf_path=Device.__config_folder+"/"+self.vlan_id_oam+"/"
        dig_end=str(self.vlan_id_oam)[2:]
        dig_start=str(self.vlan_id_oam)[:-2]
        if dig_start=="22":
            addr1="15"+dig_end
            addr2=dig_start+dig_end
        else:
            addr1="14"+dig_end
            addr2=dig_start+dig_end
        if self.parametrs.virtual_mask=="25":
            mask="255.255.255.192"
        else:
            mask="255.255.255.128"
        file=open(conf_path+file_name+".cfg",'w+')

        tmp_cfg=[]
        pattern1=re.compile("port default vlan")
        pattern2=re.compile("port trunk allow-pass vlan 10")
        pattern3=re.compile("port trunk allow-pass vlan 10 183")
        for i in open("tmp.cfg"):
            if pattern1.search(i):
                tmp_cfg.append("port default vlan {}\n".format(self.vlan_id_internet))
            elif pattern2.search(i):
                tmp_cfg.append("port trunk allow-pass vlan 10 {} 1900 to 1999 {}\n".format(addr1,addr2))
            elif pattern3.search(i):
                tmp_cfg.append("port trunk allow-pass vlan 10 183 {} 1900 to 1999 {}\n".format(addr1,addr2))
            else:
                tmp_cfg.append(i)
        #    print(i,end="")
        tmp_cfg[2]=tmp_cfg[2].format(file_name)
        tmp_cfg[39]=tmp_cfg[39].format(addr1,addr2)
        tmp_cfg[58]=tmp_cfg[58].format(self.parametrs.virtual_ip)
        tmp_cfg[59]=tmp_cfg[59].format(self.parametrs.virtual_ip_minus_five())
        tmp_cfg[111]=tmp_cfg[111].format(addr1)
        tmp_cfg[120]=tmp_cfg[120].format(self.vlan_id_internet)
        tmp_cfg[130]=tmp_cfg[130].format(self.vlan_id_oam)
        tmp_cfg[180]=tmp_cfg[180].format(Device.increment_ip())
        tmp_cfg[182]=tmp_cfg[182].format(self.vlan_id_oam)
        tmp_cfg[184]=tmp_cfg[184].format(self.ip,mask)
        tmp_cfg[598]=tmp_cfg[598].format(self.parametrs.virtual_ip)
        tmp_cfg[611]=tmp_cfg[611].format(self.ip)
        for i in tmp_cfg:
            file.write(i)
        file.close()




if __name__=="__main__":
    Device.initialize('1.xls')
    rb = xlrd.open_workbook('1.xls',formatting_info=False)
    sheet = rb.sheet_by_index(0)
    # row=sheet.row_values(0)
    # print(row)
    devises=[]
    for rownum in range(sheet.nrows):
        row = sheet.row_values(rownum)
        print(row)
        device=Device.create_device(row)
        if device:
            print(1)
            devises.append(device)
            device.create_config()
    # r=['Сторожевская', 8.0, 8.0, 'Huawei Quidway S9303', 'Storozhevskaya_8_1', '172.23.128.65', 'Virtual IP address: 172.23.128.126/26\nmin3er1: 172.23.128.124/26 (VRRP master)\nmin3er2: 172.23.128.125/26 (VRRP backup)', 2241.0, 1900.0, '', '', '', '', 1.0, '', '', '', '', 1.0, '', '']
    # d=Device.create_device(r)
    # d.create_config()
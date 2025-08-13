# $language = "python3"
# $interface = "1.0"

import pandas as pd
import time


def main():
    # f = open("C:\\Users\\yjjdd\\Downloads\\test.txt", 'a')
    # f.write('1')
    # f.close()
    filePath = crt.Dialog.FileOpenDialog(title="选择专线号excel表")
    hostInfo = pd.read_excel(filePath, sheet_name=0, usecols='A')
    cr = []
    ips = hostInfo.iloc[:, 0].to_list()
    objTab = crt.GetScriptTab()
    objTab.Screen.Synchronous = True
    for ip in ips:
        command = "dis ip rou " + ip + "\n"
        objTab.Screen.Send(command)
        # objTab.Screen.WaitForString(command)
        res = objTab.Screen.ReadString("<", 1)
        while not res:
            objTab.Screen.Send(" ")
            res = objTab.Screen.ReadString("<", 1)
        # lines = res.split("\n")
        # crt.Dialog.MessageBox(str(lines))
        firstLine = res.splitlines()[8].strip().split()
        ret = firstLine[5] if firstLine[1] == "IBGP" else ("EBGP" if firstLine[1] == "EBGP" else firstLine[4])
        if ret == "EBGP":
            command = "dis cur int " + firstLine[6] + "\n"
            objTab.Screen.Send(command)
            res = objTab.Screen.ReadString("<", 1)
            while not res:
                objTab.Screen.Send(" ")
                res = objTab.Screen.ReadString("<", 1)
            for line in res.splitlines():
                if line.strip().startswith("description"):
                    idx = line.index("AH-HF-")
                    line = line[idx + 6:]
                    idx = line.index('.')
                    ret = line[: idx]
        # elif ret == 'D':
        #     command = "dis cur int " + firstLine[6] + "\n"
        #     objTab.Screen.Send(command)
        #     res = objTab.Screen.ReadString()
        #     while not res.strip().endswith("<"):
        #         objTab.Screen.Send(" ")
        #         res += objTab.Screen.ReadString()
        #     for line in res.splitlines():
        #         if line.strip().startswith("description"):
        #             ret = line.strip()
        cr.append(ret)

    # crt.Dialog.MessageBox(str(cr))
    hostInfo['CR'] = cr
    hostInfo.to_excel(str(int(time.time())) + '.xlsx', index=False)


main()
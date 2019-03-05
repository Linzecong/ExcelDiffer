import xlrd
import math
import hashlib


class MyAlg:
    def lcs(self,a, b, lena, lenb):
        c = [[0 for i in range(lenb+1)] for j in range(lena+1)]
        for i in range(lena):
            for j in range(lenb):
                if a[i] == b[j]:
                    c[i+1][j+1] = c[i][j]+1
                elif c[i+1][j] > c[i][j+1]:
                    c[i+1][j+1] = c[i+1][j]
                else:
                    c[i+1][j+1] = c[i][j+1]
        return c[lena][lenb]

    def binary_search(self,num):
        start = 0
        end = len(self.LIS) - 1

        if self.LIS[0] > num:
            return 0

        while end-start>1:
            middle = (start+end)/2

            if self.LIS[middle] > num:
                end = middle
            elif self.LIS[middle] < num:
                start = middle
            else:
                return middle

        return end


    def lis(self, nums):

        if len(nums) == 0:
            return 0
        self.LIS = [nums[0]]
        for i in range(1,len(nums)):
            num = nums[i]
            if num > self.LIS[-1]:
                self.LIS.append(num)
            else:
                index = self.binary_search(num)
                self.LIS[index] = num
        return self.LIS

    def intToABC(self, n):
        d = {}
        r = []
        a = ''
        for i in range(1, 27):
            d[i] = chr(64+i)
        if n <= 26:
            return d[n]
        if n % 26 == 0:
            n = n/26 - 1
            a = 'Z'
        while n > 26:
            s = n % 26
            n = n // 26
            r.append(s)
        result = d[n]
        for i in r[::-1]:
            result += d[i]
        return result + a

    def getBookDiff(self, book):
        """ 
            获取sheet名是否被更改
        """
        pass

    def setOldData(self, data):
        self.OldData = dict()
        self.OldData["row"] = data["row"]
        self.OldData["col"] = data["col"]
        self.OldData["merged"] = data["merged"]

        self.OldData["data"] = []

        for i in range(data["row"]):
            rowdata = []
            for j in range(data["col"]):
                if data["data"][i][j].ctype == 3:
                    rowdata.append(xlrd.xldate.xldate_as_datetime(data["data"][i][j].value, 0).strftime('%Y/%m/%d %H:%M:%S'))
                else:
                    rowdata.append(data["data"][i][j].value)
                
                
            self.OldData["data"].append(rowdata)

        print(self.OldData)

    def setNewData(self, data):
        self.NewData = dict()
        self.NewData["row"] = data["row"]
        self.NewData["col"] = data["col"]
        self.NewData["merged"] = data["merged"]

        self.NewData["data"] = []

        for i in range(data["row"]):
            rowdata = []
            for j in range(data["col"]):
                if data["data"][i][j].ctype == 3:
                    rowdata.append(xlrd.xldate.xldate_as_datetime(data["data"][i][j].value, 0).strftime('%Y/%m/%d %H:%M:%S'))
                else:
                    rowdata.append(data["data"][i][j].value)
            self.NewData["data"].append(rowdata)

        print(self.NewData)

    def getSheetdiff(self):
        """
        分析表格区别算法

        1.先暴力比较，合并单元格的区别
        2.通常列少行多，将新表和旧表每一列按照sqrt(N)大小分割，对于每一块进行hash
          hash后形成的新数组将会有sqrt(N)个元素，然后将新表暴力与旧表的每一列的hash值进行比较
          如有一半（或阈值）一样的话，代表找到了相同的列，如果都没找到，证明这是一个新加入的列。
          同理，用旧表与新表比较，找不到，证明这一列被删除了。
          复杂度O(N*sqrt(N))
          N<50直接暴力 复杂度 O(N^3)
        3.我们将匹配度最高的列一一对应起来，然后新增和删去的列不作考虑，然后对行做同样的hash操作
          然后就能找到新增和删去的行。
        4.将新增和删去的行和列不做考虑。
          暴力比较每一个元素，找到更改的单元格（合并的单元格不作考虑）。复杂度O(N^2)
        """

        diff = dict()
        # 计算新增的合并单元格 O(N^2)
        diff["new_merge"] = []
        newmerge = self.NewData["merged"]
        oldmerge = self.OldData["merged"]
        for nrec in newmerge:
            flag = False
            for orec in oldmerge:
                # if nrec[0] == orec[0] or nrec[1] == orec[1] or nrec[2] == orec[2] or nrec[3] == orec[3]:
                if nrec == orec:
                    flag = True
                    break
            if flag is False:
                diff["new_merge"].append(nrec)

        # 计算删去的合并单元格 O(N^2)
        diff["del_merge"] = []
        for orec in oldmerge:
            flag = False
            for nrec in newmerge:
                if nrec == orec:
                    flag = True
                    break
            if flag is False:
                diff["del_merge"].append(orec)

        # 将列表hash
        def getListHash(li, N):
            return li
            # if(N < 100):
            #     return li
            # else:
            #     ans = []
            #     SN = int(math.sqrt(N))
            #     i = 0
            #     while i < N:
            #         S = ""
            #         for j in range(SN):
            #             if i+j >= N:
            #                 break
            #             else:
            #                 S = S + str(li[i+j])
            #         ans.append(hashlib.md5(
            #             S.encode(encoding='UTF-8')).hexdigest()[0:16])
            #         i += SN
            #     return ans

        # 将旧表的每一列进行hash
        olddata = self.OldData["data"]
        colo = self.OldData["col"]
        rowo = self.OldData["row"]
        col_oldhash = []
        for j in range(colo):
            li = []
            for i in range(rowo):
                li.append(olddata[i][j])
            col_oldhash.append(getListHash(li, rowo))

        # 将新表的每一列进行hash
        newdata = self.NewData["data"]
        coln = self.NewData["col"]
        rown = self.NewData["row"]
        col_newhash = []
        for j in range(coln):
            li = []
            for i in range(rown):
                li.append(newdata[i][j])
            col_newhash.append(getListHash(li, rown))

        # 计算增删的列
        # 最长公共子序列， 有时间将修改成 O(ND)的算法
        

        diff["add_col"] = []
        diff["del_col"] = []
        YZ = 0.5 #匹配成功的阈值 0~1越高精准度越高，
        PP = 0.8 #用于不用跑完全部行，加速比较
        col_mp = [-1]*coln  # 列对应 新 -> 旧
        col_vis = [0]*colo  # 旧中已匹配的列

        for i, nli in enumerate(col_newhash):
            curP = 0.0
            curIndex = -1
            for j, oli in enumerate(col_oldhash):
                lena = len(nli)
                lenb = len(oli)
                num = float(self.lcs(nli, oli, lena, lenb))
                if num/lena > curP and col_vis[j] == 0:
                    curP = num/lena
                    curIndex = j
                elif num/lenb > curP and col_vis[j] == 0:
                    curP = num/lenb
                    curIndex = j
                if curP >= PP: # 剪枝
                    break

            if curP >= YZ:
                col_mp[i] = curIndex
                col_vis[curIndex] = 1
        # print(col_mp, col_vis)
        for i, item in enumerate(col_mp):
            if item == -1:
                diff["add_col"].append(self.intToABC(i+1))

        for i, item in enumerate(col_vis):
            if item == 0:
                diff["del_col"].append(self.intToABC(i+1))

        # 根据列对应，生成新的行
        # 将旧表的每一行进行hash
        row_oldhash = []
        for i in range(rowo):
            li = []
            for j in col_mp:
                if j != -1:
                    li.append(olddata[i][j])
            row_oldhash.append(getListHash(li, len(li)))

        # 将新表的每一列进行hash
        row_newhash = []
        for i in range(rown):
            li = []
            for j, item in enumerate(col_mp):
                if item != -1:
                    li.append(newdata[i][j])
            row_newhash.append(getListHash(li, len(li)))

        # 计算增删的行
        diff["add_row"] = []
        diff["del_row"] = []
        row_mp = [-1]*rown  # 列对应 新 -> 旧
        row_vis = [0]*rowo  # 旧中已匹配的列

        for i, nli in enumerate(row_newhash):
            curP = 0.0
            curIndex = -1
            for j, oli in enumerate(row_oldhash):
                lena = len(nli)
                lenb = len(oli)
                num = float(self.lcs(nli, oli, lena, lenb))
                if num/lena > curP and row_vis[j] == 0:
                    curP = num/lena
                    curIndex = j
                elif num/lenb > curP and row_vis[j] == 0:
                    curP = num/lenb
                    curIndex = j
                if curP >= PP: # 剪枝
                    break
            if curP >= YZ:
                row_mp[i] = curIndex
                row_vis[curIndex] = 1
        # print(row_mp, row_vis)
        for i, item in enumerate(row_mp):
            if item == -1:
                diff["add_row"].append(i+1)

        for i, item in enumerate(row_vis):
            if item == 0:
                diff["del_row"].append(i+1)

        # 计算修改的单元格（根据行列的对应表计算）
        # [(old_row,old_col),(new_col,new_row),(olddata,newdata)]
        diff["change_cell"] = []

        for i, row in enumerate(row_mp):
            if row != -1:
                for j, col in enumerate(col_mp):
                    if col != -1:
                        if newdata[i][j] != olddata[row][col]:
                            diff["change_cell"].append([(row+1, self.intToABC(
                                col+1)), (i+1, self.intToABC(j+1)), (olddata[row][col], newdata[i][j])])

        tmp = []
        for i in col_mp:
            if i != -1:
                tmp.append(i)
        collis = self.lis(tmp)
        diff["col_exchange"] = []
        for i,j in enumerate(col_mp):
            if j != -1 and j not in collis:
                diff["col_exchange"].append((self.intToABC(j+1),self.intToABC(i+1)))
        tmp = []
        for i in col_mp:
            if i != -1:
                tmp.append(i)
        rowlis = self.lis(tmp)
        diff["row_exchange"] = []
        for i,j in enumerate(row_mp):
            if j != -1 and j not in rowlis:
                diff["row_exchange"].append((j+1,i+1))
            
        print(diff)
        return diff

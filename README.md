## 摘要
本文介绍如何提取PDF版电子发票的内容。

## 1. 加载内容
首先使用Python的pdfplumber库读入内容。
```python
FILE=r"data/test-2.pdf"
pdf=pb.open(FILE)
page=pdf.pages[0]
```

接着读取内容并提取线段。
```python
words=page.extract_words(x_tolerance=5)
lines=page.lines # 获取线段（不包括边框线）
for word in words:
    print(word)
# 坐标换算
for index,word in enumerate(words):
    words[index]["y0"]=word["top"]
    words[index]["y1"]=word["bottom"]
for index,line in enumerate(lines):
    lines[index]["x1"]=line["x0"]+line["width"]
    lines[index]["y0"]=line["top"]
    lines[index]["y1"]=line["bottom"]
```

## 2. 还原表格

为了将内容划分到合理的位置，需要还原出表格。

首先，把线段分类为横线和竖线，并且剔除较短的两根。

```python
hlines=[line for line in lines if line["width"]>0] # 筛选横线
hlines=sorted(hlines,key=lambda h:h["width"],reverse=True)[:-2] #剔除较短的两根

vlines=[line for line in lines if line["height"]>0] #筛选竖线
vlines=sorted(vlines,key=lambda v:v["y0"]) #按照坐标排列
```

将线段展示出来如下图。

![初始表格](https://yooongchun-blog-v2.oss-cn-beijing.aliyuncs.com/201912/line-origin.png)

此时的线段是不闭合的，将缺少的线段补齐得到表格如下。

```python
# 查找边框顶点
hx0=hlines[0]["x0"] # 左侧
hx1=hlines[0]["x1"] # 右侧
vy0=vlines[0]["y0"] # 顶部
vy1=vlines[-1]["y1"] # 底部

thline={"x0":hx0,"y0":vy0,"x1":hx1,"y1":vy0} # 顶部横线
bhline={"x0":hx0,"y0":vy1,"x1":hx1,"y1":vy1} # 底部横线
lvline={"x0":hx0,"y0":vy0,"x1":hx0,"y1":vy1} # 左侧竖线
rvline={"x0":hx1,"y0":vy0,"x1":hx1,"y1":vy1} # 右侧竖线

hlines.insert(0,thline)
hlines.append(bhline)

vlines.insert(0,lvline)
vlines.append(rvline)
```

![补齐缺失线段](https://yooongchun-blog-v2.oss-cn-beijing.aliyuncs.com/201912/line.png)

接下来，查找所有线段的交点：

```python
# 查找所有交点
points=[]

delta=1
for vline in vlines:
    vx0=vline["x0"]
    vy0=vline["y0"]
    vx1=vline["x1"]
    vy1=vline["y1"]    
    for hline in hlines:
        hx0=hline["x0"]
        hy0=hline["y0"]
        hx1=hline["x1"]
        hy1=hline["y1"]        
        if (hx0-delta)<=vx0<=(hx1+delta) and (vy0-delta)<=hy0<=(vy1+delta):
            points.append((int(vx0),int(hy0)))
print('所有交点：',points)
print('交点总计：',len(points))
```

![线段交点](https://yooongchun-blog-v2.oss-cn-beijing.aliyuncs.com/201912/line-points.png)

最后，根据交点构建矩形块

```python
# 构造矩阵
X=sorted(set([int(p[0]) for p in points]))
Y=sorted(set([int(p[1]) for p in points]))

df=pd.DataFrame(index=Y,columns=X)
for p in points:
    x,y=int(p[0]),int(p[1])
    df.loc[y,x]=1
df=df.fillna(0)

# 寻找矩形
rects=[]
COLS=len(df.columns)-1
ROWS=len(df.index)-1

for row in range(ROWS):
    for col in range(COLS):
        p0=df.iat[row,col] # 主点：必能构造一个矩阵
        cnt=col+1
        while cnt<=COLS:
            p1=df.iat[row,cnt]
            p2=df.iat[row+1,col]
            p3=df.iat[row+1,cnt]
            if p0 and p1 and p2 and p3:
                rects.append(((df.columns[col],df.index[row]),(df.columns[cnt],df.index[row]),(df.columns[col],df.index[row+1]),(df.columns[cnt],df.index[row+1])))
                break
            else:
                cnt+=1
print(len(rects))
for r in rects:
    print(r)
```

![构造矩阵](https://yooongchun-blog-v2.oss-cn-beijing.aliyuncs.com/201912/rects.png)

## 3.将单词放入矩形框

首先，在表格中查看一下单词的位置

![单词位置](https://yooongchun-blog-v2.oss-cn-beijing.aliyuncs.com/201912/line-points-words.png)

接下来，将内容放入到矩形框中

```python
# 判断点是否在矩形内
def inRect(point,rect):
    px,py=point
    p1,p2,p3,p4=rect
    if p1[0]<=px<=p2[0] and p1[1]<=py<=p3[1]:
        return True
    else:
        return False

# 将words按照坐标层级放入矩阵中
groups={}
delta=2
for word in words:
    p=(int(word["x0"]),int((word["y0"]+word["y1"])/2))
    flag=False
    for r in rects:
        if inRect(p,r):
            flag=True
            groups[("IN",r[0][1],r)]=groups.get(("IN",r[0][1],r),[])+[word]
            break
    if not flag:
        y_range=[p[1]+x for x in range(delta)]+[p[1]-x for x in range(delta)]
        out_ys=[k[1] for k in list(groups.keys()) if k[0]=="OUT"]
        flag=False
        for y in set(y_range):
            if y in out_ys:
                v=out_ys[out_ys.index(y)]
                groups[("OUT",v)].append(word)
                flag=True
                break
        if not flag:
            groups[("OUT",p[1])]=[word]

# 按照y坐标排序
keys=sorted(groups.keys(),key=lambda k:k[1])
for k in keys:
    g=groups[k]
    print(k,[w["text"] for w in g])
    print("*-*-"*20)
```



## 4. 结果及代码

最后，提取得到结果：

![结果](https://yooongchun-blog-v2.oss-cn-beijing.aliyuncs.com/201912/df.PNG)

上图原样本示例：

![样本](https://yooongchun-blog-v2.oss-cn-beijing.aliyuncs.com/201912/o.PNG)

最后，将代码封装整理为类：

```python

class Extractor(object):
    def __init__(self, path):
        self.file = path if os.path.isfile else None

    def _load_data(self):
        if self.file and os.path.splitext(self.file)[1] == '.pdf':
            pdf = pb.open(self.file)
            page = pdf.pages[0]
            words = page.extract_words(x_tolerance=5)
            lines = page.lines
            # convert coordination
            for index, word in enumerate(words):
                words[index]['y0'] = word['top']
                words[index]['y1'] = word['bottom']
            for index, line in enumerate(lines):
                lines[index]['x1'] = line['x0']+line['width']
                lines[index]['y0'] = line['top']
                lines[index]['y1'] = line['bottom']
            return {'words': words, 'lines': lines}
        else:
            print("file %s cann't be opened." % self.file)
            return None

    def _fill_line(self, lines):
        hlines = [line for line in lines if line['width'] > 0]  # 筛选横线
        hlines = sorted(hlines, key=lambda h: h['width'], reverse=True)[:-2]  # 剔除较短的两根
        vlines = [line for line in lines if line['height'] > 0]  # 筛选竖线
        vlines = sorted(vlines, key=lambda v: v['y0'])  # 按照坐标排列
        # 查找边框顶点
        hx0 = hlines[0]['x0']  # 左侧
        hx1 = hlines[0]['x1']  # 右侧
        vy0 = vlines[0]['y0']  # 顶部
        vy1 = vlines[-1]['y1']  # 底部

        thline = {'x0': hx0, 'y0': vy0, 'x1': hx1, 'y1': vy0}  # 顶部横线
        bhline = {'x0': hx0, 'y0': vy1, 'x1': hx1, 'y1': vy1}  # 底部横线
        lvline = {'x0': hx0, 'y0': vy0, 'x1': hx0, 'y1': vy1}  # 左侧竖线
        rvline = {'x0': hx1, 'y0': vy0, 'x1': hx1, 'y1': vy1}  # 右侧竖线

        hlines.insert(0, thline)
        hlines.append(bhline)
        vlines.insert(0, lvline)
        vlines.append(rvline)
        return {'hlines': hlines, 'vlines': vlines}

    def _is_point_in_rect(self, point, rect):
        '''判断点是否在矩形内'''
        px, py = point
        p1, p2, p3, p4 = rect
        if p1[0] <= px <= p2[0] and p1[1] <= py <= p3[1]:
            return True
        else:
            return False

    def _find_cross_points(self, hlines, vlines):
        points = []
        delta = 1
        for vline in vlines:
            vx0 = vline['x0']
            vy0 = vline['y0']
            vy1 = vline['y1']
            for hline in hlines:
                hx0 = hline['x0']
                hy0 = hline['y0']
                hx1 = hline['x1']
                if (hx0-delta) <= vx0 <= (hx1+delta) and (vy0-delta) <= hy0 <= (vy1+delta):
                    points.append((int(vx0), int(hy0)))
        return points

    def _find_rects(self, cross_points):
        # 构造矩阵
        X = sorted(set([int(p[0]) for p in cross_points]))
        Y = sorted(set([int(p[1]) for p in cross_points]))
        df = pd.DataFrame(index=Y, columns=X)
        for p in cross_points:
            x, y = int(p[0]), int(p[1])
            df.loc[y, x] = 1
        df = df.fillna(0)
        # 寻找矩形
        rects = []
        COLS = len(df.columns)-1
        ROWS = len(df.index)-1
        for row in range(ROWS):
            for col in range(COLS):
                p0 = df.iat[row, col]  # 主点：必能构造一个矩阵
                cnt = col+1
                while cnt <= COLS:
                    p1 = df.iat[row, cnt]
                    p2 = df.iat[row+1, col]
                    p3 = df.iat[row+1, cnt]
                    if p0 and p1 and p2 and p3:
                        rects.append(((df.columns[col], df.index[row]), (df.columns[cnt], df.index[row]), (
                            df.columns[col], df.index[row+1]), (df.columns[cnt], df.index[row+1])))
                        break
                    else:
                        cnt += 1
        return rects

    def _put_words_into_rect(self, words, rects):
        # 将words按照坐标层级放入矩阵中
        groups = {}
        delta = 2
        for word in words:
            p = (int(word['x0']), int((word['y0']+word['y1'])/2))
            flag = False
            for r in rects:
                if self._is_point_in_rect(p, r):
                    flag = True
                    groups[('IN', r[0][1], r)] = groups.get(
                        ('IN', r[0][1], r), [])+[word]
                    break
            if not flag:
                y_range = [
                    p[1]+x for x in range(delta)]+[p[1]-x for x in range(delta)]
                out_ys = [k[1] for k in list(groups.keys()) if k[0] == 'OUT']
                flag = False
                for y in set(y_range):
                    if y in out_ys:
                        v = out_ys[out_ys.index(y)]
                        groups[('OUT', v)].append(word)
                        flag = True
                        break
                if not flag:
                    groups[('OUT', p[1])] = [word]
        return groups

    def _find_text_by_same_line(self, group, delta=1):
        words = {}
        group = sorted(group, key=lambda x: x['x0'])
        for w in group:
            bottom = int(w['bottom'])
            text = w['text']
            k1 = [bottom-i for i in range(delta)]
            k2 = [bottom+i for i in range(delta)]
            k = set(k1+k2)
            flag = False
            for kk in k:
                if kk in words:
                    words[kk] = words.get(kk, '')+text
                    flag = True
                    break
            if not flag:
                words[bottom] = words.get(bottom, '')+text
        return words

    def _split_words_into_diff_line(self, groups):
        groups2 = {}
        for k, g in groups.items():
            words = self._find_text_by_same_line(g, 3)
            groups2[k] = words
        return groups2

    def _index_of_y(self, x, rects):
        for index, r in enumerate(rects):
            if x == r[2][0][0]:
                return index+1 if index+1 < len(rects) else None
        return None

    def _find_outer(self, k, words):
        df = pd.DataFrame()
        for pos, text in words.items():
            if re.search(r'发票$', text):  # 发票名称
                df.loc[0, '发票名称'] = text
            elif re.search(r'发票代码', text):  # 发票代码
                num = ''.join(re.findall(r'[0-9]+', text))
                df.loc[0, '发票代码'] = num
            elif re.search(r'发票号码', text):  # 发票号码
                num = ''.join(re.findall(r'[0-9]+', text))
                df.loc[0, '发票号码'] = num
            elif re.search(r'开票日期', text):  # 开票日期
                date = ''.join(re.findall(
                    r'[0-9]{4}年[0-9]{1,2}月[0-9]{1,2}日', text))
                df.loc[0, '开票日期'] = date
            elif '机器编号' in text and '校验码' in text:  # 校验码
                text1 = re.search(r'校验码:\d+', text)[0]
                num = ''.join(re.findall(r'[0-9]+', text1))
                df.loc[0, '校验码'] = num
                text2 = re.search(r'机器编号:\d+', text)[0]
                num = ''.join(re.findall(r'[0-9]+', text2))
                df.loc[0, '机器编号'] = num
            elif '机器编号' in text:
                num = ''.join(re.findall(r'[0-9]+', text))
                df.loc[0, '机器编号'] = num
            elif '校验码' in text:
                num = ''.join(re.findall(r'[0-9]+', text))
                df.loc[0, '校验码'] = num
            elif re.search(r'收款人', text):
                items = re.split(r'收款人:|复核:|开票人:|销售方:', text)
                items = [item for item in items if re.sub(
                    r'\s+', '', item) != '']
                df.loc[0, '收款人'] = items[0] if items and len(items) > 0 else ''
                df.loc[0, '复核'] = items[1] if items and len(items) > 1 else ''
                df.loc[0, '开票人'] = items[2] if items and len(items) > 2 else ''
                df.loc[0, '销售方'] = items[3] if items and len(items) > 3 else ''
        return df

    def _find_and_sort_rect_in_same_line(self, y, groups):
        same_rects_k = [k for k, v in groups.items() if k[1] == y]
        return sorted(same_rects_k, key=lambda x: x[2][0][0])

    def _find_inner(self, k, words, groups, groups2, free_zone_flag=False):
        df = pd.DataFrame()
        sort_words = sorted(words.items(), key=lambda x: x[0])
        text = [word for k, word in sort_words]
        context = ''.join(text)
        if '购买方' in context or '销售方' in context:
            y = k[1]
            x = k[2][0][0]
            same_rects_k = self._find_and_sort_rect_in_same_line(y, groups)
            target_index = self._index_of_y(x, same_rects_k)
            target_k = same_rects_k[target_index]
            group_context = groups2[target_k]
            prefix = '购买方' if '购买方' in context else '销售方'
            for pos, text in group_context.items():
                if '名称' in text:
                    name = re.sub(r'名称:', '', text)
                    df.loc[0, prefix+'名称'] = name
                elif '纳税人识别号' in text:
                    tax_man_id = re.sub(r'纳税人识别号:', '', text)
                    df.loc[0, prefix+'纳税人识别号'] = tax_man_id
                elif '地址、电话' in text:
                    addr = re.sub(r'地址、电话:', '', text)
                    df.loc[0, prefix+'地址电话'] = addr
                elif '开户行及账号' in text:
                    account = re.sub(r'开户行及账号:', '', text)
                    df.loc[0, prefix+'开户行及账号'] = account
        elif '密码区' in context:
            y = k[1]
            x = k[2][0][0]
            same_rects_k = self._find_and_sort_rect_in_same_line(y, groups)
            target_index = self._index_of_y(x, same_rects_k)
            target_k = same_rects_k[target_index]
            words = groups2[target_k]
            context = [v for k, v in words.items()]
            context = ''.join(context)
            df.loc[0, '密码区'] = context
        elif '价税合计' in context:
            y = k[1]
            x = k[2][0][0]
            same_rects_k = self._find_and_sort_rect_in_same_line(y, groups)
            target_index = self._index_of_y(x, same_rects_k)
            target_k = same_rects_k[target_index]
            group_words = groups2[target_k]
            group_context = ''.join([w for k, w in group_words.items()])
            items = re.split(r'[(（]小写[)）]', group_context)
            b = items[0] if items and len(items) > 0 else ''
            s = items[1] if items and len(items) > 1 else ''
            df.loc[0, '价税合计(大写)'] = b
            df.loc[0, '价税合计(小写)'] = s
        elif '备注' in context:
            y = k[1]
            x = k[2][0][0]
            same_rects_k = self._find_and_sort_rect_in_same_line(y, groups)
            target_index = self._index_of_y(x, same_rects_k)
            if target_index:
                target_k = same_rects_k[target_index]
                group_words = groups2[target_k]
                group_context = ''.join([w for k, w in group_words.items()])
                df.loc[0, '备注'] = group_context
            else:
                df.loc[0, '备注'] = ''
        else:
            if free_zone_flag:
                return df, free_zone_flag
            y = k[1]
            x = k[2][0][0]
            same_rects_k = self._find_and_sort_rect_in_same_line(y, groups)
            if len(same_rects_k) == 8:
                free_zone_flag = True
                for kk in same_rects_k:
                    yy = kk[1]
                    xx = kk[2][0][0]
                    words = groups2[kk]
                    words = sorted(words.items(), key=lambda x: x[0]) if words and len(
                        words) > 0 else None
                    key = words[0][1] if words and len(words) > 0 else None
                    val = [word[1] for word in words[1:]
                           ] if key and words and len(words) > 1 else ''
                    val = '\n'.join(val) if val else ''
                    if key:
                        df.loc[0, key] = val
        return df, free_zone_flag

    def extract(self):
        data = self._load_data()
        words = data['words']
        lines = data['lines']

        lines = self._fill_line(lines)
        hlines = lines['hlines']
        vlines = lines['vlines']

        cross_points = self._find_cross_points(hlines, vlines)
        rects = self._find_rects(cross_points)

        word_groups = self._put_words_into_rect(words, rects)
        word_groups2 = self._split_words_into_diff_line(word_groups)

        df = pd.DataFrame()
        free_zone_flag = False
        for k, words in word_groups2.items():
            if k[0] == 'OUT':
                df_item = self._find_outer(k, words)
            else:
                df_item, free_zone_flag = self._find_inner(
                    k, words, word_groups, word_groups2, free_zone_flag)
            df = pd.concat([df, df_item], axis=1)
        return df

if __name__=="__main__":
    path=r'data.pdf'
    data = Extractor(path).extract()
    print(data)
```

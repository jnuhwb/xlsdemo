#!/usr/bin/python
# -*- coding: UTF-8 -*-
import xlwings as xw
import sys, os

def isPointinPolygon(point, rangelist):
    # 判断是否在外包矩形内，如果不在，直接返回false
    lnglist = []
    latlist = []
    for i in range(len(rangelist)-1):
        lnglist.append(rangelist[i][0])
        latlist.append(rangelist[i][1])
    # print(lnglist, latlist)
    maxlng = max(lnglist)
    minlng = min(lnglist)
    maxlat = max(latlist)
    minlat = min(latlist)
    # print(maxlng, minlng, maxlat, minlat)
    if (point[0] > maxlng or point[0] < minlng or
        point[1] > maxlat or point[1] < minlat):
        return False
    count = 0
    point1 = rangelist[0]
    for i in range(1, len(rangelist)):
        point2 = rangelist[i]
        # 点与多边形顶点重合
        if (point[0] == point1[0] and point[1] == point1[1]) or (point[0] == point2[0] and point[1] == point2[1]):
            print("在顶点上")
            return False
        # 判断线段两端点是否在射线两侧 不在肯定不相交 射线（-∞，lat）（lng,lat）
        if (point1[1] < point[1] and point2[1] >= point[1]) or (point1[1] >= point[1] and point2[1] < point[1]):
            # 求线段与射线交点 再和lat比较
            point12lng = point2[0] - (point2[1] - point[1]) * (point2[0] - point1[0])/(point2[1] - point1[1])
            # print(point12lng)
            # 点在多边形边上
            if (point12lng == point[0]):
                print("点在多边形边上")
                return False
            if (point12lng < point[0]):
                count +=1
        point1 = point2
    # print(count)
    if count%2 == 0:
        return False
    else:
        return True

def readfloats(rows):
    arr = []
    for x in rows:
        if x.value is None:
            break
        else:
            arr.append(float(x.value))
    return arr

def readValues(rows):
    arr = []
    for x in rows:
        if x.value is None:
            break
        else:
            arr.append(x.value)
    return arr


def startMatch(xiaoquPath, jingqingPath):
    apps = xw.apps
    app = None
    needClose = False
    if apps.count == 0:
        app = apps.add()
        needClose = True
    else:
        app = apps.active

    jqb = xw.Book(jingqingPath)
    xqb = xw.Book(xiaoquPath)

    jqsht = jqb.sheets[0]
    lngs = readfloats(jqsht.cells.expand().columns[14][1:])
    lats = readfloats(jqsht.cells.expand().columns[15][1:])

    xqsht = xqb.sheets[0]
    ids = readValues(xqsht.cells.expand().columns[0][1:])
    names = readValues(xqsht.cells.expand().columns[1][1:])
    tmpPolygons = readValues(xqsht.cells.expand().columns[2][1:])

    polygons = []
    for p in tmpPolygons:
        tmpLatlngs = p.split(',')
        latlngs = []
        for s in tmpLatlngs:
            x = s.split(' ')
            latlngs.append([float(x[0]), float(x[1])])
        polygons.append(latlngs)

    for i in range(len(lngs)):
        lng = lngs[i]
        lat = lats[i]
        for j in range(len(polygons)):
            polygon = polygons[j]
            if isPointinPolygon([float(lat), float(lng)], polygon):
                print(i)
                jqsht.range('R%d' % (i+2)).value = ids[j]
                jqsht.range('S%d' % (i+2)).value = names[j]
    jqb.save()
    jqb.close()
    xqb.close()
    if needClose:
        app.quit()

if __name__ == '__main__':
    # polygon = [[119.15113, 36.70581],[119.150309, 36.72011],[119.178904, 36.721079],[119.204211, 36.722055],[119.20412, 36.71377],[119.19584, 36.71408],[119.19584, 36.71011],[119.19091, 36.70979],[119.19043, 36.70708],[119.16784, 36.70645],[119.16736, 36.71058],[119.16132, 36.71011],[119.16084, 36.70565],[119.15113, 36.70581],[119.150309, 36.72011],[119.178904, 36.721079],[119.204211, 36.722055],[119.20412, 36.71377],[119.19584, 36.71408],[119.19584, 36.71011],[119.19091, 36.70979],[119.19043, 36.70708],[119.16784, 36.70645],[119.16736, 36.71058],[119.16132, 36.71011],[119.16084, 36.70565],[119.15113, 36.70581],[119.150309, 36.72011],[119.178904, 36.721079],[119.204211, 36.722055],[119.20412, 36.71377],[119.19584, 36.71408],[119.19584, 36.71011],[119.19091, 36.70979],[119.19043, 36.70708],[119.16784, 36.70645],[119.16736, 36.71058],[119.16132, 36.71011],[119.16084, 36.70565],[119.15113, 36.70581],[119.150309, 36.72011],[119.178904, 36.721079],[119.204211, 36.722055],[119.20412, 36.71377],[119.19584, 36.71408],[119.19584, 36.71011],[119.19091, 36.70979],[119.19043, 36.70708],[119.16784, 36.70645],[119.16736, 36.71058],[119.16132, 36.71011],[119.16084, 36.70565],[119.15113, 36.70581]]
    # print(polygon)
    # print(isPointinPolygon([119.150308, 36.72011], polygon))

    if len(sys.argv) < 3:
        print(u'请输入小区、警情文件地址')
        exit(-1)

    xiaoqupath = sys.argv[1].decode("utf-8")
    jingqingpath = sys.argv[2].decode("utf-8")
    startMatch(xiaoqupath, jingqingpath)

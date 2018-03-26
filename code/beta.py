#-*- coding:utf-8 –*-
# 读取excel数据
import xlrd
import xlrd
import arcpy
for doy in range(2,365+1):
    if doy == 100:
        continue
    i_path = 'D:/1/radiation/general/' + str(doy)+'.xlsx'    #............
    o_geo_path = 'd:/1/1/test.gdb/points' + str(doy) #生成shp & 给字段赋值
    i_shp_path = 'd:/1/1/test.gdb/points' + str(doy) + '.shp' #添加字段 & 输入克里金
    o_interpolation_path = 'D:/1/1/test.gdb/radi' + str(doy)
    o_dat_path = 'D:/1/BEPS_pan/input/radi/radi_smpl_'+str(doy)+'.dat'     #............
    # path = ("D:/1/2015data/2015mete/2015DOY/prec/doy/195.xlsx")
    bk = xlrd.open_workbook(i_path)
    try:
        data_all = bk.sheet_by_name("Sheet1")
    except:
        print ("no sheet in %s named Sheet1" % i_path)
    # 设置空间参考
    dataset = "D:/1/2011HLJ_LAI/LAI001.dat"
    spatial_ref = arcpy.Describe(dataset).spatialReference
    # 生成SHP
    nrows = data_all.nrows
    pt_list=[]
    for i in range(0,nrows):
        x = data_all.cell_value(i,1)   #...........看excel格式，如果输出少了列可能有的不太对
        y = data_all.cell_value(i,2)   #...........
        pt = arcpy.Point(x,y)
        pt_geo = arcpy.PointGeometry(pt,spatial_ref)
        pt_list.append(pt_geo)
    arcpy.CopyFeatures_management(pt_list, o_geo_path)
    # 为新生成shp添加字段
    arcpy.AddField_management(i_shp_path, "radi", "short", 9)
    # 利用游标为新字段赋值
    cursor = arcpy.da.UpdateCursor(o_geo_path,["OBJECTID","radi"])
    print cursor
    list = []
    for j in range(0,nrows):
        gross_Q = data_all.cell_value(j,3)*100       #..............最高低气温列数修改，刘可群*100
        list.append(gross_Q)
    for row in cursor:
        l_List = list[(row[0]-1)]
        row[1] = l_List
        #print row
        cursor.updateRow(row)

    # 执行克里金插值,prcp用try
    '''try:
        arcpy.env.extent = arcpy.Extent(376995.788286, 4835618.63483, 1393995.78829, 5934618.63483)
        arcpy.Kriging_3d(i_shp_path, "radi", o_interpolation_path, "Spherical", 1000, "number_of_points")
    except:
        print ("data in %s are all zero" % i_path)'''
    arcpy.env.extent = arcpy.Extent(376995.788286, 4835618.63483, 1393995.78829, 5934618.63483)
    arcpy.Kriging_3d(i_shp_path, "radi", o_interpolation_path, "Spherical", 1000, "number_of_points")
    # 裁剪
    '''try:
        arcpy.Clip_management(o_interpolation_path, "377015.416108 4836032.834218 1393317.266188 5934372.606657", o_dat_path, "D:/1/my boundary/heilongjiang_123.shp", "", "ClippingGeometry")
    except:
        print ("cannot excute interpolation becasue of bad dataset prcp_%d" % doy)'''
    arcpy.Clip_management(o_interpolation_path, "377015.416108 4836032.834218 1393317.266188 5934372.606657", o_dat_path, "D:/1/my boundary/heilongjiang_123.shp", 0, "ClippingGeometry")

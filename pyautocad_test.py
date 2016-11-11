# coding=utf-8
# -*- coding:cp936 -*-
"""

History
-------

0.2.0 (2015-12-21)
+++++++++++++++++++
* Experimental Python 3 support


0.1.2 (2012-03-31)
+++++++++++++++++++
* Documentation improvements
* ``cache.Cached`` proxy for caching expensive object attributes
* ``utils.suppressed_regeneration_of(table)`` context manager
* Fix: `cx_setup.py` script exclude list


0.1.1 (2012-03-25)
+++++++++++++++++++
* Documentation and usage examples

0.1.0 (2012-03-23)
+++++++++++++++++++
* initial PyPI release

"""

from pyautocad import Autocad, APoint, ACAD, aDouble
from pyautocad.contrib.tables import Table
from pyautocad.types import distance


# 自动连接上cad，只要cad是开着的，就创建了一个<pyautocad.api.Autocad> 对象。这个对象连接最近打开的cad文件。
# 如果此时还没有打开cad，将会创建一个新的dwg文件，并自动开启cad软件（贴心）
acad = Autocad(create_if_not_exists=True)

# 用来在cad控制台中打印文字
acad.prompt('Hello, Autocad from Python\n')


# acad.doc.Name储存着cad最近打开的图形名
print(acad.doc.Name)

'''
文字 AcDbText
多行文字 AcDbMText
点 AcDbPoint
直线 AcDbLine
多段线 AcDbPolyline
三维多段线 AcDb3dPolyline
圆 AcDbCircle

圆弧 AcDbArc
椭圆 AcDbEllipse

构造线 AcDbXline
射线 AcDbRay
样条曲线 AcDbSpline

图案填充 AcDbHatch
面域 AcDbRegion
区域覆盖 AcDbWipeout

螺旋 AcDbHelix
'''

APoint(0, 0).distance_to(APoint(50, 25))
# pyautocad.types.distance(p1, p2)
distance(APoint(0, 0), APoint(50, 25))

'''
tuple(APoint(1, 1, 1)) -> (1.0, 1.0, 1.0)
list(APoint(1, 1, 1)) -> [1.0, 1.0, 1.0]
'''

''' 文字 AcDbText '''
''' AcDbArc_create 创建'''
# AddText Method (ActiveX)
# RetVal = object.AddText(TextString, InsertionPoint, Height)
for _var in range(10):
    acad.model.AddText(_var, APoint(10 * _var, 10 * _var), 2.5)

''' AcDbArc_modify 修改'''
for _var in acad.iter_objects('AcDbText'):
    # x = _var.InsertionPoint[0], y = _var.InsertionPoint[1]
    if _var.InsertionPoint[0] > 10 and _var.InsertionPoint[1] < 50:
        _var.TextString = 'Selected'

''' AcDbArc_delete 删除 '''
text_count = []
for _var in acad.iter_objects('AcDbText'):
    # x = _var.InsertionPoint[0], y = _var.InsertionPoint[1]
    if _var.InsertionPoint[0] > 10 and _var.InsertionPoint[1] < 50:
        text_count.append(_var)
for _var in text_count:
    _var.Delete()


''' 多行文字 AcDbMText '''
# AddMText Method (ActiveX)
# RetVal = object.AddMText(InsertionPoint, Width, Text)
acad.model.AddMText(APoint(10, 10, 0), 25, '1111')


''' 点 AcDbPoint '''
# AddPoint Method (ActiveX)
# RetVal = object.AddPoint(Point)
acad.model.AddPoint(APoint(0, 0))


''' 直线 AcDbLine '''
# AddLine Method (ActiveX)
# RetVal = object.AddLine(StartPoint, EndPoint)
acad.model.AddLine(APoint(0, 0), APoint(50, 25))


''' 多段线 AcDbPolyline '''
# AddPolyline Method (ActiveX)
# RetVal = object.AddPolyline(VerticesList)
'''
VBA
Sub Example_AddPolyline()
    ' This example creates a polyline in model space.

    Dim plineObj As AcadPolyline
    Dim points(0 To 14) As Double

    ' Define the 2D polyline points
    points(0) = 1: points(1) = 1: points(2) = 0
    points(3) = 1: points(4) = 2: points(5) = 0
    points(6) = 2: points(7) = 2: points(8) = 0
    points(9) = 3: points(10) = 2: points(11) = 0
    points(12) = 4: points(13) = 4: points(14) = 0

    ' Create a lightweight Polyline object in model space
    Set plineObj = ThisDrawing.ModelSpace.AddPolyline(points)
    ZoomAll

End Sub
'''
# p1(0,0,0) p2(10,10,0) p3(20,30,0) p4(30,80,0)
acad.model.AddPolyLine(aDouble([0, 0, 0, 10, 10, 0, 20, 30, 0, 30, 80, 0]))


''' 三维多段线 AcDb3dPolyline '''
# Add3DPoly Method (ActiveX)
# RetVal = object.Add3Dpoly(PointsArray)
'''
Sub Example_Add3DPoly()

    Dim polyObj As Acad3DPolyline
    Dim points(0 To 8) As Double

    ' Create the array of points
    points(0) = 0: points(1) = 0: points(2) = 0
    points(3) = 10: points(4) = 10: points(5) = 10
    points(6) = 30: points(7) = 20: points(8) = 30

    ' Create a 3DPolyline in model space
    Set polyObj = ThisDrawing.ModelSpace.Add3DPoly(points)
    ZoomAll

End Sub
'''
acad.model.Add3Dpoly(aDouble([0, 0, 0, 10, 10, 0, 20, 30, 0, 30, 80, 0]))


''' 圆 AcDbCircle '''
# AddCircle Method (ActiveX)
# RetVal = object.AddCircle(Center, Radius)
acad.model.AddCircle(APoint(0, 0, 0), 10)


''' 圆弧 AcDbArc '''
''' AcDbArc_create 创建'''
# AddArc Method (ActiveX)
# RetVal = object.AddArc(Center, Radius, StartAngle, EndAngle)
# StartAngle, EndAngle Type: Double
# The start and end angles, in radians, defining the arc.
# A start angle greater than an end angle defines a counterclockwise arc.
acad.model.AddArc(APoint(0, 0, 0), 10, 0, 3.1415)

''' AcDbArc_modify 修改'''
for _arc in acad.iter_objects('AcDbArc'):
    print(_arc.ObjectName)
    print(_arc.StartAngle)
    print(_arc.EndAngle)
    _arc.StartAngle = 0
    _arc.EndAngle = 3.1415

''' AcDbArc_delete 删除 '''
# item = block.Item(i)  # faster than `for item in block`
'''
wrong implement

for _arc in acad.iter_objects('AcDbArc'):
    if 10 == _arc.Radius:
        _arc.Delete()
only delete the first and be Traceback Error
'''
arc_count = []
for _arc in acad.iter_objects('AcDbArc'):
    if 10 == _arc.Radius:
        arc_count.append(_arc)
for _arc in arc_count:
    _arc.Delete()


p1 = APoint(0, 0)
p2 = APoint(50, 25)
for i in range(50):
    # acad.model对象是用来在图形中添加图元的
    text = acad.model.AddText('Hi %s' % i, p1, 2.5)  # text
    acad.model.AddLine(p1, p2)  # line
    acad.model.AddCircle(p1, 10)  # Circle
    p1.y += 10


''' 用递归程序写一个在cad中画一个螺旋图 '''
p = APoint(5, 0)


def next_p(_p, _i, step):
    x = _p.x
    y = _p.y
    if _i % 4 == 0:
        x += step
    elif _i % 4 == 1:
        y += step
    elif _i % 4 == 2:
        x -= step
    elif _i % 4 == 3:
        y -= step
    return APoint(x, y)


def recur(_p, step, layer):
    if layer == 50:
        return
    _p2 = next_p(_p, layer, step)
    acad.model.AddLine(_p, _p2)
    layer += 1
    step += 5
    print(step)
    recur(_p2, step, layer)

recur(p, 0, 1)


''' 基本的遍历图形中所有图元的方法 '''
for _obj in acad.iter_objects():
    # # ObjectName 可以打印出对象的类型
    print(_obj.ObjectName)


''' 按类型查找出所有某种图元(如所有Text对象） '''
for _text in acad.iter_objects('Text'):
    print(_text.TextString, _text.InsertionPoint)

''' 在类型选择时填入多种类型 '''
for _obj2 in acad.iter_objects(['tExt', 'Line']):
    # 按照类型查找可以混淆大小写，也可以只输入类型的一部分，
    # 比如查找”te”类型就可以自动匹配到text类型，输入”li“就自动匹配到Ellipse和Line类型
    print(_obj2.ObjectName)


def text_contains_3(text_obj):
    """ 查找符合条件的第一个对象. 查找第一个text item包含3的text """
    return '3' in text_obj.TextString

text2 = acad.find_one('Text', predicate=text_contains_3)
print(text2.TextString)


# 在文档中修改对象, 需要找到 interesting objects 并改变其属性.
# 一些属性被描述为 constants, 例如 text alignment.这些 constants 可通过 ACAD. 来访问
''' 改变所有text objects的text alignment '''
for text in acad.iter_objects('Text'):
    # text.InsertionPoint 转 APoint, 存储text的插入点
    # 当 setting properties时, 不能直接使用tuple, 例如 text.TextAlignmentPoint.
    old_insertion_point = APoint(text.InsertionPoint)
    text.Alignment = ACAD.acAlignmentRight  # 对齐方式调整为 右对齐
    text.TextAlignmentPoint = old_insertion_point  # 恢复插入点位置


# 改变对象位置要用 APoint,
''' 改变line端点位置 '''
for line in acad.iter_objects('Line'):
    ''' 该代码用在图形中只有直线时，有多段线时要精确查找对象为 AcDbLine(直线) AcDbPolyline(多段线)'''
    line.EndPoint = APoint(line.StartPoint) - APoint(20, 0)

for _ob in acad.iter_objects('AcDbText'):
    _ob.move(APoint(0, 0), APoint(100, 100))
    print(_ob.ObjectName)

''' 调用move方法(set) '''
for _text in acad.iter_objects('AcDbText'):
    _text.move(APoint(0, 0), APoint(100, 100))  # 原点，相对原点的位置


''' 访问layer '''
for _text in acad.iter_objects('AcDbText'):
    print(_text.layer)

''' 设置layer '''
for _text in acad.iter_objects('AcDbText'):
    _text.layer = "0"  # 要改变text对象的layer，直接赋值即可(layer名字必须已经存在，否则会报错)

''' 提取PolyLine的各个顶点 '''
for _poly in acad.iter_objects('AcDbPolyline'):
    pC = _poly.Coordinates
    # (0.0, 0.0, 10., 10.0, 20.0, 30.0, 30., 80.0)
    # 第1、2个元素构成第一个坐标(x,y)， 3、4个元素构成第二个坐标(x,y)

''' 使用table Excel '''
# save text and position from all text objects to Excel file, and then load it back.

# add some objects to AutoCAD
acad = Autocad()
p1 = APoint(0, 0)
for i in range(5):
    obj = acad.model.AddText(u'Hi %s!' % i, p1, 2.5)
    p1.y += 10


#  iterate this objects and save them to Excel table
table = Table()
for obj in acad.iter_objects('Text'):
    x, y, z = obj.InsertionPoint
    table.writerow([obj.TextString, x, y, z])
table.save('data.xls', 'xls')

# After saving this data to ‘data.xls’ and probably changing it
# with some table processor software (e.g. Microsoft Office Excel)
# we can retrieve our data from file
data = Table.data_from_file('data.xls')


def print_table_info(table, print_rows=0):
    """ Example of working with AutoCAD table objects at examples/dev_get_table_info.py """
    merged = set()
    column_widths = [round(table.GetColumnWidth(col), 2) for col in range(table.Columns)]
    row_heights = [round(table.GetRowHeight(row), 2) for row in range(table.Rows)]
    row_texts = []
    for row in range(table.Rows):
        columns = []
        for col in range(table.Columns):
            if print_rows > 0:
                columns.append(table.GetText(row, col))
            minRow, maxRow, minCol, maxCol, is_merged = table.IsMergedCell(row, col)
            if is_merged:
                merged.add((minRow, maxRow, minCol, maxCol,))
        if print_rows > 0:
            print_rows -= 1
            row_texts.append(columns)

    print('row_heights = %s' % str(row_heights))
    print('column_widths = %s' % str(column_widths))
    print('merged_cells = %s' % print.pformat(list(merged)))
    if row_texts:
        print('content = [')
        for row in row_texts:
            print(u"        [%s]," % u", ".join("u'%s'" % s for s in row))
        print(']')


acad = Autocad()
layout = acad.doc.ActiveLayout
table = acad.find_one('table', layout.Block)
print_table_info(table, 3)


# ActiveX technology is quite slow.
# When you are accessing object attributes like position, text, etc, every time call is passed to AutoCAD.
# It can slowdown execution time.
# For example if you have program, which combines single line text based on its relative positions,
# you probably need to get each text position several times.
# To speed this up,
# you can cache objects attributes using the pyautocad.cache.Cached proxy (see example in class documentation)

# To improve speed of AutoCAD table manipulations,
# you can use Table.RegenerateTableSuppressed = True or handy context manager suppressed_regeneration_of(table):
'''
table = acad.model.AddTable(pos, rows, columns, row_height, col_width)
with suppressed_regeneration_of(table):
    table.SetAlignment(ACAD.acDataRow, ACAD.acMiddleCenter)
    for row in range(rows):
        for col in range(columns):
            table.SetText(row, col, '%s %s' % (row, col))'''


''' api - Main Autocad interface '''

''' types - 3D Point and other AutoCAD data types '''

''' utils - Utility functions '''

''' contrib.tables - Import and export tabular data from popular formats '''

''' cache - Cache all object’s attributes '''

'''
Module

pyautocad.api
pyautocad.cache
pyautocad.contrib.tables
pyautocad.types
pyautocad.utils
'''


"""
属性/方法 第一种形式（可以直接调用）

AddRef
Application
ArrayPolar
ArrayRectangular
AttachmentPoint
BackgroundFill
Copy
Database
Delete
Document
DrawingDirection
EntityName
EntityTransparency
EntityType
Erase
FieldCode
GetBoundingBox
GetExtensionDictionary
GetIDsOfNames
GetTypeInfo
GetTypeInfoCount
GetXData
Handle
HasExtensionDictionary
Height
Highlight
Hyperlinks
InsertionPoint
IntersectWith
Invoke
Layer
LineSpacingDistance
LineSpacingFactor
LineSpacingStyle
Linetype
LinetypeScale
Lineweight
Material
Mirror
Mirror3D
Move
Normal
ObjectID
ObjectID32
ObjectName
OwnerID
OwnerID32
PlotStyleName
QueryInterface
Release
Rotate
Rotate3D
Rotation
ScaleEntity
SetXData
StyleName
TextString
TransformBy
TrueColor
Update
Visible
Width

# 属性/方法 第二种形式
_AddRef
_GetIDsOfNames
_GetTypeInfo
_IAcadEntity__com_ArrayPolar
_IAcadEntity__com_ArrayRectangular
_IAcadEntity__com_Copy
_IAcadEntity__com_GetBoundingBox
_IAcadEntity__com_Highlight
_IAcadEntity__com_IntersectWith
_IAcadEntity__com_Mirror
_IAcadEntity__com_Mirror3D
_IAcadEntity__com_Move
_IAcadEntity__com_Rotate
_IAcadEntity__com_Rotate3D
_IAcadEntity__com_ScaleEntity
_IAcadEntity__com_TransformBy
_IAcadEntity__com_Update
_IAcadEntity__com__get_EntityName
_IAcadEntity__com__get_EntityTransparency
_IAcadEntity__com__get_EntityType
_IAcadEntity__com__get_Hyperlinks
_IAcadEntity__com__get_Layer
_IAcadEntity__com__get_Linetype
_IAcadEntity__com__get_LinetypeScale
_IAcadEntity__com__get_Lineweight
_IAcadEntity__com__get_Material
_IAcadEntity__com__get_PlotStyleName
_IAcadEntity__com__get_TrueColor
_IAcadEntity__com__get_Visible
_IAcadEntity__com__get_color
_IAcadEntity__com__set_EntityTransparency
_IAcadEntity__com__set_Layer
_IAcadEntity__com__set_Linetype
_IAcadEntity__com__set_LinetypeScale
_IAcadEntity__com__set_Lineweight
_IAcadEntity__com__set_Material
_IAcadEntity__com__set_PlotStyleName
_IAcadEntity__com__set_TrueColor
_IAcadEntity__com__set_Visible
_IAcadEntity__com__set_color
_IAcadMText__com_FieldCode
_IAcadMText__com__get_AttachmentPoint
_IAcadMText__com__get_BackgroundFill
_IAcadMText__com__get_DrawingDirection
_IAcadMText__com__get_Height
_IAcadMText__com__get_InsertionPoint
_IAcadMText__com__get_LineSpacingDistance
_IAcadMText__com__get_LineSpacingFactor
_IAcadMText__com__get_LineSpacingStyle
_IAcadMText__com__get_Normal
_IAcadMText__com__get_Rotation
_IAcadMText__com__get_StyleName
_IAcadMText__com__get_TextString
_IAcadMText__com__get_Width
_IAcadMText__com__set_AttachmentPoint
_IAcadMText__com__set_BackgroundFill
_IAcadMText__com__set_DrawingDirection
_IAcadMText__com__set_Height
_IAcadMText__com__set_InsertionPoint
_IAcadMText__com__set_LineSpacingDistance
_IAcadMText__com__set_LineSpacingFactor
_IAcadMText__com__set_LineSpacingStyle
_IAcadMText__com__set_Normal
_IAcadMText__com__set_Rotation
_IAcadMText__com__set_StyleName
_IAcadMText__com__set_TextString
_IAcadMText__com__set_Width
_IAcadObject__com_Delete
_IAcadObject__com_Erase
_IAcadObject__com_GetExtensionDictionary
_IAcadObject__com_GetXData
_IAcadObject__com_SetXData
_IAcadObject__com__get_Application
_IAcadObject__com__get_Database
_IAcadObject__com__get_Document
_IAcadObject__com__get_Handle
_IAcadObject__com__get_HasExtensionDictionary
_IAcadObject__com__get_ObjectID
_IAcadObject__com__get_ObjectID32
_IAcadObject__com__get_ObjectName
_IAcadObject__com__get_OwnerID
_IAcadObject__com__get_OwnerID32
_IDispatch__com_GetIDsOfNames
_IDispatch__com_GetTypeInfo
_IDispatch__com_GetTypeInfoCount
_IDispatch__com_Invoke
_IUnknown__com_AddRef
_IUnknown__com_QueryInterface
_IUnknown__com_Release
_Invoke
_QueryInterface
_Release
__class__
__cmp__
__com_interface__
__ctypes_from_outparam__
__del__
__delattr__
__dict__
__doc__
__eq__
__format__
__getattr__
__getattribute__
__hash__
__init__
__map_case__
__metaclass__
__module__
__new__
__nonzero__
__reduce__
__reduce_ex__
__repr__
__setattr__
__setstate__
__sizeof__
__str__
__subclasshook__
__weakref__
_b_base_
_b_needsfree_
_case_insensitive_
_compointer_base__get_value
_idlflags_
_iid_
_invoke
_methods_
_needs_com_addref_
_objects
_type_
color
from_param
value
"""

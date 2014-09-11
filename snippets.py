#!/usr/bin/env python
# -*- mode: python; coding: utf-8 -*-
# (c) Valik mailto:vasnake@gmail.com

'''
Created on 2011-04-30
@author: Valik

сниппеты для работы с Автокадом

Проблема с чертежами МВК
    отрицательные координаты можно убрать выполнив преобразование такое:
    (trans '(-3195.939915040071400 1786.635070675984300) (handent "7598") 0)
    что означает - из OCS в WCS. Чтобы получить UCS (ocs2ucs) надо сделать так:
    (trans '(-3195.939915040071400 1786.635070675984300) (handent "7598") 1)
Проблема с ними в том, что в чертежах используются системы координат OCS, UCS, WCS
и все они по разному повернуты относительно друг друга. Это доставляет много хлопот
при определении углов поворота (например для блоков). Хлопоты усугубляются тем, что
API ActiveX Автокада не содержит функций для преобразования углов и, хуже того,
выдает координаты всегда в WCS (за исключением полилиний) а углы всегда в OCS.
Приходится выкручиваться. Мой совет - не используйте API ActiveX Автокада.
Используйте AutoLISP.

c:\Program Files\Common Files\Autodesk Shared\acadauto.chm
Coordinate
    Variant (three-element or two-element array of doubles); read-write
    The array of X, Y, and Z coordinates for the specified vertex.
    LightweightPolyline object: The variant has two elements representing the X and Y coordinates in OCS.
    Polyline object: The variant has three elements, representing the X and Y coordinates in OCS. The Z coordinate is present in the variant but ignored.
    All other objects: The variant has three elements, representing the X and Y coordinates in WCS; the Z coordinate will default to 0 on the active UCS.
Coordinates
    Variant (array of doubles); read-write
    The array of points.
    LightweightPolyline objects: The variant is an array of 2D points in OCS.
    Polyline objects: The variant is an array of 3D points: the X and Y coordinates are in OCS; the Z coordinate is ignored.
    All other objects: The variant is an array of 3D points in WCS.

http://exchange.autodesk.com/autocadarchitecture/enu/online-help/browse#WS73099cc142f4875516d84be10ebc87a53f-79d0.htm
OCS
    Object coordinate system—point values returned by entget are expressed in this coordinate system,
    relative to the object itself.
    These points are usually converted into the WCS, current UCS, or current DCS, according to the intended use of the object.
    Conversely, points must be translated into an OCS before they are written to the database by means of the entmod or entmake
    unctions. This is also known as the entity coordinate system.
WCS
    World coordinate system: The reference coordinate system. All other coordinate systems are
    defined relative to the WCS, which never changes. Values measured relative to the WCS are stable
    across changes to other coordinate systems. All points passed in and out of
    ActiveX methods and properties are expressed in the WCS unless otherwise specified.
UCS
    User coordinate system (UCS): The working coordinate system. The user specifies a UCS to make drawing
    tasks easier. All points passed to AutoCAD commands, including those returned from AutoLISP routines
    and external functions, are points in the current UCS (unless the user precedes them
    with an * at the Command prompt). If you want your application to send coordinates in the
    WCS, OCS, or DCS to AutoCAD commands, you must first convert them to the UCS by calling the
    TranslateCoordinates method.

http://exchange.autodesk.com/autocadarchitecture/enu/online-help/search#WS73099cc142f4875516d84be10ebc87a53f-7a23.htm
    IAcadUCS_Impl GetUCSMatrix

http://www.kxcad.net/autodesk/autocad/Autodesk_AutoCAD_ActiveX_and_VBA_Developer_Guide/ws1a9193826455f5ff1a32d8d10ebc6b7ccc-6d46.htm
    util = doc.Utility # def TranslateCoordinates(self, Point, FromCoordSystem, ToCoordSystem, Displacement, OCSNormal):
    coordinateWCS = ThisDrawing.Utility.TranslateCoordinates(firstVertex, acOCS, acWorld[acUCS], False, plineNormal)

'''

def getModule(sModuleName):
    import comtypes.client
    sLibPath = GetLibPath()
    comtypes.client.GetModule(sLibPath +'\\'+ sModuleName)

def GetLibPath():
    """Return location of Autocad type libraries as string
    c:\Program Files\Common Files\Autodesk Shared\
    c:\ObjectARX2011\inc-win32\
    """
    import _winreg, os, sys
    key = _winreg.OpenKey(_winreg.HKEY_CLASSES_ROOT, \
        "TypeLib\\{D32C213D-6096-40EF-A216-89A3A6FB82F7}\\1.0\\0\\win32")
    res = _winreg.QueryValueEx(key, '')[0]
    print >> sys.stderr, ('lib [%s]' % (res))
    return os.path.dirname(res)

def NewObj(MyClass, MyInterface):
    """Creates a new comtypes POINTER object where\n\
    MyClass is the class to be instantiated,\n\
    MyInterface is the interface to be assigned"""
    from comtypes.client import CreateObject
    try:
        ptr = CreateObject(MyClass, interface=MyInterface)
        return ptr
    except:
        return None

def CType(obj, interface):
    """Casts obj to interface and returns comtypes POINTER or None"""
    try:
        newobj = obj.QueryInterface(interface)
        return newobj
    except:
        return None

def point2str(xyz):
    return u'%0.16f, %0.16f' % (xyz[0], xyz[1])

class VacEntity (object):
    ''' EntityType adapter (c:\program...\Autodesk Topobase Client 2011\Help\acadauto.chm)
    базовый класс для работы с координатами примитивов (координаты приходят в WCS, за исключением особых случаев)
    и доп.свойствами
    '''
    def __init__(self, item=''):
        super(VacEntity, self).__init__()
        self.coords = ''
        self.angle = ''
        self.name = ''
        self.closed = ''
        self.radius = ''

    def toStr(self):
        return u'%s;%s;%s;%s;%s' % (self.coords, self.angle, self.name, self.closed, self.radius)
    def __str__(self):
        return self.toStr()
    def __repr__(self):
        return self.toStr()

    def description(self):
        s = u'Общая проблема экспорта координат заключается в том, что координаты нам дают в WCS (обычно)\n\
  а углы определены для OCS (обычно). К тому, в Автокаде нет функций перевода углов из одной КС в другую.\n\
  Примеры восстановления примитивов из вынутых данных см.в коде тестов.\n\
coords это список координат (WCS) вида х, у[, ...] в зависимости от типа элемента: \n\
  для блока - точка вставки, доп.точка (д.т.) вектора с углом поворота блока, д.т. вектора с углом 0, д.т. вект. с углом 90 град.;\n\
  список точек для линий;\n\
  список точек и булжей для полилиний;\n\
  для текста - InsertionPoint, TextAlignmentPoint, доп.точка (д.т.) вектора с углом поворота текста, д.т. вектора с углом 0, д.т. вект. с углом 90 град.;\n\
  центр для окружности;\n\
  Center, StartPoint, EndPoint, MidPoint для дуги.\n\
  Поскольку значения углов Автокадом не переводятся из одной КС в другую, приходится углы, заданные в OCS, определять\n\
  через дополнительные точки в WCS. Доп.точка вместе с точкой вставки дает вектор с определенным углом поворота\n\
  вокруг точки вставки, относительно направления оси X. Вектор (точка вставки, доп.точка вектора с углом 0) дает направление оси Х в OCS.\n\
  Точка полилинии может предваряться булжем - (bulge f) x, y - значит эта и след.точки образуют дуговой сегмент;\n\
  значение булжа записано для WCS.\n\
  Дуга всегда рисуется против часовой стрелки (start-end), но это справедливо только в CS дуги (OCS),\n\
  поэтому я сохраняю три точки дуги, однозначно ее определяющие (при этом старт может оказаться эндом).\n\
angle это Rotation для блока и текста; StartAngle, EndAngle для дуг. Углы в радианах OCS если не указано другое.\n\
text это имя для блока; текст надписи для текста.\n\
closed для полилиний True если замкнутая полилиния (полигон);\n\
    для текста - имя стиля (StyleName).\n\
radius это радиус для окружности и дуги; для блока - XScaleFactor, YScaleFactor;\n\
    для текста - Alignment, VerticalAlignment, HorizontalAlignment, Height, ScaleFactor, Backward.'
        return s

    def heads(self):
        return u'coords, angle, text, closed, radius'
    def values(self):
        return u'%s//%s//%s//%s//%s' % (self.coords, self.angle, self.name, self.closed, self.radius)

    def getWCSpointsFromOCSangle(self, pnt, norm, angle=0.0):
        ''' возвращает три точки, определяющие направления осей OCS и угол в OCS.
        Три точки дают три вектора, исходящих из pnt - точка вставки блока, к примеру.
        cx - направление оси Х, cy - оси У, sp - угол поворота.
        Эта информация потом используется для преобразования угла поворота в таковой для других СК.
        '''
        p = VAcad.trans(pnt, AutoCAD.acWorld, AutoCAD.acOCS, norm)
        sp = trig.AutoLISP.polarP(p, angle, 100)
        cx = trig.AutoLISP.polarP(p, 0.0, 100)
        cy = trig.AutoLISP.polarP(p, math.pi/2.0, 100)
        sp = VAcad.trans(sp, AutoCAD.acOCS, AutoCAD.acWorld, norm)
        cx = VAcad.trans(cx, AutoCAD.acOCS, AutoCAD.acWorld, norm)
        cy = VAcad.trans(cy, AutoCAD.acOCS, AutoCAD.acWorld, norm)
        return (sp,cx,cy)
#	def getWCSpointsFromOCSangle(self, pnt, norm, angle=0.0):
#class VacEntity


class VacBlock (VacEntity):
    ''' ##class IAcadBlockReference
    acBlockReference = 7
    доп.атрибуты: HasAttributes,GetAttributes,GetConstantAttributes,IsDynamicBlock,GetDynamicBlockProperties,
    InsUnitsFactor,X(Y)ScaleFactor,X(Y)EffectiveScaleFactor,InsUnits,EffectiveName.
    coords: InsertionPoint, SecondPoint, X-axis, Y-axis
    Три точки после точки вставки дают (для WCS) угол поворота, направление осей Х и У соответственно.
    '''
    def __init__(self, item=''):
        super(VacBlock, self).__init__(item)
        if not item: return
        o = CType(item, AutoCAD.IAcadBlockReference)
        self.coords = o.InsertionPoint # point2str(o.InsertionPoint)
        self.angle = o.Rotation
        self.name = o.Name
        self.radius = u'%s, %s' % (o.XScaleFactor, o.YScaleFactor)

        norm = o.Normal
        sp,cx,cy = self.getWCSpointsFromOCSangle(self.coords, norm, self.angle)
        self.coords = u'%s, %s, %s, %s' % (point2str(self.coords), point2str(sp), point2str(cx), point2str(cy))
        #~ self.angle = u'%s, %s' % (self.angle, VAcad.ocs2wcsAngle(self.angle, norm))
#	def __init__(self, item=''):
#class VacBlock (VacEntity)

class VacLWPolyline (VacEntity):
    ''' ##class IAcadLWPolyline
    # acPolylineLight = 24
    доп.атрибуты: Thickness,ConstantWidth
    coords: x, y[,...] или (bulge n) x, y[,...] и добавляется еще одна точка для замыкания, если есть атрибут Closed.
    bulge это характеристика каждой точки (сегмента, начинающегося с данной точки),
    если булж != 0.0 то сегмент суть дуга, определяемая через две точки сегмента и булж.
    Сохраняются данные в WCS, хотя в самом AutoCAD координаты полилиний даны в OCS, булж определен для для OCS
    (все остальные координаты в WCS).
    coordinateWCS = ThisDrawing.Utility.TranslateCoordinates(firstVertex, acOCS, acWorld, False, plineNormal)
    coordinateUCS = ThisDrawing.Utility.TranslateCoordinates(firstVertex, acOCS, acUCS, False, plineNormal)
    '''
    def __init__(self, item=''):
        super(VacLWPolyline, self).__init__(item)
        if not item: return
        o = CType(item, AutoCAD.IAcadLWPolyline)
        self.closed = o.Closed
        self.coords = o.Coordinates
        if self.closed:
            # если клозед, то надо добавить еще одну точку в конец, идентичную первой, и проверить последний булж!
            self.coords = self.coords + (self.coords[0], self.coords[1])

        llen = len(self.coords)
        norm = o.Normal
        bsign = self.getWCSBulgeSign(norm)
        s = u''
        for n,e in zip(range(llen), self.coords):
            if n%2 == 0: # 0, 2, 4,...
                if s: s += u', '
                p = VAcad.trans((e, self.coords[n+1], 0.0), AutoCAD.acOCS, AutoCAD.acWorld, norm)
                if n < llen-2: # not last pair
                    ind = n/2
                    b = bsign * o.GetBulge(ind)
                    if not b == 0.0:
                        s += u'(bulge %0.5f) ' % b
                        print '  polyline [%s] have bulge [%0.3f] at segment [%u]' % (o.Handle, b, ind+1)
                s += point2str(p)
        self.coords = s
#	def __init__(self, item=''):

    def getWCSBulgeSign(self, norm):
        ''' надо правильно определить знак булжа в WCS, ибо изначально он для OCS.
        Метод детектирования смены направления хода часовой стрелки
        основан на измерении углов, известных для OCS.
        '''
        p0 = VAcad.trans((0.0, 0.0, 0.0), AutoCAD.acOCS, AutoCAD.acWorld, norm)
        p1 = VAcad.trans((2.0, 1.0, 0.0), AutoCAD.acOCS, AutoCAD.acWorld, norm)
        p2 = VAcad.trans((2.0, 2.0, 0.0), AutoCAD.acOCS, AutoCAD.acWorld, norm)
        return trig.getBulgeSign(p0, p1, p2)
#class VacLWPolyline (VacEntity)


class VacText (VacEntity):
    ''' ##class IAcadText
    # acText = 32
    доп.атрибуты: Thickness,Height,VerticalAlignment,ObliqueAngle,HorizontalAlignment,
        StyleName,ScaleFactor,Alignment
    To position text whose justification is other than left, aligned, or fit, use the TextAlignmentPoint.
    coords: InsertionPoint, TextAlignmentPoint, SecondPoint, X-axis, Y-axis
    расшифровку см. в описании VacBlock.
    closed: StyleName
    radius: Alignment, VerticalAlignment, HorizontalAlignment, Height, ScaleFactor, Backward
    '''
    def __init__(self, item=''):
        super(VacText, self).__init__(item)
        if not item: return
        o = CType(item, AutoCAD.IAcadText)
        self.coords = (o.InsertionPoint, o.TextAlignmentPoint) # u'%s, %s' % (point2str(o.InsertionPoint), point2str(o.TextAlignmentPoint))
        self.name = o.TextString
        self.angle = o.Rotation
        self.radius = u'%s, %s, %s, %s, %s, %s' % \
            (o.Alignment, o.VerticalAlignment, o.HorizontalAlignment, o.Height, o.ScaleFactor, o.Backward)
        self.closed = o.StyleName

        norm = o.Normal
        sp,cx,cy = self.getWCSpointsFromOCSangle(self.coords[0], norm, self.angle)
        self.coords = u'%s, %s, %s, %s, %s' % (point2str(self.coords[0]), point2str(self.coords[1]), \
            point2str(sp), point2str(cx), point2str(cy))
        #~ self.angle = u'%s, %s' % (self.angle, VAcad.ocs2wcsAngle(self.angle, norm))
#class VacText (VacEntity)

class VacLine (VacEntity):
    ''' ##class IAcadLine
    # acLine = 19
    доп.атрибуты: Thickness
    '''
    def __init__(self, item=''):
        super(VacLine, self).__init__(item)
        if not item: return
        o = CType(item, AutoCAD.IAcadLine)
        #~ self.coords = (o.StartPoint, o.EndPoint)
        self.coords = u'%s, %s' % (point2str(o.StartPoint), point2str(o.EndPoint))
#class VacLine (VacEntity)

class VacCircle (VacEntity):
    ''' ##class IAcadCircle
    # acCircle = 8
    доп.атрибуты: Thickness
    '''
    def __init__(self, item=''):
        super(VacCircle, self).__init__(item)
        if not item: return
        o = CType(item, AutoCAD.IAcadCircle)
        #~ self.coords = o.Center
        self.coords = point2str(o.Center)
        self.radius = o.Radius
#class VacCircle (VacEntity)

class VacArc (VacEntity):
    ''' ##class IAcadArc
    # acArc = 4
    доп.атрибуты: Thickness
    coords: Center, StartPoint, EndPoint, MidPoint
    Поскольку углы из OCS в WCS не преобразуются, определение дуги пишется четырьмя точками.
    MidPoint это точка на середине дуги.
    c:\Program Files\Common Files\Autodesk Shared\acadauto.chm
        An arc is always drawn counterclockwise from the start point to the endpoint.
        The StartPoint and EndPoint properties of an arc are calculated through the
        StartAngle, EndAngle, and Radius properties.
    Направление дуги определяется для UCS а координаты в WCS.
    '''
    def __init__(self, item=''):
        super(VacArc, self).__init__(item)
        if not item: return
        o = CType(item, AutoCAD.IAcadArc)
        c,s,e = (o.Center, o.StartPoint, o.EndPoint)
        sa,ea = (o.StartAngle, o.EndAngle)
        self.coords = u'%s, %s, %s' % (point2str(c), point2str(s), point2str(e)) # WCS
        self.radius = o.Radius
        self.angle = u'%s, %s' % (sa, ea) # radians
# get midpoint
        norm = o.Normal
        c = VAcad.trans(c, AutoCAD.acWorld, AutoCAD.acOCS, norm)
        s = VAcad.trans(s, AutoCAD.acWorld, AutoCAD.acOCS, norm)
        e = VAcad.trans(e, AutoCAD.acWorld, AutoCAD.acOCS, norm)
        m = trig.getArcMidpointP(c, self.radius, s, e)
        m = VAcad.trans(m, AutoCAD.acOCS, AutoCAD.acWorld, norm)
        self.coords = u'%s, %s' % (self.coords, point2str(m))
#class VacArc (VacEntity)

class VacPoint (VacEntity):
    ''' ##class IAcadPoint
    acPoint = 22
    доп.атрибуты: Thickness
    '''
    def __init__(self, item=''):
        super(VacPoint, self).__init__(item)
        if not item: return
        o = CType(item, AutoCAD.IAcadPoint)
        self.coords = point2str(o.Coordinates)
#class VacPoint (VacEntity):


class VAcadServices:
    ''' для полилиний необходимо трансформировать координаты из OCS в WCS.
    '''
    def __init__(self):
        import comtypes.client
        self.acad = comtypes.client.GetActiveObject('AutoCAD.Application')
        self.ac = CType(self.acad, AutoCAD.IAcadApplication)
        self.docs = self.ac.Documents
        self.doc = self.ac.ActiveDocument
        self.ms = self.doc.ModelSpace
        self.u = self.doc.Utility
        self.bulgeSign = ''
        self.zeroAngle = ''
        self.norm = ''

    def openDWG(self, fname, ro=True):
        self.docs.Close()
        self.doc = self.docs.Open(fname, ro)
        self.ms = self.doc.ModelSpace
        self.u = self.doc.Utility

    def trans(self, point, csFrom, csTo, norm='', disp=False):
        if len(point) < 3:
            point = (point[0], point[1], 0.0)
        p = array.array('d', [point[0], point[1], point[2]])
        if norm:
            norm = array.array('d', [norm[0], norm[1], norm[2]])
            p = self.u.TranslateCoordinates(p, csFrom, csTo, disp, norm)
        else:
            p = self.u.TranslateCoordinates(p, csFrom, csTo, disp)
        return p


    def ocs2wcsAngle(self, angle, norm):
        ''' transform angle from OCS to WCS
        '''
        #~ if not (self.norm and self.bulgeSign) or self.norm != norm:
        p0 = self.trans((0.0, 0.0, 0.0), AutoCAD.acOCS, AutoCAD.acWorld, norm)
        p1 = self.trans((2.0, 1.0, 0.0), AutoCAD.acOCS, AutoCAD.acWorld, norm)
        p2 = self.trans((2.0, 2.0, 0.0), AutoCAD.acOCS, AutoCAD.acWorld, norm)
        p10 = self.trans((1.0, 0.0, 0.0), AutoCAD.acOCS, AutoCAD.acWorld, norm)
        self.bulgeSign = trig.getBulgeSign(p0, p1, p2)
        self.zeroAngle = trig.AutoLISP.angleP(p0, p10)
        self.norm = norm

        ta = (self.bulgeSign * self.zeroAngle) + (self.bulgeSign * angle)
        return trig.normAngle2pi(ta)


    def getUCSMatrix(self):
        ''' Autodesk Shared\acadauto.chm
        RetVal = object.GetUCSMatrix()
        Object
            UCS (Document.ActiveUCS)
            The object this method applies to.
        RetVal
            Variant (4x4 array of doubles)
            The UCS matrix.
        To transform an entity into a given UCS, use the TransformBy method, using the matrix returned by this method as the input for that method.
        http://exchange.autodesk.com/autocadarchitecture/enu/online-help/browse#WS73099cc142f4875516d84be10ebc87a53f-7a23.htm
            matrix: (a1,a2,a3) (b1,b2,b3), (c1,c2,c3), (d1,d2,d3)
            x' = x*a1 + y*b1 + z*c1 + d1
            y' = x*a2 + y*b2 + z*c2 + d2
            z' = x*a3 + y*b3 + z*c3 + d3
        Example:
            #~ UCSMatrix: ((0.0, 1.0, 0.0), (1.0, 0.0, 0.0), (0.0, 0.0, 0.0), (0.0, 0.0, 0.0))
            #~ WCS (bulge 0.26052) 3195.9399150400714000, 1786.6350706759843000,
            #~ UCS at point  X=1786.6351  Y=3195.9399  Z=   0.0000
            a,b,c,d = (0.0, 1.0, 0.0), (1.0, 0.0, 0.0), (0.0, 0.0, 0.0), (0.0, 0.0, 0.0)
            x,y,z = 3195.9399150400714000, 1786.6350706759843000, 0.0
            xt = x*a[0] + y*b[0] + z*c[0] + d[0]
            yt = x*a[1] + y*b[1] + z*c[1] + d[1]
            zt = x*a[2] + y*b[2] + z*c[2] + d[2]
        '''
        m = None
        print 'getUCSMatrix...'
        t = self.doc.GetVariable('UCSNAME')
        print 'UCSNAME: [%s]' % t
        if t:
            t = self.doc.ActiveUCS
            print 'ActiveUCS: [%r]' % t
        if t:
            ucs = CType(t, AutoCAD.IAcadUCS)
            print 'IAcadUCS: [%r]' % ucs
            m = ucs.GetUCSMatrix()
            m = (self.doc.GetVariable('UCSNAME'), m[0], m[1], m[2], m[3])
        if not m:
            m = (self.doc.GetVariable('UCSNAME'), \
                self.doc.GetVariable('UCSXDIR'), self.doc.GetVariable('UCSYDIR'), \
                (0.0, 0.0, 0.0), self.doc.GetVariable('UCSORG'))
        return m
#class VAcadServices:


try:
    import comtypes.gen.AutoCAD as AutoCAD # c:\Python25\Lib\site-packages\comtypes\gen\_D32C213D_6096_40EF_A216_89A3A6FB82F7_0_1_0.py
except:
    getModule('acax18ENU.tlb') # comtypes.gen.AutoCAD
    import comtypes.gen.AutoCAD as AutoCAD # c:\Python25\Lib\site-packages\comtypes\gen\_D32C213D_6096_40EF_A216_89A3A6FB82F7_0_1_0.py
import math, array
import trig

VAcad = VAcadServices()

#!/usr/bin/env python
# -*- mode: python; coding: utf-8 -*-
# (c) Valik mailto:vasnake@gmail.com

'''
Created on 2011-05-28
@author: Valik

Python >= 2.5

Утилиты по преобразованию координат и примитивов, вынутых ранее из AutoCAD
'''

import os, sys, math
import time, traceback

cp = 'utf-8'
ecErr = 1
ecOK = 0


class AutoLISP:
    ''' AutoLISP functions
    '''
    @staticmethod
    def angle(x1,y1, x2,y2):
        ''' AutoLISP angle in range [0 - 2Pi]
        Returns the angle between line [p1,p2] and X axis.
        http://www.afralisp.net/reference/autolisp-functions.php#A
        (angle <PT1> <PT2>) Returns the angle in radians between two points (no way!)
        Second version:
            angle = math.acos((x1 - cx) / radius)
            if not sign(y1 - cy) > 0:
                angle = (2.0 * math.pi) - angle
        '''
        res = math.atan2(y2-y1, x2-x1)
        if res < 0:
            res = (2*math.pi) + res
        return res

    @staticmethod
    def angleP(p1, p2):
        return AutoLISP.angle(p1[0],p1[1], p2[0],p2[1])

    @staticmethod
    def polar(x1,y1, phi, dist):
        ''' AutoLISP polar
        This function (polar) returns the point at an angle (in radians) and distance from a given point
        http://www.afralisp.net/reference/autolisp-functions.php#P
        (polar <PT><ANGLE><DISTANCE>)
        '''
        x = x1 + dist * math.cos(phi)
        y = y1 + dist * math.sin(phi)
        return (x,y)

    @staticmethod
    def polarP(pnt, phi, dist):
        return AutoLISP.polar(pnt[0], pnt[1], phi, dist)
#class AutoLISP:

def pyAngle(x1,y1, x2,y2):
    ''' Returns the angle between line [p1,p2] and X axis.
    Range [-Pi - Pi]
    '''
    return math.atan2(y2-y1, x2-x1)

def sign(n):
    if n < 0.0: return -1.0
    return 1.0

def floatIsEqual(v1, v2, accuracy=0.001):
    if abs(abs(float(v1)) - abs(float(v2))) > accuracy:
        return False
    if abs(float(v1) - float(v2)) <= accuracy:
        return True
    if abs((-1.0*float(v1)) + float(v2)) <= accuracy:
        return True
    return False

def normAngle2pi(angle):
    ''' return angle in range [0 - 2*pi]
    '''
    if angle >= 0.0 and angle <= math.pi*2:
        return angle
    if angle < 0.0:
        return normAngle2pi(angle + math.pi*2)
    return normAngle2pi(angle - math.pi*2)

def rotationAngle(cx, cy, centerP, origAngle):
    ''' Get mirroring mark and rotation angle (in current CS) for block or text,
    based on two vectors and saved angle.
    @type	cx: tuple
    @param	cx: Point (x,y) for vector (centerP,cx). This vector is origin zero angle vector.
    @type	cy: tuple
    @param	cy: Point (x,y) for vector (centerP,cy). This vector is origin PI/2 angle vector.
    @type	centerP: tuple
    @param	centerP: Center Point (x,y) for two vectors, e.g. this is insertion point for block.
    @type	origAngle: number
    @param	origAngle: Block or text rotation angle (0-2*PI) in OCS.

    @rtype: tuple
    @returns: (zDir, rotAngle) Z-axis direction and block rotation angle in current CS.
    '''
    zDir = 1.0
    ucsA = origAngle
#	print 'rotationAngle: cp [%s], cx [%s], cy [%s]' % (centerP, cx, cy)
    xa = AutoLISP.angleP(centerP, cx)
    ya = AutoLISP.angleP(centerP, cy)
#	print 'rotationAngle: xa [%s], ya [%s]' % (math.degrees(xa), math.degrees(ya))
    # norm. x,y angles
    nxa = xa
    nya = ya
    if abs(nxa - nya) > math.radians(90.5):
        nxa = normAngle2pi(xa + math.pi)
        nya = normAngle2pi(ya + math.pi)
#	print 'rotationAngle: nxa [%s], nya [%s]' % (math.degrees(nxa), math.degrees(nya))
    # detect Z-axis direction
    if (nya - nxa) < 0:
        zDir = -1.0
    # calc angle
    ucsA = (origAngle * zDir) + (xa * zDir)
    return (zDir, ucsA)
#def rotationAngle(cx, cy, centerP, origAngle):


class Vwcs2ucs:
    '''
    ucsMatrix = ((0.0, 1.0, 0.0), (1.0, 0.0, 0.0), (0.0, 0.0, 0.0), (0.0, 0.0, 0.0))
    # detect origin angle for UCS
    zp = wcs2ucs(ucsMatrix, 0.0, 0.0)
    p10 = wcs2ucs(ucsMatrix, 1.0, 0.0)
    ucsZA = AutoLISP.angleP(zp, p10)
    # detect clock hands directions for UCS
    ucsRotSign = getUCSBulgeSign(ucsMatrix)
    print 'ucs angle origin [%s] deg, rotation sign [%s]' % (math.degrees(ucsZA), ucsRotSign)
    ucsDlt = math.radians(360)
    # if a <= ucsZA: ucsDlt = 0.0
    # ucsA = ucsZA + (a * ucsRotSign) + ucsDlt
    '''
    def __init__(self):
        self.ucsMatrix = ((1.0, 0.0, 0.0), (0.0, 1.0, 0.0), (0.0, 0.0, 1.0), (0.0, 0.0, 0.0))
        self.ucsZA = 0.0
        self.ucsDlt = math.radians(360)
        self.bulgeSign = 1.0

    def config(self, matrix):
        self.ucsMatrix = matrix
        # detect clock hands directions for UCS
        p0 = self.wcs2ucs(0.0, 0.0)
        p1 = self.wcs2ucs(2.0, 1.0)
        p2 = self.wcs2ucs(2.0, 2.0)
        self.bulgeSign = getBulgeSign(p0, p1, p2)
        # detect origin angle for UCS
        p10 = self.wcs2ucs(1.0, 0.0)
        self.ucsZA = AutoLISP.angleP(p0, p10)
        print 'ucs angle origin [%s] deg, rotation sign [%s]' % (math.degrees(self.ucsZA), self.bulgeSign)

    def wcs2ucsAngle(self, a):
        ta = (self.bulgeSign * self.ucsZA) + (self.bulgeSign * a)
        return normAngle2pi(ta)

    def getUCSBulgeSign(self):
        return self.bulgeSign

    def wcs2ucs(self, x, y, z=0.0):
        ''' transform point from WCS to UCS
        http://exchange.autodesk.com/autocadarchitecture/enu/online-help/browse#WS73099cc142f4875516d84be10ebc87a53f-7a23.htm
        (a1,a2,a3) (b1,b2,b3), (c1,c2,c3), (d1,d2,d3)
        xt = x*a1 + y*b1 + z*c1 + d1
        yt = x*a2 + y*b2 + z*c2 + d2
        zt = x*a3 + y*b3 + z*c3 + d3
        ucs2wcx (http://ru.wikipedia.org/wiki/%D0%9C%D0%B5%D1%82%D0%BE%D0%B4_%D0%9A%D1%80%D0%B0%D0%BC%D0%B5%D1%80%D0%B0):
        x = delta1/delta
        y = delta2/delta
        z = delta3/delta
        delta = a[0]*b[1]*c[2] - a[0]*c[1]*b[2] - b[0]*a[1]*c[2] + b[0]*c[1]*a[2] + c[0]*a[1]*b[2] - c[0]*b[1]*a[2]
        if (delta <> 0)
            delta1 = (xt - d[0])*b[1]*c[2] - (xt - d[0])*c[1]*b[2] - b[0]*(yt - d[1])*c[2] + b[0]*c[1]*(zt - d[2]) + c[0]*(yt - d[1])*b[2] - c[0]*b[1]*(zt - d[2])
            delta2 = a[0]*(yt - d[1])*c[2] - a[0]*c[1]*(zt - d[2]) - (xt - d[0])*a[1]*c[2] + (xt - d[0])*c[1]*a[2] + c[0]*a[1]*(zt - d[2]) - c[0]*(yt - d[1])*a[2]
            delta3 = a[0]*b[1]*(zt - d[2]) - a[0]*(yt - d[1])*b[2] - b[0]*a[1]*(zt - d[2]) + b[0]*(yt - d[1])*a[2] + (xt - d[0])*a[1]*b[2] - (xt - d[0])*b[1]*a[2]
        '''
        matrix = self.ucsMatrix
        a,b,c,d = matrix
        xt = x*a[0] + y*b[0] + z*c[0] + d[0]
        yt = x*a[1] + y*b[1] + z*c[1] + d[1]
        zt = x*a[2] + y*b[2] + z*c[2] + d[2]
        return (xt, yt, zt)
#	def wcs2ucs(self, x, y, z=0.0):

    def wcs2ucsP(self, pnt):
        z = 0.0
        if len(pnt) > 2: z = pnt[2]
        return self.wcs2ucs(pnt[0], pnt[1], z)
#class Vwcs2ucs:


def getBulgeSign(p0, p1, p2):
    ''' Detect clockhand directions reversion for new CS.
    Angle p0-p1 < p0-p2 in first CS, check how true is it in new CS.
    ref. snippets.VacLWPolyline.getWCSBulgeSign(self, norm)
    '''
    res = 1.0
    p1a = AutoLISP.angleP(p0, p1)
    p2a = AutoLISP.angleP(p0, p2)
    if p1a > p2a: res = -1.0
    return res

def getUCSBulgeSign(ucsMatrix):
    t = Vwcs2ucs()
    t.config(ucsMatrix)
    return t.getUCSBulgeSign()

def wcs2ucs(matrix, x, y, z=0.0):
    ''' transform point from WCS to UCS
    '''
    t = Vwcs2ucs()
    t.config(matrix)
    return t.wcs2ucs(x,y,z)

def wcs2ucsP(matrix, pnt):
    z = 0.0
    if len(pnt) > 2: z = pnt[2]
    return wcs2ucs(matrix, pnt[0], pnt[1], z)


def unzipBulge(x1, y1, x2, y2, bulge, sublen=0.0, algo=1):
    ''' Convert AutoCAD polyline segment with bulge to arc (radius, center, angles, start-stop points)
    and to approximating line segments (facets).
    If sublen == 0 then facetlen = 2*pi*r / 30
    Formula sources:
    att\FacetBulge_rev1.zip\FacetBulge\BulgeCalc_Rev1.xls (wrong formula)
    http://www.cadtutor.net/forum/showthread.php?51511-Points-along-a-lwpoly-arc&
    http://www.afralisp.net/archive/lisp/Bulges1.htm (good formula)
Because a circular arc describes a portion of the circumference of a circle, it has all the attributes of a circle:
    Radius (r) is the same as in the circle the arc is a portion of.
    Center point (P) is also the same as in the circle.
    Included angle (θ). In a circle, this angle is 360 degrees.
    Arc length (le). The arc length is equal to the perimeter in a full circle.
Adding to these attributes are some that are specific for an arc:
    Start point and end point (P1 and P2) a.k.a. vertices (although sometimes it is practical to talk about specific points that a circle passes through, there are no distinct vertices on the circumference of a circle).
    Chord length (c). An infinite amount of chords can be described by both circles and arcs, but for an arc there is only one distinct chord that passes through its vertices (for a circle, there is only one distinct chord that passes through the center, the diameter, but it doesn't describe any specific vertices).
    Given two fixed vertices, there is also a specific midpoint (P3) of an arc.
    The apothem (a). This line starts at the center and is perpendicular to the chord.
    The sagitta (s) a.k.a. height of the arc. This line is drawn from the midpoint of an arc and perpendicular to its chord.
The bulge is the tangent of 1/4 of the included angle for the arc between the
    selected vertex and the next vertex in the polyline's vertex list.
    A negative bulge value indicates that the arc goes clockwise from the selected vertex to the next vertex.
    A bulge of 0 indicates a straight segment, and a bulge of 1 is a semicircle.
    '''
    print 'calcBulge: p1 [%s], p2 [%s], bulge [%s], sublen [%s]' % ((x1,y1), (x2,y2), bulge, sublen)
    res = {}
# midpoint (F9, G9) # arc midpoint [(12.41925, 13.5352)]; // 16.73, 11.77
    mx1 = (x1 + x2) / 2.0
    my1 = (y1 + y2) / 2.0
    s1 = 'chord midpoint [%s]' % ((mx1, my1),)
    res['chordmidpoint'] = (mx1, my1)
# angle (K13) # angle/2 (K14)
    angle = math.atan(bulge) * 4.0
    angleDeg = angle * (180.0 / math.pi)
    s2 = 'arc angle [%0.5f] rad, [%0.5f] deg' % (angle, angleDeg)
    res['angleRad'] = angle
    res['angleDeg'] = angleDeg
# dist (F5)
    dist = math.sqrt((x2-x1)**2 + (y2-y1)**2)
    s3 = 'chord len [%0.5f]' % dist
    res['chordLen'] = dist
# sagitta length (http://en.wikipedia.org/wiki/Sagitta_%28geometry%29)
    sagitta = dist/2.0 * bulge
# radius (K15) # r = ((dist/2.0)**2+sagitta**2)/2.0*sagitta
    radius = 1.e99
    if not angle == 0.0:
        radius = (dist/2.0) / math.sin( abs(angle/2.0) )
    if radius == 0.0: radius = 1.e-99
    s4 = 'radius [%0.5f]' % radius
    res['radius'] = radius
# arc length (F14) l = 2*pi*r
    alen = abs(radius * angle)
    s5 = 'arc length [%0.5f]' % alen
    res['arcLen'] = alen
# center (F19, G19)
    t = '''
    bad algo:
        k5 = ( (math.sqrt(radius**2 - (dist / 2.0)**2))*2 ) / dist
        cx = mx1 + (((y1-y2)/2.0) * k5 * sign(bulge))
        cy = my1 + (((x2-x1)/2.0) * k5 * sign(bulge))
        s6 = 'arc center [%s]' % ((cx, cy),)
        res['center'] = (cx, cy)
    good algo:
    (setq bulge 2.5613 p1 (list 11.7326 11.8487) p2 (list 13.1059 15.2217) r 2.68744
      theta (* 4.0 (atan (abs bulge)))
      gamma (/ (- pi theta) 2.0)
      phi (+ (angle p1 p2) gamma)
      p (polar p1 phi r) )
    '''
    theta = 4.0 * math.atan(abs(bulge))
    gamma = (math.pi - theta) / 2.0
    #~ phi = AutoLISP.angle(x1,y1, x2,y2) + gamma
    phi = pyAngle(x1,y1, x2,y2) + gamma * sign(bulge)
    cx,cy = AutoLISP.polar(x1,y1, phi, radius)
    s6 = 'arc center [%s]' % ((cx, cy),)
    res['center'] = (cx, cy)
# start, end angle (G21, G22)
    startAngle = math.acos((x1 - cx) / radius)
    if not sign(y1 - cy) > 0:
        startAngle = (2.0 * math.pi) - startAngle
    endAngle = startAngle + angle
    s7 = 'start, end arc angles [%s]' % ((startAngle, endAngle),)
    res['seAngles'] = (startAngle, endAngle)
# subangle (F27), numsub (# of Divisions K26)
    if sublen <= 0.0:
        numsub = abs(angle / 0.2) # 30 segments for full circle
        try: sublen = alen / numsub
        except: sublen = alen
    try: numsub = round(alen/sublen, 0)
    except: numsub = 1
    if numsub < 2:
        numsub = 1
    subangle = angle / numsub
    s8 = 'numsub, subangle [%s]' % ((numsub, subangle),)
# length of subarc L27
    realSublen = abs(2 * (math.sin(subangle/2.0) * radius) )
    s9 = 'real sublen [%0.5f]' % realSublen
# sub points
    listPoints = [(x1,y1)]
    t = ''' diff between 2 algo:
    unzipBulge(50.0, 0.0, -50.0, 0.0, 1, 1)
    first algo
    -49.989990187139483, 1.0004404476901365)
    second algo
    -49.989990187142631, 1.0004404476947728
    '''
    if algo == 1: # first algo
        currangle = startAngle + (subangle/2.0) + (math.pi/2.0 * sign(bulge))
        sx = x1
        sy = y1
        for cnt in range(int(numsub-1)):
            if not cnt == 0:
                currangle = currangle + subangle
            sx = sx + (realSublen * math.cos(currangle))
            sy = sy + (realSublen * math.sin(currangle))
            listPoints.append((sx, sy))
    else: # second algo
        for cnt in range(int(numsub-1)):
            currangle = startAngle + abs((cnt+1)*subangle) * sign(bulge)
            sx,sy = AutoLISP.polar(cx,cy, currangle, radius)
            listPoints.append((sx, sy))

    listPoints.append((x2, y2))
    res['points'] = listPoints
    print 'doBulge done [%s; %s; %s; %s; %s; %s; %s; %s; %s; subpoints %s]' % (s1,s2,s3,s4,s5,s6,s7,s8,s9,listPoints)
    return res
#def unzipBulge(x1, y1, x2, y2, bulge, sublen=0.0, algo=1):

def unzipBulge2(p1, p2, bulge, sublen):
    ''' p1, p2 - points like (x,y,z)
    '''
    return unzipBulge(p1[0], p1[1], p2[0], p2[1], bulge, sublen)


def getArcBulge(center, start, end, startangle=-1, endangle=-1):
    ''' get bulge value for AutoCAD arc.
    It works only if start and end are in correct positions.
    center, start, end - is points like (x,y)
    An arc is always drawn counterclockwise from the start point to the endpoint (in UCS!).
    The bulge is the tangent of 1/4 of the included angle for the arc between the
    selected vertex and the next vertex.
    A negative bulge value indicates that the arc goes clockwise from the selected vertex to the next vertex.
    '''
    if startangle == -1:
        startangle = AutoLISP.angle(center[0], center[1], start[0],start[1])
        endangle = AutoLISP.angle(center[0], center[1], end[0],end[1])
    print 'startangle [%s], endangle [%s]' % (math.degrees(startangle), math.degrees(endangle))
    ang = endangle - startangle
    if ang < 0: ang = math.radians(360) + ang
    bulge = math.tan(ang/4.0)
    return bulge


def normArcAngles(startangle, endangle, delta=0.0, back=False):
    ''' in normal form arc always have endangle > startangle
    So, let's rotate pie.
    '''
    if back:
        startangle = math.radians(360) - delta
        endangle -= delta
        return (startangle, endangle, 0.0)

    delta = 0.0
    if startangle > endangle:
        delta = math.radians(360) - startangle
        endangle += delta
        startangle = 0.0
    return (startangle, endangle, delta)
#def normArcAngles(startangle, endangle, delta=0.0, back=False):

def getArcMidpointP(center, radius, startpoint, endpoint):
    startangle = AutoLISP.angle(center[0], center[1], startpoint[0],startpoint[1])
    endangle = AutoLISP.angle(center[0], center[1], endpoint[0],endpoint[1])
    return getArcMidpointA(center, radius, startangle, endangle)

def getArcMidpointA(center, radius, startangle, endangle):
    startangle, endangle, delta = normArcAngles(startangle, endangle)
    a = ((endangle - startangle) / 2.0) - delta + startangle
    x,y = AutoLISP.polar(center[0], center[1], a, radius)
    return (x,y,0)

def detectArcStartEnd(center, start, end, midpoint):
    ''' s,e = detectArcStartEnd(c, s, e, m)
    An arc is always drawn counterclockwise from the start point to the endpoint.
    '''
    startangle = AutoLISP.angle(center[0], center[1], start[0],start[1])
    endangle = AutoLISP.angle(center[0], center[1], end[0],end[1])
    midangle = AutoLISP.angle(center[0], center[1], midpoint[0],midpoint[1])
    startangle, endangle, delta = normArcAngles(startangle, endangle)
    midangle += delta
    if midangle > math.radians(360):
        midangle -= math.radians(360)
    if midangle > startangle and midangle < endangle:
        return (start, end)
    return (end, start)


################################################################################


def test (pl, norm):
    if pl == norm:
        return
    raise NameError('test NOT passed sample [%s], paragon [%s]' % (pl, norm))

def testArcMidpoint():
    ''' 'DRAWING3';'AcDbArc';'4';'0';'2127698424';'
    377';'';'
    7.2943541954524846, 7.6562227951962285,
    6.5885851277066587, 10.6607789870368300,
    6.4885743713744901, 4.6769297981878202,
    8.0502341927713292, 10.6485652416030180';'1.80151314583, 4.44824837595';'';'';'3.08633567308'
    '''
    c = (7.2943541954524846, 7.6562227951962285)
    s = (6.5885851277066587, 10.6607789870368300)
    e = (6.4885743713744901, 4.6769297981878202)
    r = 3.08633567308
    m = getArcMidpointP(c, r, s, e)
    test(m, (4.2084494996092445, 7.7077989049644016, 0))

def testUCSMatrix():
    ''' http://exchange.autodesk.com/autocadarchitecture/enu/online-help/browse#WS73099cc142f4875516d84be10ebc87a53f-7a23.htm
(a1,a2,a3) (b1,b2,b3), (c1,c2,c3), (d1,d2,d3)
x = x*a1 + y*b1 + z*c1 + d1
y = x*a2 + y*b2 + z*c2 + d2
z = x*a3 + y*b3 + z*c3 + d3

(setq h (handent "7598") o (command "zoom" "o" h "") o (redraw h 3))

WCS
'+01+02';'AcDbPolyline';'24';'В_ТЕКСТ_УЗЛЫ';'2127975424';'7598';'';'
    (bulge 0.26052) 3195.9399150400714000, 1786.6350706759843000,
    (bulge 0.58605) 3196.4541498123590000, 1785.7151197495145000,
    3194.8100081997472000, 1784.7941722803723000
    ';'';'';'False';''
UCSMatrix: ((0.0, 1.0, 0.0), (1.0, 0.0, 0.0), (0.0, 0.0, 0.0), (0.0, 0.0, 0.0))
UCS
at point  X=1786.6351  Y=3195.9399  Z=   0.0000
at point  X=1785.7151  Y=3196.4541  Z=   0.0000
at point  X=1784.7942  Y=3194.8100  Z=   0.0000
    '''
    print
    print 'testUCSMatrix...'
    a,b,c,d = (0.0, 1.0, 0.0), (1.0, 0.0, 0.0), (0.0, 0.0, 0.0), (0.0, 0.0, 0.0)
    x,y,z = 3195.9399150400714000, 1786.6350706759843000, 0.0
    xt = x*a[0] + y*b[0] + z*c[0] + d[0]
    yt = x*a[1] + y*b[1] + z*c[1] + d[1]
    zt = x*a[2] + y*b[2] + z*c[2] + d[2]
    print 'a [%s], b [%s], c [%s], d [%s]' % (a,b,c,d)
    print 'x [%s], y [%s], z [%s]' % (x,y,z)
    print 'xt [%s], yt [%s], zt [%s]' % (xt,yt,zt)
    print
#def testUCSMatrix():

def testAngle():
    ''' result
    45.0
    135.0
    180.0
    225.0
    315.0
    '''
    print
    print 'testAngle...'
    print math.degrees(AutoLISP.angle(0,0, 1,1))
    print math.degrees(AutoLISP.angle(0,0, -1,1))
    print math.degrees(AutoLISP.angle(0,0, -1,0))
    print math.degrees(AutoLISP.angle(0,0, -1,-1))
    print math.degrees(AutoLISP.angle(0,0, 1,-1))
    print
#def testAngle():


def testTrig():
    testArcMidpoint()
    testAngle()
    testUCSMatrix()
    return ecOK

if __name__ == '__main__':
    argc = len(sys.argv)
    res = ecErr
    print 'begin test [%s], argc: [%s], argv: [%s]' % (time.strftime('%Y-%m-%d %H:%M:%S'), argc, sys.argv)
    try:
        res = testTrig()
        print 'done [%s]' % res
    except Exception, e:
        if type(e).__name__ == 'COMError': print 'COM Error, msg [%s]' % e
        else:
            print 'Error, tests failed:'
            traceback.print_exc(file=sys.stderr)
    print 'end test [%s], argc: [%s], argv: [%s]' % (time.strftime('%Y-%m-%d %H:%M:%S'), argc, sys.argv)
    sys.exit(res)

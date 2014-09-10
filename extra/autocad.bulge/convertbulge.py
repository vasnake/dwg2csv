#!/usr/bin/env python
# -*- coding: utf-8 -*-
# (c) Valik mailto:vasnake@gmail.com

import sys, os, math



class AutoLISP:
	''' AutoLISP functions analog
(command "zoom" "o" (handent "7598") "") (redraw (handent "7598") 3)
(setq o (command "zoom" "o" (handent "7598") "") o1 (redraw (handent "7598") 3))
	'''
	@staticmethod
	def angle(x1,y1, x2,y2):
		''' AutoLISP angle
		Returns the angle between line [p1,p2] and X axis.
		http://www.afralisp.net/reference/autolisp-functions.php#A
		(angle <PT1> <PT2>) Returns the angle in radians between two points (no way!)
		'''
		return math.atan2(y2-y1, x2-x1)

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
#class AutoLISP:

def testBulge(x1, y1, x2, y2, bulge, sublen=0.0):
	''' Convert AutoCAD polyline segment with bulge to arc (radius, center, angles, start-stop points)
	and to approximating line segments (facets).
	If sublen == 0 then facetlen = arclen / 30
	Formula sources:
	att\FacetBulge_rev1.zip\FacetBulge\BulgeCalc_Rev1.xls (wrong formula)
	http://www.cadtutor.net/forum/showthread.php?51511-Points-along-a-lwpoly-arc&
	http://www.afralisp.net/archive/lisp/Bulges1.htm (good formula)
	'''
	print 'calcBulge: p1 [%s], p2 [%s], bulge [%s], sublen [%s]' % ((x1,y1), (x2,y2), bulge, sublen)
	res = {}
	def sign(n):
		if n < 0.0: return -1.0
		return 1.0
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
  p (polar p1 phi r)
)
	'''
	theta = 4.0 * math.atan(abs(bulge))
	gamma = (math.pi - theta) / 2.0
	phi = AutoLISP.angle(x1,y1, x2,y2) + gamma * sign(bulge)
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
	if sublen <= 0.0 and alen > 0.0: sublen = alen / 30.0
	numsub = 0
	if sublen > 0.0: numsub = round(alen/sublen, 0)
	if numsub < 2:
		numsub = 1.0
	subangle = angle / numsub
	s8 = 'numsub, subangle [%s]' % ((numsub, subangle),)
# length of subarc L27
	realSublen = abs(2 * (math.sin(subangle/2.0) * radius) )
	s9 = 'real sublen [%0.5f]' % realSublen
# sub points
	listPoints = [(x1,y1)]
	t = '''
diff between 2 algo:
testBulge(50.0, 0.0, -50.0, 0.0, 1, 1)
first algo
-49.989990187139483, 1.0004404476901365)
second algo
-49.989990187142631, 1.0004404476947728
	'''
	#~ t = ''' first algo
	currangle = startAngle + (subangle/2.0) + (math.pi/2.0 * sign(bulge))
	sx = x1
	sy = y1
	for cnt in range(int(numsub-1)):
		if not cnt == 0:
			currangle = currangle + subangle
		sx = sx + (realSublen * math.cos(currangle))
		sy = sy + (realSublen * math.sin(currangle))
		listPoints.append((sx, sy))
	#~ '''
	t = ''' second algo
	for cnt in range(int(numsub-1)):
		currangle = startAngle + (cnt+1)*subangle * sign(bulge)
		sx = cx + radius * math.cos(currangle)
		sy = cy + radius * math.sin(currangle)
		listPoints.append((sx, sy))
	'''
	listPoints.append((x2, y2))
	res['points'] = listPoints
	print 'doBulge done [%s; %s; %s; %s; %s; %s; %s; %s; %s; subpoints %s]' % (s1,s2,s3,s4,s5,s6,s7,s8,s9,listPoints)
	return res
#def testBulge(self, x1, y1, x2, y2, bulge, sublen=10):


def main():
	print >> sys.stderr, 'argc: [%s], argv: [%s]' % (len(sys.argv), sys.argv)
	#~ проверить дуговые сегменты и дуги в чертежах - что это есть и как они интерполируются
	#~ сделать интерполяцию дуг (не сегментов а дуг - центр, две точки + радиус)
	#~ как видны внешние блоки? чем блоки отличаются от референсов?
	#~ грузить в Оракл
	#~ видуализировать выборку в Автокад
	testBulge(13.0, 9.0, 21.68, 33.65, -0.45, 7)
	#~ testBulge(11.7326, 11.8487, 13.1059, 15.2217, 2.5613, 1)
	#~ testBulge(5.9228, 12.0274, 8.1062, 22.8887, 2.2893, 3)
	#~ testBulge(24.2884, 10.3276, 27.3493, 14.9170, -4.5373, 3)
	#~ testBulge(-34.8952, -21.6100, -32.7117, -10.7486, 2.2893, 5)
	#~ testBulge(-16.5296, -23.3098, -13.4687, -18.7203, -4.5373, 5)
	#~ testBulge(10.7004, -22.2329, 12.8839, -11.3716, 2.2893, 2)
	#~ testBulge(29.0660, -23.9327, 32.1270, -19.3433, -4.5373, 2)
	#~ testBulge(-34.2720, 12.0274, -32.0886, 22.8887, 2.2893, 2)
	#~ testBulge(-15.9064, 10.3276, -12.8455, 14.9170, -4.5373, 2)
	#~ testBulge(8.0555, 5.0696, -6.8401, -6.4998, -0.3324, 2)
	#~ testBulge(-3.6195, 4.3654, 6.0425, -6.4998, 2.1734, 2)
	#~ testBulge(6.0425, -6.4998, -3.6195, 4.3654, -2.1734, 2)
	#~ testBulge(1.0, 0.0, 0.0, 1.0, 1, 0.5)
	#~ testBulge(0, 1, -1, 0, 1, 0.1)
	#~ testBulge(1.0, 0.0, -1.0, 0.0, 1, 0.03)
	#~ testBulge(50.0, 0.0, -50.0, 0.0, 1, 1)
	#~ testBulge(1.0, 0.0, 0.99, -0.044, 44.53, 0.1)
	#~ testBulge(0.0, 0.0, 0.0, 0.0, 45, 0.3)
	#~ testBulge(0.0, 0.0, 1.0, 1.0, 0.0)



import time, traceback
print time.strftime('%Y-%m-%d %H:%M:%S')
if __name__ == "__main__":
	try:
		main()
	except Exception, e:
		if type(e).__name__ == 'COMError': print 'COM Error, msg [%s]' % e
		else:
			print 'Error, program failed:'
			traceback.print_exc(file=sys.stderr)
print time.strftime('%Y-%m-%d %H:%M:%S')

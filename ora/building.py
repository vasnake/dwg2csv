#!/usr/bin/env python
# -*- mode: python; coding: utf-8 -*-
# (c) Valik mailto:vasnake@gmail.com

'''
Created on 2011-07-04
@author: Valik

Python >= 2.5
cx_Oracle-5.1-10g.win32-py2.5.msi

Task
Get buildings data from raw data table (xdata like '%99:44110000%') and
insert rows into featureclass table building.

After loading we need to repair geometry
select * from (
  select a.fid, sdo_geom.validate_geometry_with_context( a.geom, 0.00001 ) res
  from
  MKV.building
  a order by res
) where not res = 'TRUE' ;
alter session set current_schema=MKV;
begin
  SDO_MIGRATE.TO_CURRENT('BUILDING');
end;
/
update building a
  set a.geom = sdo_util.rectify_geometry(a.geom, 0.00001)
  where not (sdo_geom.validate_geometry_with_context( a.geom, 0.00001 ) = 'TRUE') ;

Tables
CREATE TABLE "MKV"."BIGTAB"
  (
    "FID"      NUMBER(10,0),
    "CLASSIF"  VARCHAR2(80 CHAR),
    "CLOSTY"   VARCHAR2(10 CHAR),
    "COMMENTS" VARCHAR2(2000 CHAR),
    "COORDS" CLOB,
    "DWG"      VARCHAR2(255 CHAR),
    "EID"      VARCHAR2(12 CHAR),
    "HAND"     VARCHAR2(10 CHAR),
    "LYR"      VARCHAR2(80 CHAR),
    "RAD"      VARCHAR2(40 CHAR),
    "ROTANG"   VARCHAR2(40 CHAR),
    "STATS"    VARCHAR2(80 CHAR),
    "TXT"      VARCHAR2(200 CHAR),
    "TYPENAME" VARCHAR2(32 CHAR),
    "TYPENUM"  NUMBER(3,0),
    "XDATA"    VARCHAR2(400 CHAR)
  ) ;
'''

import os, sys, time
import traceback
import csv
import cx_Oracle

USERNAME = 'MKV'
PASSWORD = os.environ.get('as2217_cgisdb_rgogrid')
TNSENTRY = 'tb12'
ARRAY_SIZE = 50

cp = 'utf-8'
ecErr = 1
ecOK = 0


def OutputTypeHandler(cursor, name, defaultType, size, precision, scale):
    if defaultType == cx_Oracle.CLOB:
        return cursor.var(cx_Oracle.LONG_STRING, 70000, cursor.arraysize)


class VoraDriver:
    def __init__(self):
        self.connection = cx_Oracle.connect(USERNAME, PASSWORD, TNSENTRY)
        self.connection.autocommit = False
        self.cursor = self.connection.cursor()
        self.cursor.arraysize = ARRAY_SIZE

    def __del__(self):
        del self.cursor
        del self.connection
#class VoraDriver:


def doWork(inp='', dryrun=True):
    ora = VoraDriver()
    ora.connection.outputtypehandler = OutputTypeHandler
    print ora.connection.encoding
    ss = u'%99:44110000%'
    #~ ora.cursor.execute("""select * from MKV.bigtab where xdata like :p_Value""", p_Value = ss)
    ora.cursor.execute('''select coords, fid
        from MKV.bigtab
        where xdata like :p_Value
        order by fid''',
        p_Value = ss
    )
    selected = []
    for coords, fid in ora.cursor:
        #~ print 'len [%s], xy [%s]' % (len(coords), coords[:99])
        selected.append((coords, fid))

    for recnum,rec in zip(range(len(selected)),selected):
        coords, fid = rec
        print 'recnum [%s], reclen [%s], xy [%s...]' % (recnum, len(coords), coords[:33])

        wkt = coords.split(', ')
        points = []
        for n,e in zip(range(len(wkt)), wkt):
            if n%2 == 0: # 0, 2, 4,...
                points.append((e, wkt[n+1]))
        wkt = ''
        for p in points:
            if wkt: wkt += ', '
            wkt += '%s %s' % (p[0], p[1])
        geom = 'POLYGON ((%s))' % wkt

        ora.cursor.setinputsizes(coords = cx_Oracle.CLOB)
        ora.cursor.execute("""Insert into mkv.building (GEOM) values (
            SDO_GEOMETRY( :coords, 82353)
            )""",
            coords = geom
        )

    if dryrun:
        ora.connection.rollback()
    else:
        ora.connection.commit()
    del ora
    return ecOK
#def doWork(inp):


if __name__ == '__main__':
    argc = len(sys.argv)
    res = ecErr
    inp = ''
    print 'begin [%s], argc: [%s], argv: [%s]' % (time.strftime('%Y-%m-%d %H:%M:%S'), argc, sys.argv)
    if argc > 1: inp = sys.argv[1]

    try:
        res = doWork(inp)
        print 'done [%s]' % res
    except Exception, e:
        if type(e).__name__ == 'COMError': print 'COM Error, msg [%s]' % e
        else:
            print 'Error, doWork failed'
            traceback.print_exc(file=sys.stderr)
    print 'end [%s], argc: [%s], argv: [%s]' % (time.strftime('%Y-%m-%d %H:%M:%S'), argc, sys.argv)
    sys.exit(res)

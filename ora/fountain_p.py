#!/usr/bin/env python
# -*- mode: python; coding: utf-8 -*-
# (c) Valik mailto:vasnake@gmail.com

'''
Created on 2011-07-04
@author: Valik

Python >= 2.5
cx_Oracle-5.1-10g.win32-py2.5.msi

Task
select fountains (xdata like '00:"Фонтан питьевой";99:56045100')
from big table;
insert them to featureclass table fountain_p.

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

CREATE TABLE "MKV"."FOUNTAIN_P"
  (
    "FID" NUMBER(10,0),
    "GEOM" "MDSYS"."SDO_GEOMETRY" ,
    "ORIENTATION" NUMBER(6,3) DEFAULT 90.0 NOT NULL ENABLE,
    "Z"           NUMBER(20,8),
    "QUALITY"     NUMBER(10,0),
    "ATTRIBS"     VARCHAR2(255 CHAR),
    "BLKNAME"     VARCHAR2(20 CHAR),
    "DWG"         VARCHAR2(255 CHAR),
    "ENTTYPE"     VARCHAR2(20 CHAR),
    "HANDL"       VARCHAR2(20 CHAR),
    "LYR"         VARCHAR2(80 CHAR),
    "ROTANG"      NUMBER(9,3)
  )
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


class VoraDriver:
    def __init__(self):
        self.connection = cx_Oracle.connect(USERNAME, PASSWORD, TNSENTRY)
        self.cursor = self.connection.cursor()
        self.cursor.arraysize = ARRAY_SIZE

    def __del__(self):
        del self.cursor
        del self.connection
#class VoraDriver:


def doWork(inp='', dryrun=True):
    ora = VoraDriver()
    print ora.connection.encoding
    ss = u'%00:"Фонтан питьевой";99:56045100%'
    #~ ora.cursor.execute("""select * from MKV.bigtab where xdata like :p_Value""", p_Value = ss)
    ora.cursor.execute(
        '''select dbms_lob.substr(coords, 4000, 1) coords,
            xdata, txt, dwg, typename, hand, lyr
        from MKV.bigtab
        where xdata like :p_Value
        order by fid''',
        p_Value = ss
    )
    dat = ora.cursor.fetchall()

    for r in dat:
        #~ print r
        coords, xdata, txt, dwg, typename, hand, lyr = r
        #~ print ('lyr [%s]' % lyr.decode('cp1251')).encode(cp)
        wktpoint = coords.split(', ')
        wktpoint = '%s %s' % (wktpoint[0], wktpoint[1])
        xdata = xdata.decode('cp1251').encode(cp)
        txt = txt.decode('cp1251').encode(cp)
        lyr = lyr.decode('cp1251').encode(cp)
        sql = '''Insert into mkv.fountain_p (
                GEOM, attribs, blkname, dwg, enttype, handl, lyr)
            values (
                SDO_GEOMETRY('POINT (%s)', 82353),
                '%s', '%s', '%s', '%s', '%s', '%s');''' % (
            wktpoint, xdata, txt, dwg, typename, hand, lyr)
        print sql

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

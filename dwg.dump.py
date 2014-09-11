#!/usr/bin/env python
# -*- mode: python; coding: utf-8 -*-
# (c) Valik mailto:vasnake@gmail.com

'''
Created on 2011-04-30
@author: Valik

Python >= 2.5
comtypes >= 0.6.2

Extract spatial data and attributes from DWG file.
AutoCAD must be running and current file will be processed if not filename parameter given.

Supported entities:
    AcDbBlockReference
    AcDbPolyline
    AcDbCircle
    AcDbLine
    AcDbText
    AcDbArc
    AcDbPoint

changelog
    - transform coords to UCS or WCS, UCS transform matrix
    - find arc segments in polylines and convert them to polylines
    - convert arcs and bulges to polylines
    - check exported data by importing it back

TODO
    - export other types of entities

inspired by
    rngis.db.updater\src\regingdb.py\ (http://int.allgis.org/tracrn/browser/rn/db.updater/src/regingdb.py/regingdb.py)
    http://tjriley.infogami.com/python_autocad
    http://www.cadtutor.net/forum/showthread.php?48626-On-Python-for-autocad

extra info:
    Autodesk Topobase Client 2011\Help\acad_aag.chm
    http://usa.autodesk.com/adsk/servlet/index?siteID=123112&id=770215#vbatonet
    2011-05-03\Whitepaper.pdf (topobase\DevTV_VBA_To_VBdotNet_Migration_English.zip)
    http://usa.autodesk.com/adsk/servlet/index?id=1911627&siteID=123112
    topobase\DevTV_VBA_To_VBdotNet_Migration_English.zip
    topobase\dotnetwizard_v10_0_4_0.zip
    topobase\objectarx_2011_documentation.exe
    topobase\objectarx_2011_win_64_and_32bit.exe
'''

import os, sys, time
import traceback
import csv
import comtypes.client

from snippets import *

# c:\Python25\Lib\site-packages\comtypes\gen\_D32C213D_6096_40EF_A216_89A3A6FB82F7_0_1_0.py
import comtypes.gen.AutoCAD as AutoCAD

cp = 'utf-8'
ecErr = 1
ecOK = 0


def doWork(dwg=''):
    ''' Dump data from DWG
    '''
    print 'doWork...'
    #~ axDump()

    if dwg:
        print 'try load DWG [%s]' % dwg
        VAcad.openDWG(dwg)
        #~ acad = comtypes.client.GetActiveObject('AutoCAD.Application')
        #~ ac = CType(acad, AutoCAD.IAcadApplication)
        #~ docs = ac.Documents
        #~ docs.Close()
        #~ doc = docs.Open(dwg, True)
        #~ print 'dwg name [%s], fullname [%s], dwgprefix var [%s]' % (doc.Name, doc.FullName, doc.GetVariable('DWGPREFIX'))

    comtypesDump()
    return ecOK
# def doWork(dwg=''):


def comtypesDump():
    '''
    Enumerate objects from ModelSpace in current DWG;
    output objects data to file dwgname.csv.

    eXtended data sample:
    xd(
        (1001, 1000),
        (u'ESMA', u'/99:T1000000/00:"\u041e\u043f\u043e\u0440\u043d\u044b\u0435 \u0442\u043e\u0447\u043a\u0438"/')
    )

    dumped object sample:
    num [428], item [name [AcDbPolyline], type [24], lyr [В_ГИЛЬЗА], id [2127705120],
        attr [00:"Гильза водопровода";99:56044000;ДМ:800;МТ:жб],
        xd [/99:56044000/ДМ:800/МТ:жб/00:"Гильза водопровода"/],
        ent [(2227.5499999999993, 1791.6200000000003, 2222.6399999999994, 1798.1900000000001);;;False;]]

    stats sample:
    objects count [18809]
    layers [
        0:16;
        БРОШЕННЫЕ_СЕТИ:100;
        В_:878;
        ...
        УЛ_ПЛОЩАДКИ_ПРОЕЗДЫ:265;
        УЛ_ПОДЗ_ПЕРЕХОД:6;
        УЛ_ТРАМВАЙНЫЕ_ЛИНИИ:7]
    types [AcDbArc:1;AcDbBlockReference:5322;AcDbCircle:4;AcDbLine:7;AcDbPolyline:10421;AcDbText:3054]
    noXDCount [3257], dupIDs [0]
    '''
    print 'comtypesDump...'
    #~ acad = comtypes.client.GetActiveObject('AutoCAD.Application')
    #~ ac = CType(acad, AutoCAD.IAcadApplication) # ac = NewObj(AutoCAD.AcadApplication, AutoCAD.IAcadApplication)
    #~ print '[%s] loaded' % ac.Name
    #~ doc = ac.ActiveDocument
    #~ print 'dwg name [%s], fullname [%s], dwgprefix var [%s]' % (doc.Name, doc.FullName, doc.GetVariable('DWGPREFIX'))
    #~ ms = doc.ModelSpace

    count = VAcad.ms.Count
    print 'objects count [%i]' % count
    lyrDict = VCountStrings()
    nameDict = VCountStrings()
    idDict = VCountStrings()
    noXDCount = 0
    csvWriter = VcsvWriter('%s.csv' % VAcad.doc.Name)

    for i in range(count):
        #~ if i > 100: break
        item = VacItem(VAcad.ms.Item(i))

        if nameDict.push(item.name) == 1:
            print >> sys.stderr, (u'new type, num [%i], item [%s]' % (i+1, item)).encode(cp)
            if not item.ent:
                raise NameError('Need EntityType adapter! [%s]' % item.name)

        if lyrDict.push(item.lyr) == 1:
            print >> sys.stderr, (u'new lyr, num [%i], item [%s]' % (i+1, item)).encode(cp)
            pass

        if idDict.push('%s' % item.id) > 1:
            print >> sys.stderr, 'num [%i], duplicate ID [%s %s]' % (i+1, item.name, item.id)

        if not item.xd2[1]:
            #~ print >> sys.stderr, 'num [%i], XData is None! [%s %s]' % (i+1, item.name, item.id)
            noXDCount += 1

        if i == 0:
            csvWriter.writerow([(u'DWG file: %s\nObjects: %u\nUCSMatrix: %r\n' %
                (VAcad.doc.FullName, count, VAcad.getUCSMatrix())).encode(cp) +
                item.description(cp)])
            csvWriter.writerow([u'dwg'] + item.listHeads(cp))

        csvWriter.writerow([VAcad.doc.Name[:-4]] + item.listValues(cp))

        #~ if item.handle == '7598': break
        #~ if item.handle == '73DD': break
    #for i in range(count): for each entity

    del csvWriter
    print (u'layers [%s]' % lyrDict.toStr()).encode(cp)
    print (u'types [%s]' % nameDict.toStr()).encode(cp)
    print 'noXDCount [%i], dupIDs [%i]' % (noXDCount, count - len(idDict.dict))
    VAcad.doc.Utility.Prompt("There are " + str(count) + " objects in ModelSpace \n")
#def comtypesDump():


class VcsvWriter:
    ''' Wrapper for csv.writer

    writer = csv.writer(csv_stream, delimiter=exportDelimiter, quotechar=exportEscape,
        quoting=csv.QUOTE_ALL, lineterminator='\n')
    writer.writerow([description])
    '''
    def __init__(self, fname='t.csv'):
        self.fname = fname
        self.fileObj = open(fname, 'wb')
        self.csvWriter = csv.writer(self.fileObj,
            delimiter=';', quotechar="'", quoting=csv.QUOTE_ALL,
            lineterminator='\n')

    def __del__(self):
        del self.csvWriter
        self.fileObj.close()

    def writerow(self, row):
        self.csvWriter.writerow(row)
#class VcsvWriter:


class VacItem:
    ''' Wrapper for ACAD.ModelSpace.item.

    Unprocessed attribs: color, TrueColor, Visible, Material, Linetype, Lineweight
    '''
    def __init__(self, item=''):
        self.name = ''
        self.id = ''
        self.xd1,self.xd2 = ('', '')
        self.attr = {}
        self.lyr = ''
        self.etype = ''
        self.ent = ''
        self.handle = ''
        if item: self.configure(item)

    def toStr(self):
        return u'name [%s], type [%s], lyr [%s], id [%s], hndl [%s], attr [%s], xd [%s], ent [%s]' % \
            (self.name, self.etype, self.lyr, self.id, self.handle, self.attr2str(), self.xd2[1], self.ent)

    def __str__(self):
        return self.toStr()

    def __repr__(self):
        return self.toStr()

    def attr2str(self):
        return dict2string(self.attr)

    def description(self, codepage='utf-8'):
        s = u'typename, typenum это название и номер для EntityType Автокада.\n\
id это внутренний идент.элемента в чертеже.\n\
handle это атрибут Handle элемента.\n\
attribs это расширенные данные элемента (XData).\n%s\n' % self.ent.description()
        return s.encode(codepage)

    def listHeads(self, codepage='utf-8'):
        s = u'typename, typenum, layer, id, handle, attribs, %s' % self.ent.heads()
        return s.encode(codepage).split(', ')

    def listValues(self, codepage='utf-8'):
        s = u'%s//%u//%s//%u//%s//%s//%s' % \
            (self.name, self.etype, self.lyr, self.id, self.handle, self.attr2str(), self.ent.values())
        return s.encode(codepage).split('//')

    def configure(self, acItem):
        self.name = acItem.ObjectName
        self.etype = acItem.EntityType
        self.id = acItem.ObjectID
        self.lyr = acItem.Layer
        self.handle = acItem.Handle

        self.xd1,self.xd2 = acItem.GetXData('ESMA')
        if not self.xd2: self.xd2 = (u'ESMA', u'')
        xd = self.xd2[1].encode('cp1251').split(r'/')
        for d in xd:
            pair = d.split(r':', 1)
            if len(pair) > 1:
                self.attr[pair[0].decode('cp1251')] = pair[1].decode('cp1251')

        self.ent = self.makeEntity(self.etype, acItem)

        #~ en = acItem.EntityName
        #~ if en != self.name:
            #~ raise NameError('ObjectName != EntityName [%s %s]' % (self.name, en))
#    def configure(self, acItem):


    def makeEntity(self, etype, acItem):
        if etype == AutoCAD.acBlockReference:
            self.ent = VacBlock(acItem)
        elif etype == AutoCAD.acPolylineLight:
            self.ent = VacLWPolyline(acItem)
        elif etype == AutoCAD.acText:
            self.ent = VacText(acItem)
        elif etype == AutoCAD.acLine:
            self.ent = VacLine(acItem)
        elif etype == AutoCAD.acCircle:
            self.ent = VacCircle(acItem)
        elif etype == AutoCAD.acArc:
            self.ent = VacArc(acItem)
        elif etype == AutoCAD.acPoint:
            self.ent = VacPoint(acItem)
        else:
            self.ent = ''

        return self.ent
#    def makeEntity(self, etype, acItem):
#class VacItem:


class VCountStrings:
    ''' Strings collection with counting number of adding for every string
    '''
    def __init__(self):
        self.dict = {}

    def push(self, str):
        num = 0
        if str in self.dict:
            num = self.dict[str]
        num += 1
        self.dict[str] = num
        return num

    def toStr(self):
        return dict2string(self.dict)

    def __str__(self):
        return self.toStr()

    def __repr__(self):
        return self.toStr()
#class VCountStrings:


def dict2string(dct):
    s = u''
    for k in sorted(dct.keys()):
        if s: s += u';'
        s += u'%s:%s' % (k, dct[k])
    return s
#def dict2string(dct):


def axDump():
    ''' ActiveX interop example

    objects [{
    u'AcDbBlockReference': 5323,
    u'AcDbPolyline': 10422,
    u'AcDbCircle': 5,
    u'AcDbLine': 8,
    u'AcDbText': 3055,
    u'AcDbArc': 2}]
    '''
    print 'axDump...'
    #~ import win32com.client
    #~ acad = win32com.client.Dispatch("AutoCAD.Application")
    acad = comtypes.client.GetActiveObject("AutoCAD.Application")
    doc = acad.ActiveDocument
    print 'dwg name [%s], fullname [%s]' % (doc.Name, doc.FullName)
    ms = doc.ModelSpace

    count = ms.Count
    total_text = 0
    res = {}
    for i in range(count):
        if i > 100: break
        item = ms.Item(i)
        on = item.ObjectName
        try:
            num = res[on]
        except:
            print 'objname [%s]' % on
            num = 0
        res[on] = num + 1

        if 'text' in item.ObjectName.lower():
            total_text = total_text + 1

    print 'objects [%s]' % res
    doc.Utility.Prompt("There are " + str(total_text) + " text objects in ModelSpace\n")
    #~ raise NameError('not yet')
#def axDump():


if __name__ == '__main__':
    argc = len(sys.argv)
    res = ecErr
    dwg = ''
    print 'begin [%s], argc: [%s], argv: [%s]' % (time.strftime('%Y-%m-%d %H:%M:%S'), argc, sys.argv)
    if argc > 1: dwg = sys.argv[1]

    try:
        res = doWork(dwg)
        print 'done [%s]' % res
    except Exception, e:
        if type(e).__name__ == 'COMError': print 'COM Error, msg [%s]' % e
        else:
            print 'Error, doWork failed:'
            traceback.print_exc(file=sys.stderr)
    print 'end [%s], argc: [%s], argv: [%s]' % (time.strftime('%Y-%m-%d %H:%M:%S'), argc, sys.argv)
    sys.exit(res)

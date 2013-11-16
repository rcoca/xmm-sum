#!/usr/bin/env python

from xml.dom import minidom
import sys,os,glob
from optparse import OptionParser
import pyExcelerator
import datetime

def walkTree(root,extra=True):
    yield root
    if root.hasChildNodes():
        for f in root.childNodes:
            for x in walkTree(f,extra=extra):
                yield x
        if extra: yield root

def isTextNode(node):
    return node.nodeType==1 and node.hasAttribute('TEXT')


def PrintTree(tree):
    used = []
    for node in walkTree(tree,extra=True):
        if node in used:
            if isTextNode(node):nused=[]
            for k in node.childNodes:
                if isTextNode(k):nused.append(used.pop())
            nused=nused[::-1]
            for k in nused: print "%s->%s"%\
                    (k.getAttribute('TEXT'),node.getAttribute('TEXT'))
            
        if not node in used and isTextNode(node):used.append(node)

def DictFromTree(tree):
    tupleList=[]
    used = []
    for node in walkTree(tree,extra=True):
        if node in used:
            if isTextNode(node):nused=[]
            for k in node.childNodes:
                if isTextNode(k):nused.append(used.pop())
            nused=nused[::-1]
            for k in nused: 
                    tupleList.append((k.getAttribute('TEXT'),node.getAttribute('TEXT')))
            
        if not node in used and isTextNode(node):used.append(node)
    return tupleList



def findNodeByPath(tree,dottedpath):
    path=dottedpath.split('.')
    i=0
    for k in walkTree(tree,extra=False):
        if isTextNode(k) and k.getAttribute('TEXT').strip()==path[i].strip():
            if i==len(path)-1: return k
            else: i+=1
    return None

def LoadEstimFile(fname):
    fname=open(fname,"r")
    paths=fname.read().splitlines()
    fname.close()
    return paths

import re
def MergeEstimMMap(xmltree,paths):
    for path in paths:
        Path=path.strip()
        m=re.search("^([A-z\.0-9\:\s]+)\s+\-\s*([0-9\.?]+\s*[A-z?]+\s*)$",Path)
        if m:
            #print "found path=",path,m.groups()
            pathname=m.group(1).strip()
            estimation=m.group(2).strip()
            Node=findNodeByPath(xmltree,pathname)
            if Node:
                print "Adding estimation:%s:%s"%(Node.getAttribute("TEXT"),estimation)
                Node.setAttribute('TEXT',Node.getAttribute("TEXT")+":"+estimation)

def ExtractEstimTxt(xmltree,paths,p=True,FD=sys.stdout,ret=False):
    unestim_paths =[]
    estim_paths   =[]
    unk_paths     =[]
    head_rex="^([A-z\.0-9,\:\s/\(\)?\-&]+)\s*"
    estim_rex="\:\s*([0-9\.?]+\s*[A-z?]+\s*)$"
    for path in paths:
        m=re.search(head_rex+estim_rex,path)
        if m:
            pathname=m.group(1).strip()
            estimation=m.group(2)
            estim_paths.append((pathname,estimation))
        else:
            m=re.search(head_rex,path)
            if m:
                pathname=m.group(1).strip()
                unestim_paths.append(pathname)
            else:
                unk_paths.append(path)
            
    print >> FD, ""
    print >> FD, ""
    print >> FD, "Estimated:%d"%len(estim_paths)
    for k in estim_paths:    print >> FD, "%s\t-\t%s"%k
    print >> FD, ""
    print >> FD, ""
    print >> FD, "Not Estimated:%d"%len(unestim_paths)
    for k in unestim_paths:  print >> FD, "%s"%k
    print >> FD, "Estimated/Not estimated/[sum]/Total:%d/%d/[%d]/%d"%(len(estim_paths),len(unestim_paths),len(estim_paths)+len(unestim_paths),len(paths))
    print >> FD, "Unknown Paths"
    for k in unk_paths:  print >> FD, "%s"%k
    if ret:return estim_paths

def PostOrderWalkTree(root,extra=False):
    if extra:yield root
    if root.hasChildNodes():
        for f in root.childNodes:
            for x in PostOrderWalkTree(f,extra=extra):
                yield x
            yield f

def EnumeratePaths(xmltree):
    paths =[]
    for leaf in PostOrderWalkTree(xmltree):
        if leaf.hasChildNodes():continue
        else:
            c=leaf
            newpath=[leaf]
            while True:
                if c.parentNode:
                    c=c.parentNode
                    newpath.append(c)
                else:break
            newpath=filter(lambda x:isTextNode(x),newpath)
            if len(newpath) and newpath[0].hasChildNodes():continue
            paths.append(map(lambda x:x.getAttribute('TEXT'),newpath)[::-1])
            
    return paths

def PreOrderWalkTree(root,depth=0):
    yield root,depth
    if root.hasChildNodes():
        for f in root.childNodes:
            for x,d in PreOrderWalkTree(f,depth=depth+1):
                yield x,d

def GetDepth(root):
    d=0
    for n,dp in PreOrderWalkTree(root):
        if isTextNode(n):
            d=max(d,dp)
    return d
def GetRootDepth(root):
    d=None
    for n,dp in PreOrderWalkTree(root):
        if isTextNode(n):
            if not d:d=dp
            d=min(d,dp)
    return d


def CreateXLS(xmltree):
    wb = pyExcelerator.Workbook()
    Root = filter(lambda x:isTextNode(x),xmltree.firstChild.childNodes)[0].getAttribute('TEXT').encode('ascii')
    sh = wb.add_sheet(Root)

    color_list=[ 'aqua','orange','silver']#'white','grey' ,'red','lime','yellow','silver','white' ]
    lineNo=2
    paths =[]
    Depth=GetDepth(xmltree)
    depth=GetRootDepth(xmltree)
    style   = pyExcelerator.XFStyle()

    style.font.bold=True
    style.font.outline=True
    style.font.underline=True
    style.font.name='Arial'
    style.borders.right  = pyExcelerator.Formatting.Borders.THICK
    style.borders.top    = pyExcelerator.Formatting.Borders.THICK
    style.borders.bottom = pyExcelerator.Formatting.Borders.THICK
    style.borders.left   = pyExcelerator.Formatting.Borders.THICK

    for i in range(Depth-depth):sh.write(0,i,'Level %d'%(i+1),style)
    sh.write(0,Depth-depth,"Schedule:",style)
    sh.write(0,Depth+depth,"Detailed Schedule:",style)
    sh.write(0,Depth+depth+2,"StartDate",style)
    style.num_format_str = 'mm/dd/yyyy'
    
    sh.write(1,Depth+depth+2,datetime.datetime.now(),style)
    #print "Depth=%d,%d"%(Depth,depth)
    for node,colNo in PreOrderWalkTree(xmltree):
        if not isTextNode(node):continue
        node_text=node.getAttribute('TEXT').encode('ascii')
        #print (colNo-depth)*' ',node_text,color_list[(lineNo%len(color_list))]
        m=re.search("(^.*)\:\s*([0-9\.]+\s*[dhmw]\s*$)",node_text)

        style   = pyExcelerator.XFStyle()
        color=color_list[(lineNo%len(color_list))]
        style.pattern.pattern_fore_colour= pyExcelerator.Formatting.get_colour_val(color)
        style.pattern.pattern=pyExcelerator.Formatting.Pattern.SOLID_PATTERN

        if m:
            dur=ConvertToDay(m.group(2)).replace('d','')
            node_text=m.group(1)
            for i in range(Depth+(colNo-2*depth)+1):
                sh.write(lineNo,i,' ',style)
            style.borders.right  = pyExcelerator.Formatting.Borders.THIN
            style.borders.top    = pyExcelerator.Formatting.Borders.THIN
            style.borders.bottom = pyExcelerator.Formatting.Borders.THIN
            sh.write(lineNo,Depth+(colNo-2*depth),dur,style)
            if not node.hasChildNodes():
                sh.write(lineNo,Depth+depth,dur,style)
                sh.write(lineNo,Depth+depth+1,str(float(dur)*7.0/5.0),style)
                startname=chr(ord('A')+Depth+depth+2)
                colname=chr(ord('A')+Depth+depth+1)
                style.num_format_str = 'mm/dd/yyyy'
                formula=pyExcelerator.Formula('$%s2+$%s%d'%(startname,colname,lineNo+1))
                sh.write(lineNo,Depth+depth+2,formula,style)
                sh.col(Depth+depth+2).width = 0x24E1/2
        style.borders.right  = pyExcelerator.Formatting.Borders.NO_LINE
        style.borders.left  = pyExcelerator.Formatting.Borders.THIN
        sh.write(lineNo,(colNo-depth),node_text,style)
        sh.col(colNo-depth).width = 0x24E1
        lineNo+=1
    return wb

def ConvertToDay(arg):
    D={'d':1,'w':5,'h':1.0/8.0,'m':23.0}
    val=arg
    for scale in D.keys():
        val=val.replace(scale,"*%f"%D[scale])
    val=str(eval(val))
    return val
    
def UpdateSumEstim(xmltree):
    used=[]
    opqueue=[]
    nused=[]
    for k in PostOrderWalkTree(xmltree,extra=True):
        #repeat root of subtree walked
        if isTextNode(k):
            if not k in used:   used.append(k) #push
            else:
                #get arity operands from stack
                while k!=used[-1]: nused.append(used.pop(-1))
                xx=map(lambda x:x!=k and x.getAttribute('TEXT'),nused)
                collect=filter(None,map(lambda x:re.search("\:\s*([0-9\.]+\s*[dhmw]\s*$)",x),xx))
                collect=map(lambda x:ConvertToDay(x.groups(0)[0]),collect)

                if  k.hasChildNodes():
                    #execute op
                    childrenSum=sum(map(float,collect))
                    #inline update result
                    #print collect,childrenSum
                    newval=re.sub('\:\s*[0-9\.]+\s*[dhmw]','',k.getAttribute('TEXT'))
                    uval="%s:%3.3f d"%(newval,childrenSum)
                    k.setAttribute('TEXT',uval)
                    nused=[]
        
if __name__=='__main__':
    flist=glob.glob('*.mm')
    if len(flist)>0:f=flist[0]
    parser=OptionParser()
    parser.add_option('-u','--update-sum',dest='USum')
    parser.add_option('-t','--text',dest='Text')#default =f
    (options,args) = parser.parse_args()
    if not options.USum and not options.Text:
        parser.print_help()
        exit(2)
    elif options.USum:
	xmltree=minidom.parse(options.USum)
	UpdateSumEstim(xmltree)
	mapstring=xmltree.toxml().replace('<?xml version="1.0" ?>','')
	with open(options.USum,"wt+") as FD:
	    print >>FD,mapstring
    elif options.Text:
	xmltree=minidom.parse(options.Text)
	UpdateSumEstim(xmltree)
        wb=CreateXLS(xmltree)
        outxls=options.Text.replace('.mm','.xls')
        wb.save(outxls)

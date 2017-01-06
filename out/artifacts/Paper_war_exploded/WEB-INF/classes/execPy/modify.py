# -*- coding:utf-8 -*-

import zipfile
from lxml import etree
import re
import sys
import shutil
import time
import os
import tempfile
# import chardet

#Docx_Filename = input("please input the path of your docx:")
# Docx_Filename = "test.docx"
# Checkout_Filename='check_out1'
# ModifySpace_FileName = 'space'

Data_DirPath = sys.argv[1]#输入文件夹的路径,命令行的第二个参数

Docx_Filename = Data_DirPath + 'origin.docx'
Checkout_Filename = Data_DirPath + 'check_out1'
ModifySpace_FileName = Data_DirPath + 'space'

word_schema='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
Unicode_bt = 'utf-8'

#以下4个函数是关于 解压缩word文档和读取xml内容并按节点标签对树形结构进行迭代
def get_word_xml(docx_filename):
    zipF = zipfile.ZipFile(docx_filename)
    xml_content = zipF.read('word/document.xml')
    style_content=zipF.read('word/styles.xml')
    return xml_content,style_content

def get_xml_tree(xml_string):
    return etree.fromstring(xml_string)

def _iter(my_tree,type_char):
    for node in my_tree.iter(tag=etree.Element):
        if _check_element_is(node,type_char):
            yield node

def _check_element_is(element,type_char):
    return element.tag == '%s%s' % (word_schema,type_char)

def has_key(node,attribute):
    return '%s%s' %(word_schema,attribute) in node.keys()

def get_val(node,attribute):
    if has_key(node,attribute):
        return node.get('%s%s' %(word_schema,attribute))
    else:
        return '未获取属性值'

def read_error(filename):
    #因为文件里段落可能不是递增的，这里最好写一个排序。对以后改错有帮助
    rp = open(filename,'r')
    error_dct=[]
    for line in rp:
        group = line.split('_')
        group1 = {}
        group1['paraNum'] = group[0]
        group1['location'] = group[1]
        group1['kind'] = group[2]
        group1['type'] = group[3]
        group1['rightValue'] = group[4][:-1]#去掉'\n'
        error_dct.append(group1)
    return error_dct

def modify_rpr(r,rnode,label,rightValue):
    found_rPr = False
    for rPr in r.iter(tag=etree.Element):
        if _check_element_is(rPr, 'rPr'):
            found_rPr = True
            found_node = False
            for node in rPr:
                if _check_element_is(node, rnode):
                    found_node = True
                    node.set('%s%s' % (word_schema, label), rightValue)
                    break
            if found_rPr == True and found_node == False:
                rPr.insert(0, etree.Element('%s%s' % (word_schema, rnode)))
                node = rPr[0]
                node.set('%s%s' % (word_schema, label),rightValue)
    if not found_rPr:
        r.insert(0, etree.Element('%s%s' % (word_schema, 'rPr')))
        rPr = r[0]
        rPr.insert(0, etree.Element('%s%s' % (word_schema, rnode)))
        node = rPr[0]
        node.set('%s%s' % (word_schema, label), rightValue)

def modify(xml_tree,errorlist):
    paraNum = 0
    listCount = 0
    spacelistCount = 0
    for paragr in _iter(xml_tree,'p'):
        paraNum +=1
        if spacelistCount< len(spacelist) and paraNum == spacelist[spacelistCount]:
            spacelistCount = spacelistCount + 1
            pat = re.compile('[0-9]')
            for r in paragr.iter(tag = etree.Element):
                if _check_element_is(r,'r'):
                    for t in r.iter(tag=etree.Element):
                        if _check_element_is(t,'t'):
                            if pat.match(t.text):
                                pos = 0
                                for char in t.text:
                                    if char.isdigit() or char == '.':
                                        pos = pos+1
                                        continue
                                    else:
                                        break
                                txt=t.text
                                t.text=txt[:pos]+' '+txt[pos:]
                                t.set('%s%s' % ('{http://www.w3.org/XML/1998/namespace}','space'), 'preserve')
        while listCount<len(errorlist) and paraNum == int(errorlist[listCount]['paraNum']):
            if errorlist[listCount]['type'] == 'startWithTabs':
                for r in _iter(paragr, 'r'):
                    rtext = ""
                    for tab in _iter(r,"tab"):
                        r.remove(tab)
                    pat = re.compile(' |　+')
                    for t in _iter(r, 't'):
                        rtext = t.text
                    if len(pat.sub("", rtext)) == 0:
                        continue
                    else:
                        break
                listCount += 1
            elif errorlist[listCount]['type'] == "startWithSpace":
                pat = re.compile(' |　+')
                for r in _iter(paragr, 'r'):
                    rtext = ''
                    for t in _iter(r, 't'):
                        rtext = t.text
                    if pat.match(rtext):
                        if len(pat.sub("", rtext)) == 0:
                            for t in _iter(r, 't'):
                                t.text = ""
                        else:
                            for t in _iter(r, 't'):
                                t.text = pat.sub("", rtext)
                            break
                    else:
                        break
                listCount += 1
            elif errorlist[listCount]['type'] == 'paraAlign':
                for pPr in paragr:
                    if _check_element_is(pPr,'pPr'):
                        found_jc = False
                        #print etree.tostring(pPr,pretty_print = True)
                        for jc in pPr:
                            if _check_element_is(jc,'jc'):
                                found_jc = True
                                jc.set('%s%s' %(word_schema,'val'),errorlist[listCount]['rightValue'])
                                break
                        if not found_jc:
                            pPr.insert(0,etree.Element('%s%s' %(word_schema,'jc')))
                            jc = pPr[0]
                            jc.set('%s%s' %(word_schema,'val'),errorlist[listCount]['rightValue'])
                        #print etree.tostrin(pPr,pretty_print = True)
                        break
                listCount +=  1
            elif errorlist[listCount]['type'] == 'paraIsIntent' or errorlist[listCount]['type'] == 'paraIsIntent1':
                for pPr in paragr:
                    if _check_element_is(pPr,'pPr'):
                        found_ind = False
                        #print etree.tostring(pPr,pretty_print = True)
                        for ind in pPr:
                            if _check_element_is(ind,'ind'):
                                found_ind = True
                                if errorlist[listCount]['rightValue'] != '0':
                                    ind.set('%s%s' %(word_schema,'firstLineChars'),'200')
                                else:
                                    ind.set('%s%s' %(word_schema,'firstLineChars'),'0')
                                    ind.set('%s%s' %(word_schema,'firstLine'),'0')
                                break
                        if not found_ind:
                            pPr.insert(0,etree.Element('%s%s' %(word_schema,'ind')))
                            ind = pPr[0]
                            if errorlist[listCount]['rightValue'] != '0':
                                ind.set('%s%s' %(word_schema,'firstLineChars'),'200')
                            else:
                                ind.set('%s%s' %(word_schema,'firstLineChars'),'0')
                                ind.set('%s%s' %(word_schema,'firstLine'),'0')
                        #print etree.tostring(pPr,pretty_print = True)
                        break
                listCount += 1
            elif errorlist[listCount]['type'] == 'paraleft' or errorlist[listCount]['type'] == 'paraleftChars':
                for pPr in paragr:
                    if _check_element_is(pPr, 'pPr'):
                        found_ind = False
                        # print etree.tostring(pPr,pretty_print = True)
                        for ind in pPr:
                            if _check_element_is(ind, 'ind'):
                                found_ind = True
                                ind.set('%s%s' % (word_schema, 'leftChars'), '0')
                                ind.set('%s%s' % (word_schema, 'left'), '0')
                                break
                        if not found_ind:
                            pPr.insert(0, etree.Element('%s%s' % (word_schema, 'ind')))
                            ind = pPr[0]
                            ind.set('%s%s' % (word_schema, 'leftChars'), '0')
                            ind.set('%s%s' % (word_schema, 'left'), '0')
                        # print etree.tostring(pPr,pretty_print = True)
                        break
                listCount += 1
            elif errorlist[listCount]['type'] == 'parahanging' or errorlist[listCount]['type'] == 'parahangingChars':
                for pPr in paragr:
                    if _check_element_is(pPr, 'pPr'):
                        found_ind = False
                        # print etree.tostring(pPr,pretty_print = True)
                        for ind in pPr:
                            if _check_element_is(ind, 'ind'):
                                found_ind = True
                                ind.set('%s%s' % (word_schema, 'hangingChars'), '0')
                                ind.set('%s%s' % (word_schema, 'hanging'), '0')
                                break
                        if not found_ind:
                            pPr.insert(0, etree.Element('%s%s' % (word_schema, 'ind')))
                            ind = pPr[0]
                            ind.set('%s%s' % (word_schema, 'hangingChars'), '0')
                            ind.set('%s%s' % (word_schema, 'hanging'), '0')
                        # print etree.tostring(pPr,pretty_print = True)
                        break
                listCount +=  1
            elif errorlist[listCount]['type'] == 'paraSpace':
                for pPr in paragr:
                    if _check_element_is(pPr,'pPr'):
                        found_node = False
                        #print etree.tostring(pPr,pretty_print = True)
                        for node in pPr:
                            if _check_element_is(node,'spacing'):
                                found_node = True
                                node.set('%s%s' %(word_schema,'line'),errorlist[listCount]['rightValue'])
                                break
                        if not found_node:
                            pPr.insert(0,etree.Element('%s%s' %(word_schema,'spacing')))
                            node = pPr[0]
                            node.set('%s%s' %(word_schema,'line'),errorlist[listCount]['rightValue'])
                        #print etree.tostring(pPr,pretty_print = True)
                        break
                listCount += 1
            elif errorlist[listCount]['type'] == 'paraFrontSpace':
                for pPr in paragr:
                    if _check_element_is(pPr,'pPr'):
                        found_node = False
                        #print etree.tostring(pPr,pretty_print = True)
                        for node in pPr:
                            if _check_element_is(node,'spacing'):
                                found_node = True
                                node.set('%s%s' %(word_schema,'before'),errorlist[listCount]['rightValue'])
                                break
                        if not found_node:
                            pPr.insert(0,etree.Element('%s%s' %(word_schema,'spacing')))
                            node = pPr[0]
                            node.set('%s%s' %(word_schema,'before'),errorlist[listCount]['rightValue'])
                        #print etree.tostring(pPr,pretty_print = True)
                        break
                listCount +=  1
            elif errorlist[listCount]['type'] == 'paraAfterSpace':
                for pPr in paragr:
                    if _check_element_is(pPr,'pPr'):
                        found_node = False
                        for node in pPr:
                            if _check_element_is(node,'spacing'):
                                found_node = True
                                node.set('%s%s' %(word_schema,'after'),errorlist[listCount]['rightValue'])
                                break
                        if not found_node:
                            pPr.insert(0,etree.Element('%s%s' %(word_schema,'spacing')))
                            node = pPr[0]
                            node.set('%s%s' %(word_schema,'after'),errorlist[listCount]['rightValue'])
                        break
                listCount += 1
            elif errorlist[listCount]['type'] == 'fontCN':
                for r in _iter(paragr,"r"):
                    found_rPr = False
                    for rPr in r.iter(tag=etree.Element):
                        if _check_element_is(rPr,'rPr'):
                            found_rPr = True
                            found_node = False
                            for node in rPr:
                                if _check_element_is(node,'rFonts'):
                                    found_node = True
                                    node.set('%s%s' %(word_schema,'eastAsia'),errorlist[listCount]['rightValue'])
                                    break
                            if found_rPr == True and found_node == False:
                                rPr.insert(0,etree.Element('%s%s' %(word_schema,'rFonts')))
                                node = rPr[0]
                                node.set('%s%s' %(word_schema,'eastAsia'),errorlist[listCount]['rightValue'])
                    if not found_rPr:
                        r.insert(0,etree.Element('%s%s' %(word_schema,'rPr')))
                        rPr = r[0]
                        rPr.insert(0,etree.Element('%s%s' %(word_schema,'rFonts')))
                        node = rPr[0]
                        node.set('%s%s' %(word_schema,'eastAsia'),errorlist[listCount]['rightValue'])
                listCount += 1
            elif errorlist[listCount]['type'] == 'fontEN':
                for r in _iter(paragr,"r"):
                    found_rPr = False
                    for rPr in r.iter(tag=etree.Element):
                        if _check_element_is(rPr,'rPr'):
                            found_rPr = True
                            found_node = False
                            for node in rPr:
                                if _check_element_is(node,'rFonts'):
                                    found_node = True
                                    node.set('%s%s' %(word_schema,'ascii'),errorlist[listCount]['rightValue'])
                                    break
                            if found_rPr == True and found_node == False:
                                rPr.insert(0,etree.Element('%s%s' %(word_schema,'rFonts')))
                                node = rPr[0]
                                node.set('%s%s' %(word_schema,'ascii'),errorlist[listCount]['rightValue'])
                    if not found_rPr:
                        r.insert(0,etree.Element('%s%s' %(word_schema,'rPr')))
                        rPr = r[0]
                        rPr.insert(0,etree.Element('%s%s' %(word_schema,'rFonts')))
                        node = rPr[0]
                        node.set('%s%s' %(word_schema,'ascii'),errorlist[listCount]['rightValue'])
                listCount += 1
            elif errorlist[listCount]['type'] == 'fontTheme':
                for rPr in _iter(paragr, "rPr"):
                    if rPr.getparent().tag != "%s%s" % (word_schema, "r"):
                        continue
                    for rFonts in _iter(rPr, "rFonts"):
                        for theme in ['asciiTheme', 'cstheme', 'eastAsiaTheme', 'hAnsiTheme']:
                            if has_key(rFonts, theme):
                                rFonts.set('%s%s' %(word_schema,theme),"")
                listCount += 1
            elif errorlist[listCount]['type'] == 'fontSize':
                for r in _iter(paragr,"r"):
                    found_rPr = False
                    for rPr in r.iter(tag=etree.Element):
                        if _check_element_is(rPr,'rPr'):
                            found_rPr = True
                            found_node = False
                            for node in rPr:
                                if _check_element_is(node,'sz'):
                                    found_node = True
                                    node.set('%s%s' %(word_schema,'val'),errorlist[listCount]['rightValue'])
                                    break
                            if found_rPr == True and found_node == False:
                                rPr.insert(0,etree.Element('%s%s' %(word_schema,'sz')))
                                node = rPr[0]
                                node.set('%s%s' %(word_schema,'val'),errorlist[listCount]['rightValue'])
                    if not found_rPr:
                        r.insert(0,etree.Element('%s%s' %(word_schema,'rPr')))
                        rPr = r[0]
                        rPr.insert(0,etree.Element('%s%s' %(word_schema,'sz')))
                        node = rPr[0]
                        node.set('%s%s' %(word_schema,'val'),errorlist[listCount]['rightValue'])
                listCount += 1
            elif errorlist[listCount]['type'] == 'fontShape':
                for r in _iter(paragr,"r"):
                    found_rPr = False
                    for rPr in r.iter(tag=etree.Element):
                        if _check_element_is(rPr,'rPr'):
                            found_rPr = True
                            found_node = False
                            for node in rPr:
                                if _check_element_is(node,'b'):
                                    found_node = True
                                    node.set('%s%s' %(word_schema,'val'),errorlist[listCount]['rightValue'])
                                    break
                            if not found_node:
                                rPr.insert(0,etree.Element('%s%s' %(word_schema,'b')))
                                node = rPr[0]
                                node.set('%s%s' %(word_schema,'val'),errorlist[listCount]['rightValue'])
                    if not found_rPr:
                        r.insert(0,etree.Element('%s%s' %(word_schema,'rPr')))
                        rPr = r[0]
                        rPr.insert(0,etree.Element('%s%s' %(word_schema,'b')))
                        node = rPr[0]
                        node.set('%s%s' %(word_schema,'val'),errorlist[listCount]['rightValue'])
                listCount += 1
            elif errorlist[listCount]['type'] == "ref":
                flag = 0
                refnum = 0
                ref_value = {}
                ref_value['ref'] = None
                pat = re.compile(' *REF *')
                for r in _iter(paragr,'r'):
                    fldflag = 0
                    for fldchar in _iter(r,"fldChar"):
                        fldflag = 1
                        if has_key(fldchar,'fldCharType'):
                            if get_val(fldchar,'fldCharType') == 'begin':
                                flag = 1
                            elif get_val(fldchar,'fldCharType') == 'separate':
                                flag = 2
                            elif get_val(fldchar,'fldCharType') == 'end' and ref_value['ref'] != None:
                                flag = 0
                                #ref_value['ref'] == None
                    if flag == 1 and fldflag == 0:
                        for instrText in _iter(r,'instrText'):
                            if pat.match(instrText.text):
                                ref_value['ref'] = pat.sub("",instrText.text)[0:13]
                                refnum += 1
                                #print ref_value['ref']
                    if flag == 2 and refnum == int(errorlist[listCount]['location']) and fldflag == 0:
                        num = 0
                        for text in _iter(r,'t'):
                            if num == 0:
                                text.text = errorlist[listCount]['rightValue'] 
                                num = 1
                            else:
                                text.text = ''
                listCount += 1
            elif errorlist[listCount]['type'] == "refVertAlign":
                flag = 0
                refnum = 0
                ref_value = {}
                ref_value['ref'] = None
                pat = re.compile(' *REF *')
                rlist = []
                for r in _iter(paragr,'r'):
                    fldflag = 0
                    for fldchar in _iter(r,"fldChar"):
                        rlist.append(r)
                        fldflag = 1
                        if has_key(fldchar,'fldCharType'):
                            if get_val(fldchar,'fldCharType') == 'begin':
                                flag = 1
                            elif get_val(fldchar,'fldCharType') == 'separate':
                                flag = 2
                            elif get_val(fldchar,'fldCharType') == 'end' and ref_value['ref'] != None and refnum == int(errorlist[listCount]['location']) :
                                flag = 0
                                for rlistitem in rlist:
                                    vertflag = 0
                                    for vertAlign in _iter(rlistitem,'vertAlign'):
                                        vertAlign.set('%s%s' %(word_schema,'val'),errorlist[listCount]['rightValue'])
                                        vertflag = 1
                                    if vertflag == 0:
                                        #print 1
                                        for rpr in _iter(rlistitem,"rPr"):
                                            rpr.insert(0,etree.Element('%s%s' %(word_schema,'vertAlign')))
                                            vertAlign = rpr[0]
                                            vertAlign.set('%s%s' %(word_schema,'val'),errorlist[listCount]['rightValue'])
                                rlist = []
                                #ref_value['ref'] == None
                    if flag == 1 and fldflag == 0:
                        rlist.append(r)
                        for instrText in _iter(r,'instrText'):
                            if pat.match(instrText.text):
                                ref_value['ref'] = pat.sub("",instrText.text)[0:13]
                                refnum += 1
                                #print ref_value['ref']
                    if flag == 2 and fldflag == 0 and ref_value['ref'] != None: 
                        rlist.append(r)
                listCount += 1
            else:
                listCount += 1

startTime=time.time()  
#主程序__main__入口差不多在这里了
xml_from_file,style_from_file = get_word_xml(Docx_Filename)
xml_tree = get_xml_tree(xml_from_file)
#sys.exit()
errorlist = read_error(Checkout_Filename)
spacelist=[]
spacefile = open(ModifySpace_FileName,'r')
for line in spacefile:
    spacelist.append(int(line))
i = 0
s = ''
modify(xml_tree,errorlist)
endTime=time.time()

#---zqd 新增 20160121-------------------------------------------------
zipF = zipfile.ZipFile(Docx_Filename)
#创建临时目录
tmp_dir = tempfile.mkdtemp()
#把原来的docx解压到临时目录
zipF.extractall(tmp_dir)
#把新的xml写到里面
with open(os.path.join(tmp_dir,'word/document.xml'),'w',encoding="utf-8") as f:
    xmlstr = etree.tostring (xml_tree, pretty_print=False,encoding="UTF-8").decode("utf-8")#此处不识别'gb2312'
    f.write(xmlstr)

# Get a list of all the files in the original docx zipfile
filenames = zipF.namelist()
# Now, create the new zip file and add all the filex into the archive
# zip_copy_filename = 'result.docx'
zip_copy_filename = Data_DirPath + 'result.docx'
with zipfile.ZipFile(zip_copy_filename, "w") as docx:
    for filename in filenames:
        docx.write(os.path.join(tmp_dir,filename),  filename)

# Clean up the temp dir
shutil.rmtree(tmp_dir)
#-----------------------------------------------             
print(endTime - startTime)

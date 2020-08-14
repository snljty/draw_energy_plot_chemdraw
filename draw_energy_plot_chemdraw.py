#! /usr/bin/env python3
# -*- Coding: UTF-8 -*-

r"""
Draw energy plot.
From Excel to ChemDraw.
"""

import pandas as pd

itemid: int = 0
factor: float = 10.0
iname = 'input.xlsx'
oname = 'output.cdxml'
levelLength = 20
connectionLength = 30

data = pd.read_excel(iname)

def PrintBegin() -> None:
    r'''
print the initial part, including xml version, CDXML begin, 
colortable, fonttable, page begin, ...
'''
    global oname
    f = open(oname, 'w')
    print('''<?xml version="1.0" encoding="UTF-8" ?>
<CDXML
 CreationProgram="ChemDraw"
 Name="output.cdxml"
>
    <colortable>
        <color r="1" g="1" b="1"/>
        <color r="0" g="0" b="0"/>
        <color r="1" g="0" b="0"/>
        <color r="1" g="1" b="0"/>
        <color r="0" g="1" b="0"/>
        <color r="0" g="1" b="1"/>
        <color r="0" g="0" b="1"/>
        <color r="1" g="0" b="1"/>
    </colortable>
    <fonttable>
        <font id="0" charset="iso-8859-1" name="Arial"/>
    </fonttable>
    <page>''', file = f)
    f.close()
    return

def PrintEnd() -> None:
    r'''
print the final part, including CDXML end and page end.
'''
    global oname
    f = open(oname, 'a')
    print('''    </page>
</CDXML>


''', file = f)
    f.close()
    return

def PrintLine(posx: float, posy: float) -> None:
    r'''
print a bar line.
'''
    global oname, itemid, levelLength
    itemid += 1
    f = open(oname, 'a')
    print('''        <arrow
         id="%d"
         FillType="None"
         ArrowheadType="Solid"
         Head3D="%.1lf %.1lf 0"
         Tail3D="%.1lf %.1lf 0"
        />''' % (itemid, posx, posy, posx + levelLength, posy), file = f)
    f.close()
    return

def PrintConnect(posx1: float, posy1: float, posx2: float, posy2: float) -> None:
    r'''
print a bar line.
'''
    global oname, itemid
    itemid += 1
    f = open(oname, 'a')
    print('''        <arrow
         id="%d"
         LineType="Dashed"
         FillType="None"
         ArrowheadType="Solid"
         Head3D="%.1lf %.1lf 0"
         Tail3D="%.1lf %.1lf 0"
        />''' % (itemid, posx1, posy1, posx2, posy2), file = f)
    f.close()
    return

def PrintText(posx: float, posy: float, content: str) -> None:
    r'''
print text.
'''
    global oname, itemid
    itemid += 1
    f = open(oname, 'a')
    print('''        <t
         id="%d"
         p="%.1lf %.1lf"
         LineHeight="auto"
        >
            <s font="0" color="0" size="8">%s</s>
        </t>''' % (itemid, posx, posy, content), file = f)
    f.close()
    return

if __name__ == '__main__':
    initPosx = 20
    initPosy = 400
    PrintBegin()
    for i in data.index:
        l = list(data.loc[i])
        newPosx = initPosx
        newPosy = initPosy
        status = False # if any level of this line has been drawn, status is True, otherwise False
        for levelIndex in range(len(l)):
            if not pd.isna(l[levelIndex]):
                # this level should be drawn
                newPosy = initPosy - factor * l[levelIndex]
                if status:
                    # draw a connection line here
                    newPosx += connectionLength
                    PrintConnect(oldPosx, oldPosy, newPosx, newPosy)
                # draw a level line here, as well as the text
                PrintLine(newPosx, newPosy)
                PrintText(newPosx, newPosy + 8, '%.1lf' % l[levelIndex])
                newPosx += levelLength
                # at least one level has already been drawn, set status to True, so
                # before any new level in this line is drawn, a connection has to be drawn.
                status = True
                # update posx and posy
                oldPosx = newPosx
                oldPosy = newPosy
            else:
                newPosx += (levelLength + connectionLength)
    newPosx = initPosx
    for struname in list(data.columns):
        PrintText(newPosx, 600, struname)
        newPosx += 50
    PrintEnd()


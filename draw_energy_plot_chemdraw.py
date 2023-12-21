#! /usr/bin/env python3
# -*- Coding: UTF-8 -*-

r"""
Draw energy plot.
From Excel to ChemDraw.
"""

import pandas as pd

itemid: int = 0
factor: float = 10.
levelLength: float = 20.
connectionLength: float = 30.

def PrintBegin(oname: str) -> None:
    r'''
print the initial part, including xml version, CDXML begin, 
colortable, fonttable, page begin, ...
'''
    with open(oname, 'wt') as f:
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
    return

def PrintEnd(oname: str) -> None:
    r'''
print the final part, including CDXML end and page end.
'''
    with open(oname, 'at') as f:
        print('''    </page>
</CDXML>


''', file = f)
    return

def PrintLine(oname: str, posx: float, posy: float) -> None:
    r'''
print a bar line.
'''
    global itemid, levelLength
    itemid += 1
    with open(oname, 'at') as f:
        print(f'''        <arrow
         id="{itemid:d}"
         FillType="None"
         ArrowheadType="Solid"
         Head3D="{posx:.1f} {posy:.1f} 0"
         Tail3D="{posx + levelLength:.1f} {posy:.1f} 0"
        />''', file = f)
    return

def PrintConnect(oname: str, posx1: float, posy1: float, posx2: float, posy2: float) -> None:
    r'''
print a bar line.
'''
    global itemid
    itemid += 1
    with open(oname, 'at') as f:
        print(f'''        <arrow
         id="{itemid:d}"
         LineType="Dashed"
         FillType="None"
         ArrowheadType="Solid"
         Head3D="{posx1:.1f} {posy1:.1f} 0"
         Tail3D="{posx2:.1f} {posy2:.1f} 0"
        />''', file = f)
    return

def PrintText(oname: str, posx: float, posy: float, content: str) -> None:
    r'''
print text.
'''
    global itemid
    itemid += 1
    with open(oname, 'at') as f:
        print(f'''        <t
         id="{itemid:d}"
         p="{posx:.1f} {posy:.1f}"
         LineHeight="auto"
        >
            <s font="0" color="0" size="8">{content:s}</s>
        </t>''', file = f)
    return

def GenerateFile(iname: str='input.xlsx', oname: str='output.cdxml') -> None:
    # global data
    if not iname:
        iname = 'input.xlsx'
    if not oname:
        oname = 'output.cdxml'
    if not iname.endswith('.xlsx'):
        iname += '.xlsx'
    if not oname.endswith('.cdxml'):
        oname += '.cdxml'
    initPosx = 20
    initPosy = 400
    data = pd.read_excel(iname)
    PrintBegin(oname)
    for i in data.index:
        ene = list(data.loc[i])
        newPosx = initPosx
        newPosy = initPosy
        status = False # if any level of this line has been drawn, status is True, otherwise False
        for levelIndex in range(len(ene)):
            if not pd.isna(ene[levelIndex]):
                # this level should be drawn
                newPosy = initPosy - factor * ene[levelIndex]
                if status:
                    # draw a connection line here
                    newPosx += connectionLength
                    PrintConnect(oname, oldPosx, oldPosy, newPosx, newPosy)
                # draw a level line here, as well as the text
                PrintLine(oname, newPosx, newPosy)
                PrintText(oname, newPosx, newPosy + 8, f'{ene[levelIndex]:.2f}')
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
        if isinstance(struname, (float, int)):
            label = f'{float(struname):.2f}'
        else:
            label = str(struname)
        PrintText(oname, newPosx, 600, str(label))
        newPosx += 50
    PrintEnd(oname)

if __name__ == '__main__':
    import sys
    if len(sys.argv) - 1 >= 1:
        iname = sys.argv[1]
    else:
        iname = ""
    if len(sys.argv) - 1 >= 2:
        oname = sys.argv[2]
    else:
        oname = ""
    GenerateFile(iname, oname)


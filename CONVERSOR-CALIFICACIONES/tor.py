"""
This file is part of ToR Conversor.

    ToR Conversor is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    ToR Conversor is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with ToR Conversor.  If not, see <https://www.gnu.org/licenses/>.
"""

import numpy as np
import re

americanScores = [ 'F-', 'F+', 'E-', 'E+', 'D-', 'D+',
                   'C-', 'C+', 'B-',  'B+', 'A-', 'A+',
                   'F',  'E,',  'D',  'C',  'B', 'A' ]
americanTranslation = [ 0,     2,   3,     5,    6,    8,
                        9,    11,   12,    14,   15,  17,
                        1,     4,    7,    10,   13,  16 ]


def isNumber(text):
    """ Devuelve True si el texto es un número convertible """
    try:
        i = float(text)
    except ValueError as verr:
        return False
    except Exception as ex:
        return None
    return True

def extractRangesAliases(scorelist):
    # detect subranges (delimited by ;)
    list2 = []
    maxDelim = 0
    for cell in scorelist:
        if type(cell) == float or type(cell) == int:
            cell = str(cell)
        maxDelim = max(cell.count(';'), maxDelim)
        cell = cell.split(';')
        list2.append(cell)
    # detect ranges (delimited by -) and aliases
    out = []
    alias = []
    for alternative in list2:
        out.append([])
        alias.append([])
        for text in alternative:
            text = str(text)
            txt = re.findall(r"\([^\)]+\)",text)
            if len(txt) != 0:
                text = text.replace(txt[0], "")
                aliasText = txt[0][1:-1]
            else:
                aliasText = ""
            text = text.split('-')
            value = []
            for t in text:
                value.append(float(t))
            if len(value) == 1:
                value.append(value[0])
            out[-1].append(value)
            alias[-1].append(aliasText)
    return [out, alias]

def expandScores(origin, destination, american=False):
    """ Divide the intervals of a row of the score table"""
    src = destination[:] # original scores should be in aliases (TODO)
    i = 0
    for d in destination:
        for a in americanScores:
            newScore = str(americanTranslation[americanScores.index(a)])
            destination[i] = destination[i].replace(a,  newScore)
        i += 1
    maxDelim = 0
    # detect ;

    y, aliasy = extractRangesAliases(destination)
    x, aliasx = extractRangesAliases(origin)
    index = 0
    for vy in y:
        if vy[0][0] == vy[0][1] and index < len(y)-1:
            vy[0][1] = y[index+1][0][0]
            y[index] = vy
        index += 1
    return [x, aliasx, y, aliasy]

def simpleScore(minx, maxx, miny, maxy, score):
    sx = abs(maxx - minx)
    sy = abs(maxy - miny)
    if sx == 0 or sy == 0:
        return maxx
    nscore = (score - min(miny, maxy)) / sy
    if miny < maxy:
        return nscore * sx + minx
    else:
        return (1 - nscore) * sx + minx

def score(x, aliasx, y, aliasy, score, american):
    """ Return the transformed score, unless the student has no score in the subject """
    if american and not isNumber(score):
        score = americanTranslation[americanScores.index(score)]
    for j in range(len(x)-1,-1,-1): # reverse order, if the range is shared, take the upper bound
        ax = aliasx[j]
        vx = x[j]
        ay = aliasy[j]
        vy = y[j]
        if isNumber(score):
            score = float(score)
            for r in vy: # this part only works if source doesn't have subranges
                if (score <= r[0] and score >= r[1]) or (score >= r[0] and score <= r[1]):
                    out = simpleScore(vx[0][0], vx[-1][-1],
                                      vy[0][0], vy[-1][-1], score)
                    return [out, vx, vy, ay]
        else:
            if score in ay:
                i = ay.index(score)
                score = vy[i]
                score = (score[0]+score[1])/2
                out = simpleScore(vx[0][0], vx[-1][-1],
                                  vy[0][0], vy[-1][-1], score)
                return [out, vx, vy, ay]
    print("SCORE OUT OF RANGE: ", score)
    print(y)
    print(x)
    #exit(1)
    return [-1, ax, vx, ay, vy]

def parseToR(data):
    american = False
    output = {}
    for d in data:
        #print(d[0:5])
        nameDST, creditsDST, nothing, blockDST, scoreDST = d[0:5]
        nameORIG, creditsORIG, nothing, blockORIG = d[5:]
        blockDST = str(blockDST).strip()
        blockORIG = str(blockORIG).strip()
        nameDST = str(nameDST).strip()
        nameORIG = str(nameORIG).strip()
        scoreDST = str(scoreDST).replace(",", ".").strip()
        try:
            scoreDST = float(scoreDST)
        except:
            if scoreDST in americanScores:
                american = True
            else:
                scoreDST = 0
        try:
            creditsDST = int(creditsDST)
        except:
            creditsDST = 0
        try:
            creditsORIG = int(creditsORIG)
        except:
            creditsORIG = 0
        if blockDST != "" and blockDST not in output:
            output[blockDST] = [[], []]
        if blockORIG != "" and blockORIG not in output:
            print("Warning, {} has not been created previously.".format(blockORIG))
            output[blockORIG] = [[], []]

        if blockDST != "" and nameDST.strip() != "":
            output[blockDST][0].append([nameDST, creditsDST, scoreDST])
        if blockORIG != "" and nameORIG.strip() != "":
            output[blockORIG][1].append([nameORIG, creditsORIG])

    return (american, output)

def extendToR(ToR, x, aliasx, y, aliasy, american):
    idv = 1
    output = {}
    for data in ToR:
        output[idv] = [[],[]]
        for d in ToR[data][0]:
            subject = d[0]
            credits = d[1]
            scdata = score(x, aliasx, y, aliasy, d[2], american)
            d.append(scdata[0])
            d.extend(scdata[1][0])
            d.extend(scdata[2][0])
            d.append(scdata[3][0])
            output[idv][0].append(d)
            # print(" {}".format(d))
        output[idv][1] = ToR[data][1]
        idv += 1
    return output


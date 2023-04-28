import pandas as pd
import openpyxl
from array import *

wb = openpyxl.load_workbook('iddaa.xlsx')
ex = wb.active
arr = [];
for i in range(1, 24):
  cell_obj = ex.cell(row = 17, column = i) #burdaki rowu exceldeki girilen rowla ayni degistir
  arr.append(cell_obj.value)
  
print("mac:", arr[22])

#sadece ic saha ve deplasman maclarini ayri baz alarak, mesela deplasman takiminin sadece deplasmandaki son 5 mac puan ortalamasi vs

## puan ortalamalari
icpuan = arr[0]
deppuan = arr[1]
son5pic = arr[2]
son5pdep = arr[3]
icrbport = arr[4] #puan durumunda benzer rakiplere karsi puan ort
deprbport = arr[5]

## gol ortalamalari
icgol = arr[6]
icyen = arr[7]
depyen = arr[8]
depgol = arr[9]
son5adep = arr[10]
son5ydep = arr[11]
son5aic = arr[12]
son5yic = arr[13]

## xg
son5icxg = arr[14]
son5depxg = arr[15]
son5icverxg = arr[16]
son5depverxg = arr[17]

## gol atilmayan / yenmeyen maclar
csic = arr[18]
golsuzic = arr[19]
csdep = arr[20]
golsuzdep = arr[21]

def gol(son5aic, son5yic, son5adep, son5ydep, son5icxg, son5icverxg, son5depxg, son5depverxg, csic, csdep, golsuzic, golsuzdep, icgol, depgol):
  #toplam, son5 ve xg ortalamasi (xg sacma skorlarin ve sifirlarin etkisini azaltmak icin)
  icyenen = (icyen + son5yic + son5icverxg)/3
  depyenen = (depyen + son5ydep + son5depverxg)/3
  icgol = (son5aic + icgol + son5icxg)/3
  depgol = (son5adep + depgol + son5depxg)/3
  #2 takimi birbiriyle karsilastir
  ic = (icgol + depyenen)/2
  dep = (depgol + icyenen)/2
  #mac basi golsuz mac ortalamasinin 0.3u takimlarin gol skorundan eksiliyor
  csbonusic = 0.3*csic
  tersic = 0.3*golsuzic
  csbonusdep = 0.3*csdep
  tersdep = 0.3* golsuzdep
  ic = ic - csbonusdep - tersic
  dep = dep - csbonusic - tersdep
  return ic, dep

[ic, dep] = gol(son5aic, son5yic, son5adep, son5ydep, son5icxg, son5icverxg, son5depxg, son5depverxg, csic, csdep, golsuzic, golsuzdep,icgol, depgol)
##print("ic gol skoru: ",ic)
##print("dep gol skoru:", dep)

def ms(ic, dep, icpuan, son5pic, icrbport, deppuan, son5pdep, deprbport):
  #takimlarin 3 farkli puan ortalamasinin ortalamasi (genel ve son5 2, benzer rakiplere karsi 1 sayiliyor 1-2 macin etkisini dusurmek icin)
    icort = (icpuan + icpuan + son5pic + son5pic + icrbport)/5
    deport =(deppuan + deppuan + son5pdep + son5pdep + deprbport)/5
  #gol skoruyla puan skoru ortalamasi
    evskor = (ic + icort)/2
    depskor = (dep + deport)/2
    return evskor, depskor

[evskor, depskor] = ms(ic, dep, icpuan, son5pic, icrbport, deppuan, son5pdep, deprbport)
##print("\nevskor: ",evskor)
##print("depskor:",depskor)

#toplam skora gore oran
toplam = evskor+depskor
ms1 = 1/(evskor/toplam)*1.05
ms2 = 1/(depskor/toplam)*1.05

#takimlardan biri net favoriyse onun orani daha dusup digerininki daha da yukselt
if ms1 > (ms2+1):
  ms1 = ms1/0.70
  ms2 = ms2*0.90
if (ms1+1) < ms2:
    ms1 = ms1*0.90
    ms2 = ms2/0.70
  #takimlarin oranlar yakinsa ikisi de yukselt
if abs(ms1-ms2) < 1 : 
    ms1 = ms1*1.1
    ms2 = ms2*1.1
  #oranlar cok yakinsa daha da yukselt
if abs(ms1-ms2) < 0.4 :
    ms1 = ms1*1.1
    ms2 = ms2*1.1
print("\nms1: ", "{0:.2f}".format(ms1))
print("ms2: ","{0:.2f}".format(ms2))

#burasi toplam gol, alt-ust
tg = (ic + dep)*1.2 
e = 2.71828
def pois(tg, e):
  alt1 = (pow(tg,2)/2)*pow(e,-tg) + (pow(tg,1)/1)*pow(e,-tg) + pow(e,-tg)
  ust1 = (pow(tg,3)/6)*pow(e,-tg) + (pow(tg,4)/24)*pow(e,-tg) + (pow(tg,5)/125)*pow(e,-tg) + (pow(tg,6)/750)*pow(e,-tg) + (pow(tg,7)/5250)*pow(e,-tg)
  return alt1, ust1
[alt1, ust1] = pois(tg, e)

'''
alt oranlari gercege gore dusuk ve ust taraftaki 
golsuz mac mevzusu da dusuruyor, oyuzden daha arttiriliyor.
1in altina dusmesin diye 1den basliyor, toplam oran 3.3 civari olucak sekilde buyutuluyor (gercekte hep 3.3 civari)
'''
alt = 1 + ust1 * 1.4 #bunlar ters cunku "olmama olasiligi" seklinde hesap
ust = 1 + alt1 * 1.4 

if abs(ust-alt) < 0.1: #oranlar yakinsa yukselt
  alt = alt * 1.03
  ust = ust * 1.03
  
print("\n2.5 alt: ", "{0:.2f}".format(alt))
print("2.5 ust: ", "{0:.2f}".format(ust))   
 

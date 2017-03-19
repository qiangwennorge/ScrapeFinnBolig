# -*- coding: utf-8 -*-
# encoding=utf8 

'''
Created on Feb 21, 2017

@author: wenqiang
'''
#import requests
from bs4 import BeautifulSoup
#import lxml
import urllib2
import re
import xlsxwriter
import sys  

reload(sys)  
sys.setdefaultencoding('utf8')

FinnBoligOutputFile = xlsxwriter.Workbook('FinnBoligResult.xlsx')
worksheet = FinnBoligOutputFile.add_worksheet()

worksheet.write(0,0,'Finn Code')
worksheet.write(0,1,'Bolig Address')
worksheet.write(0,2,'Prisantydning')
worksheet.write(0,3,'Fellesgjeld')
worksheet.write(0,4,'Fellesformue')
worksheet.write(0,5,'Felleskost/mnd.')
worksheet.write(0,6,'Totalpris')
worksheet.write(0,7,'Ligningsverdi')
worksheet.write(0,8,'Verditakst')
worksheet.write(0,9,'Kommunale avg.')
worksheet.write(0,10,'Lånetakst')
worksheet.write(0,11,'Primærrom')
worksheet.write(0,12,'Bruksareal')
worksheet.write(0,13,'Bruttoareal')
worksheet.write(0,14,'Soverom')
worksheet.write(0,15,'Boligtype')
worksheet.write(0,16,'Eieform')
worksheet.write(0,17,'Tomteareal')
worksheet.write(0,18,'Byggeår')
worksheet.write(0,19,'Etasje')

BoligStartNum = 0

def GetLinkOfEachRealestate(LinkList, soup):
    for ParentalTag in soup.find_all('div','unit flex align-items-stretch result-item'):
        OneLinkOfEachRealestate = ParentalTag.contents[1].get('href')
        #print OneLinkOfEachRealestate
        LinkList.append(OneLinkOfEachRealestate)
    return LinkList

def GetBoligInfo(LinkNum, BoligUrl):
    try:
        BoligResponse = urllib2.urlopen(BoligUrl)
        BoligHtml = BoligResponse.read()

        BoligSoup = BeautifulSoup(BoligHtml,"html.parser")

        BoligObject = BoligSoup.find_all('dl', 'r-prl mhn multicol col-count1upto640 col-count2upto768 col-count1upto990 col-count2from990')
        #print BoligObject[0].prettify()
        
        try:
            Address = BoligObject[0].parent.p.get_text()
            worksheet.write(LinkNum,1,Address)
            print "Address = " + Address
        except Exception:
            pass
        
        try:
            Prisantydning = BoligObject[0].parent.dl.dd.get_text().strip()
            worksheet.write(LinkNum,2,Prisantydning)
            print "Prisantydning = " + Prisantydning
        except Exception:
            pass
        
        try:
            Prisfra = BoligObject[0].parent.find('div', 'h2 mbn r-margin').find_next_sibling('div','h1 mtn r-margin').get_text().strip()
            Pristil = BoligObject[0].parent.find('div', 'h1 mtn r-margin').find_next_sibling('div','h1 mtn r-margin').get_text().strip()
            Prisantydning = Prisfra + " <-> " + Pristil
            print Prisantydning
            worksheet.write(LinkNum,2,Prisantydning)
        except Exception:
            pass
        
        try:
            Pris = BoligObject[0].parent.find('div', 'h2 mbn r-margin').find_next_sibling('div','h1 mtn r-margin').get_text().strip()
            worksheet.write(LinkNum,2,Pris)
            print Pris
        except Exception:
            pass    
                
        for dlitem in BoligObject:  
            
            try:
                Fellesgjeld = dlitem.find('dt', text='Fellesgjeld').find_next_sibling('dd').text.strip()
                worksheet.write(LinkNum,3,Fellesgjeld)
                print "Fellesgjeld = " + Fellesgjeld
            except Exception:
                pass
            
            try:
                Fellesformue = dlitem.find('dt', text='Fellesformue').find_next_sibling('dd').text.strip()
                worksheet.write(LinkNum,4,Fellesformue)
                print "Fellesformue = " + Fellesformue
            except Exception:
                pass
            
            try:
                Felleskost = dlitem.find('dt', text='Felleskost/mnd.').find_next_sibling('dd').text.strip()
                worksheet.write(LinkNum,5,Felleskost)
                print "Felleskost/mnd. = " + Felleskost
            except Exception:
                pass
            
            try:
                Totalpris = dlitem.find('dt', text='Totalpris').find_next_sibling('dd').text.strip()
                worksheet.write(LinkNum,6,Totalpris)
                print "Totalpris = " + Totalpris
            except Exception:
                pass
            
            try:
                Verditakst = dlitem.find('dt', text='Verditakst').find_next_sibling('dd').text.strip()
                worksheet.write(LinkNum,7,Verditakst)
                print "Verditakst = " + Verditakst
            except Exception:
                pass
                       
            try:
                Ligningsverdi = dlitem.find('dt', text='Ligningsverdi').find_next_sibling('dd').text.strip()
                worksheet.write(LinkNum,8,Ligningsverdi)
                print "Ligningsverdi = " + Ligningsverdi
            except Exception:
                pass
     
            try:
                KommunaleAvg = dlitem.find('dt', text='Kommunale avg.').find_next_sibling('dd').text.strip()
                worksheet.write(LinkNum,9,KommunaleAvg)
                print "Kommunale avg. = " + KommunaleAvg 
            except Exception:
                pass

            try:
                Lanetakst = dlitem.find('dt', text='Lånetakst').find_next_sibling('dd').text.strip()
                worksheet.write(LinkNum,10,Lanetakst)
                print "Lanetakst = " + Lanetakst
            except Exception:
                pass
                            
            try:
                Primarrom = dlitem.find('dt', text='Primærrom').find_next_sibling('dd').text.strip()
                worksheet.write(LinkNum,11,Primarrom)
                print "Primarrom = " + Primarrom
            except Exception:
                pass
      
            try:
                Bruksareal = dlitem.find('dt', text='Bruksareal').find_next_sibling('dd').text.strip()
                worksheet.write(LinkNum,12,Bruksareal)
                print "Bruksareal = " + Bruksareal 
            except Exception:
                pass   
            
            try:
                Bruttoareal = dlitem.find('dt', text='Bruttoareal').find_next_sibling('dd').text.strip()
                worksheet.write(LinkNum,13,Bruttoareal)
                print "Bruttoareal = " + Bruttoareal
            except Exception:
                pass

            try:
                Soverom = dlitem.find('dt', text='Soverom').find_next_sibling('dd').text.strip()
                worksheet.write(LinkNum,14,Soverom)
                print "Soverom = " + Soverom  
            except Exception:
                pass

            try:
                Boligtype = dlitem.find('dt', text='Boligtype').find_next_sibling('dd').text.strip()
                worksheet.write(LinkNum,15,Boligtype)
                print "Boligtype = " + Boligtype  
            except Exception:
                pass

            try:
                Eieform = dlitem.find('dt', text='Eieform').find_next_sibling('dd').text.strip()
                worksheet.write(LinkNum,16,Eieform)
                print "Eieform = " + Eieform
            except Exception:
                pass
 
            try:
                Tomteareal = dlitem.find('dt', text='Tomteareal').find_next_sibling('dd').text.strip()
                Tomteareal = re.sub(r"\n +", " ", Tomteareal)
                worksheet.write(LinkNum,17,Tomteareal)
                print "Tomteareal = " + Tomteareal
            except Exception:
                pass
                            
            try:
                Byggear = dlitem.find('dt', text='Byggeår').find_next_sibling('dd').text.strip()
                worksheet.write(LinkNum,18,Byggear)
                print "Byggear = " + Byggear
            except Exception:
                pass
                                
            try:
                Eierskifteforsikring = dlitem.find('dt', text='Eierskifteforsikring').find_next_sibling('dd').text.strip()
                print "Eierskifteforsikring = " + Eierskifteforsikring
            except Exception:
                pass
            
            try:
                Etasje = dlitem.find('dt', text='Etasje').find_next_sibling('dd').text.strip()
                worksheet.write(LinkNum,19,Etasje)
                print "Etasje = " + Etasje
            except Exception:
                pass
            
    except Exception:
        pass
        #return BoligDetailedInfo


InitialUrl = 'https://m.finn.no/realestate/homes/search.html?filters='
InitialResponse = urllib2.urlopen(InitialUrl)
InitialHtml = InitialResponse.read()

InitialSoup = BeautifulSoup(InitialHtml,"html.parser")

InitialTagsForNextPage = InitialSoup.find_all('div','t4 centerify r-margin')[0]
SecondPage = InitialTagsForNextPage.find_all('a',class_='pam')[0].get('href')

NextPageNum = int(re.findall(r'page=(\d+)', SecondPage)[0])

LinkList = []

SecondLinkList = GetLinkOfEachRealestate(LinkList, InitialSoup)

for PageNum in range(NextPageNum,278):
    NextPage = re.sub(r'page=(\d+)', 'page='+str(PageNum), SecondPage)
    NextUrl = 'https://m.finn.no' + NextPage
    NextResponse = urllib2.urlopen(NextUrl)
    NextHtml = NextResponse.read()
    print PageNum
    NextSoup = BeautifulSoup(NextHtml,"html.parser")    
    NextLinkList = GetLinkOfEachRealestate(SecondLinkList, NextSoup)

for BoligLink in NextLinkList:
    BoligStartNum = BoligStartNum + 1
    FinnCode = re.findall(r'finnkode=(\d+)', BoligLink)[0]
    print FinnCode
    worksheet.write(BoligStartNum,0,FinnCode)
    EachBoligUrl = 'https://m.finn.no' + BoligLink
    BoligDetailedInfo = GetBoligInfo(BoligStartNum, EachBoligUrl)

#EachBoligUrl = 'https://m.finn.no/realestate/newbuildings/ad.html?finnkode=91702810&q=91702810'
#GetBoligInfo(1, EachBoligUrl)       
    
FinnBoligOutputFile.close()

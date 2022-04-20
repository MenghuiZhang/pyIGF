# coding: utf8
import sys
sys.path.append(r'R:\pyRevit\xx_Skripte\libs\IGF_libs')
from IGF_log import getlog
from rpw import revit, DB
from pyrevit import script, forms
from IGF_lib import get_value

__title__ = "3.00 Überströmung aktualisieren"
__doc__ = """
Eine Baugruppe besteht aus zwei Auslässe und ein Luftkanal.

Baugruppen: schreibt Raumnummer von Überströmungsfamilie in entsprechende Baugruppe.

MEP-Räume: schreibt Volumen von Baugruppen in MEP-Räume

Überströmungsfamilie: schreibt Volumen von Baugruppen in Überströmungsfamilie
-------------------------

imput Parameter:
-------------------------
Baugruppen: 
IGF_RLT_Überströmung: Überstrommengen
-------------------------
output Parameter:
*************************
Luftdurchlässe:
IGF_RLT_Überströmung: Überstrommengen

IGF_RLT_ÜberströmungAusRaum: Einbauort von Eingangsauslass

IGF_RLT_ÜberströmungInRaum: Einbauort von Ausgangsauslass

*************************
Baugruppen: 
IGF_RLT_Überströmung_Eingang: Eingangsraum

IGF_RLT_Überströmung_Ausgang: Ausganssraum

*************************
Luftkanäle:
IGF_RLT_Überströmung_Eingang: Eingangsraum

IGF_RLT_Überströmung_Ausgang: Ausganssraum

IGF_RLT_Überströmung: Überstrommengen

*************************
MEP-Räume
IGF_RLT_ÜberströmungSummeIn: Überstromluft einströmend in m3/h

IGF_RLT_ÜberströmungSummeAus: Überstromluft ausströmend in m3/h

IGF_RLT_ÜberströmungRaum: Überstrommengen

IGF_RLT_ÜberstromVerteilung: Verteilung von Überstromluft

-------------------------

[2021.11.22]
Version: 2.0
"""
__author__ = "Menghui Zhang"

logger = script.get_logger()
output = script.get_output()

uidoc = revit.uidoc
doc = revit.doc
name = doc.ProjectInformation.Name
number = doc.ProjectInformation.Number
phase = doc.Phases[0]

try:
    getlog(__title__)
except:
    pass

# Baugruppen
Baugruppen_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Assemblies).WhereElementIsNotElementType()
Baugruppen = Baugruppen_collector.ToElementIds()
Baugruppen_collector.Dispose()

class Baugruppe:
    def __init__(self,elemid):
        self.elemid = elemid
        self.elem = doc.GetElement(self.elemid)
        self.members = self.get_members()
        if self.Pruefen():
            self.Volumen = get_value(self.elem.LookupParameter('IGF_RLT_Überströmung'))
            self.Eingang = ''
            self.Ausgang = ''

            self.Eingangsraum = ''
            self.Ausgangsraum = ''

            self.Eingangauslass = ''
            self.Ausgangauslass = ''

            

            self.Leitung = ''
            self.Analyse()
            self.Leitungermitteln()
            self.EingangauslassforSchreiben = Durchlass(self.Eingangauslass.elem,self.Volumen,self.Eingangsraum,self.Ausgangsraum)
            self.AusgangauslassforSchreiben = Durchlass(self.Ausgangauslass.elem,self.Volumen,self.Eingangsraum,self.Ausgangsraum)
    
    def get_members(self):
        return self.elem.GetMemberIds()
    
    def Pruefen(self):
        if len(self.members) == 3:
            return True
        else:
            return False
    
    def Analyse(self):
        for elemid in self.members:
            elem = doc.GetElement(elemid)
            Category = elem.Category.Id.ToString()
            if Category == '-2008013':
                auslass = Auslass(elem,self.Volumen)
                if auslass.typ == 'Out':
                    self.Eingangauslass = auslass
                    self.Eingangsraum = auslass.Einbauort
                    self.Eingang = auslass.Raumid
                elif auslass.typ == 'In':
                    self.Ausgangauslass = auslass
                    self.Ausgangsraum = auslass.Einbauort
                    self.Ausgang = auslass.Raumid

    def Leitungermitteln(self):
        for elemid in self.members:
            elem = doc.GetElement(elemid)
            Category = elem.Category.Id.ToString()
            if Category == '-2008000':
                self.Leitung = Luftkanal(elem,self.Volumen,self.Eingangsraum,self.Ausgangsraum)
            
           

    def werte_schreiben(self):
        def wert_schreiben(param_name,wert):
            param = self.elem.LookupParameter(param_name)
            if param and wert:         
                param.Set(wert)
        wert_schreiben('IGF_RLT_Überströmung_Eingang',self.Eingangsraum)
        wert_schreiben('IGF_RLT_Überströmung_Ausgang',self.Ausgangsraum) 
    def table_row(self):
        return [
            self.Volumen,
            self.Eingangsraum,
            self.Ausgangsraum,
        ]

        

class Luftkanal:
    def __init__(self,elem,vol,eingang,ausgang):
        self.elem = elem
        self.vol = vol
        self.eingang = eingang
        self.ausgang = ausgang
    
    def werte_schreiben(self):
        def wert_schreiben(param_name,wert):
            param = self.elem.LookupParameter(param_name)
            if param and wert:         
                if param.StorageType.ToString() == 'Double':
                    param.SetValueString(str(wert))
                else:
                    param.Set(wert)

        wert_schreiben('IGF_RLT_Überströmung_Eingang',self.eingang)
        wert_schreiben('IGF_RLT_Überströmung_Ausgang',self.ausgang)  
        wert_schreiben('IGF_RLT_Überströmung',self.vol)

class Durchlass:
    def __init__(self,elem,vol,eingang,ausgang):
        self.elem = elem
        self.vol = vol
        self.eingang = eingang
        self.ausgang = ausgang
    
    def werte_schreiben(self):
        def wert_schreiben(param_name,wert):
            param = self.elem.LookupParameter(param_name)
            if param and wert:         
                if param.StorageType.ToString() == 'Double':
                    param.SetValueString(str(wert))
                else:
                    param.Set(wert)

        wert_schreiben('IGF_RLT_ÜberströmungInRaum',self.ausgang)
        wert_schreiben('IGF_RLT_ÜberströmungAusRaum',self.eingang)  
        wert_schreiben('IGF_RLT_Überströmung',self.vol) 


class Auslass:
    def __init__(self,elem,Vol):
        self.elem = elem
        self.vol = Vol
        self.conns = self.get_connector()
        self.typ = self.get_typ()
        self.Einbauort = self.get_einbauort()
        self.Raumid = self.get_raum()

    def get_connector(self):
        return list(self.elem.MEPModel.ConnectorManager.Connectors)

    def get_typ(self):
        conn = self.conns[0]
        if conn.Direction.ToString() == 'Out':
            typ = 'Out'
        elif conn.Direction.ToString() == 'In':
            typ = 'In'

        return typ

    def get_raum(self):
        try:
            return self.elem.Space[phase].Id.ToString()
        except:
            return ''
    
    def get_einbauort(self):
        try:
            return self.elem.Space[phase].Number
        except:
            logger.error('kein Raum ist mit Luftauslass {} verknüpft.'.format(self.elem.Id.ToString()))
            return ''

class MEPRaum:
    def __init__(self,elemid,inListe,ausListe):
        self.elem = doc.GetElement(DB.ElementId(int(elemid)))
        self.summein = 0
        self.summeaus = 0
        self.summe = 0
        self.verteilin = '[m3/h] Aus: '
        self.verteilaus = '[m3/h] In: '
        self.verteiltext = ''
        self.Auslassin = inListe
        self.Auslassaus = ausListe
        self.Analyse()

    
    def Analyse(self):
        SumIn = {}
        SumAus = {}
        for aus_in in self.Auslassin:
            if aus_in.Einbauort in SumIn.keys():
                SumIn[aus_in.Einbauort] += aus_in.vol
            else:
                SumIn[aus_in.Einbauort] = aus_in.vol
            self.summein += aus_in.vol
        for aus_aus in self.Auslassaus:
            if aus_aus.Einbauort in SumAus.keys():
                SumAus[aus_aus.Einbauort] += aus_aus.vol
            else:
                SumAus[aus_aus.Einbauort] = aus_aus.vol
            self.summeaus += aus_aus.vol

        sumin = SumIn.keys()[:]
        sumin.sort()
        sumaus = SumAus.keys()[:]
        sumaus.sort()


        if len(sumin) > 0:
            for aus_in in sumin:
                self.verteilin += aus_in + ': ' + str(int(round(SumIn[aus_in]))) + ', '
        if len(sumaus) > 0:
            for aus_aus in sumaus:
                self.verteilaus += aus_aus + ': ' + str(int(round(SumAus[aus_aus]))) + ', '
        self.summe = self.summein - self.summeaus

        if self.summein > 1:
            self.verteiltext += self.verteilin
        if self.summeaus > 1:
            if self.verteiltext:
                self.verteiltext += self.verteilaus[6:]
            else:
                self.verteiltext += self.verteilaus
        self.verteiltext = self.verteiltext[:-2]
                

    def werte_schreiben(self):
        def wert_schreiben(param_name,wert):
            param = self.elem.LookupParameter(param_name)
            if param:     
                if param.StorageType.ToString() == 'Double':
                    param.SetValueString(str(wert))
                    
                else:
                    param.Set(wert)
        
        wert_schreiben('IGF_RLT_ÜberströmungSummeIn',round(self.summein))
        wert_schreiben('IGF_RLT_ÜberströmungSummeAus',round(self.summeaus))
        wert_schreiben('IGF_RLT_ÜberströmungRaum',round(self.summe))
        wert_schreiben('IGF_RLT_ÜberstromVerteilung',self.verteiltext)



baugruppe_liste = []
luftkanal_liste = []
mepraum_in_dict = {}
mepraum_aus_dict = {}
luftauslass_liste = []
with forms.ProgressBar(title="{value}/{max_value} Baugruppen",cancellable=True, step=10) as pb0:
    for n, bg_id in enumerate(Baugruppen):
        if pb0.cancelled:
            script.exit()

        pb0.update_progress(n + 1, len(Baugruppen))
        baugruppe = Baugruppe(bg_id)

        if baugruppe.Pruefen():
            baugruppe_liste.append(baugruppe)
            luftauslass_liste.append(baugruppe.EingangauslassforSchreiben)
            luftauslass_liste.append(baugruppe.AusgangauslassforSchreiben)
            if baugruppe.Ausgang not in mepraum_in_dict.keys():
                mepraum_in_dict[baugruppe.Ausgang] = [baugruppe.Eingangauslass]
            else:
                mepraum_in_dict[baugruppe.Ausgang].append(baugruppe.Eingangauslass)

            if baugruppe.Eingang not in mepraum_aus_dict.keys():
                mepraum_aus_dict[baugruppe.Eingang] = [baugruppe.Ausgangauslass]
            else:
                mepraum_aus_dict[baugruppe.Eingang].append(baugruppe.Ausgangauslass)

            luftkanal_liste.append(baugruppe.Leitung)


if forms.alert('Raumnummer in Baugruppen schreiben?', ok=False, yes=True, no=True):
    with forms.ProgressBar(title="{value}/{max_value} Baugruppen",cancellable=True, step=10) as pb1:
        t = DB.Transaction(doc,'Infos in Baugruppen schreiben')
        t.Start()
        for n,bg in enumerate(baugruppe_liste):
            if pb1.cancelled:
                script.exit()
            pb1.update_progress(n + 1, len(baugruppe_liste))

            bg.werte_schreiben()
        t.Commit()
        t.Dispose()


if forms.alert('Informationen in Luftkanäle schreiben?', ok=False, yes=True, no=True):
    with forms.ProgressBar(title="{value}/{max_value} Luftkanäle",cancellable=True, step=10) as pb2:
        t = DB.Transaction(doc,'Infos in Luftkanäle schreiben')
        t.Start()
        for n,lk in enumerate(luftkanal_liste):
            if pb2.cancelled:
                script.exit()
            pb2.update_progress(n + 1, len(luftkanal_liste))
            
            lk.werte_schreiben()
        t.Commit()
        t.Dispose()

if forms.alert('Informationen in Luftauslässe schreiben?', ok=False, yes=True, no=True):
    with forms.ProgressBar(title="{value}/{max_value} Luftauslässe",cancellable=True, step=10) as pb3:
        t = DB.Transaction(doc,'Infos in Luftauslässe schreiben')
        t.Start()
        for n,la in enumerate(luftauslass_liste):
            if pb3.cancelled:
                script.exit()
            pb3.update_progress(n + 1, len(luftauslass_liste))
            
            la.werte_schreiben()
        t.Commit()
        t.Dispose()

mepraum_liste = []
for el in mepraum_aus_dict.keys():
    if not el in mepraum_liste:
        mepraum_liste.append(el)
for el in mepraum_in_dict.keys():
    if not el in mepraum_liste:
        mepraum_liste.append(el)

MEP_Liste = []

with forms.ProgressBar(title="{value}/{max_value} MEP Räume",cancellable=True, step=10) as pb4:
    for n,elemid in enumerate(mepraum_liste):
        if pb4.cancelled:
            script.exit()
        pb4.update_progress(n + 1, len(mepraum_liste))

        if elemid in mepraum_aus_dict.keys():
            ausraum = mepraum_aus_dict[elemid]
        else:
            ausraum = []
        if elemid in mepraum_in_dict.keys():
            inraum = mepraum_in_dict[elemid]
        else:
            inraum = []
        
        mepraum = MEPRaum(elemid,inraum,ausraum)
        MEP_Liste.append(mepraum)

if forms.alert('Informationen in MEP Räume schreiben?', ok=False, yes=True, no=True):
    with forms.ProgressBar(title="{value}/{max_value} MEP Räume",cancellable=True, step=10) as pb5:
        t = DB.Transaction(doc,'Infos in MEP Räume schreiben')
        t.Start()
        for n,mepraum in enumerate(MEP_Liste):
            if pb5.cancelled:
                script.exit()
            pb5.update_progress(n + 1, len(MEP_Liste))

            mepraum.werte_schreiben()

        t.Commit()
        t.Dispose()
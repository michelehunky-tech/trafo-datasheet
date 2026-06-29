import yaml
# (it_label, en, section, unit, decimals, family, include, required, suppress_if_zero)
F = [
 ("Codifica foglio calcolo","Calc sheet code","_internal",None,None,"both",False,False,False),
 ("Codice prodotto","Product code","header",None,None,"both",True,False,False),
 ("Serie","Series","general",None,None,"both",True,True,False),
 ("Tipo casa","Construction","general",None,None,"olio",True,False,False),
 ("Tipo sistema raffreddamento","Cooling system","general",None,None,"both",True,False,False),
 ("Raffreddamento","Cooling","general",None,None,"both",True,True,False),
 ("Temp. Ambiente MIN","Ambient temperature min","environmental","°C",0,"both",True,False,False),
 ("Temp. Ambiente MAX","Ambient temperature max","environmental","°C",0,"both",True,False,False),
 ("Altitudine","Max installation altitude","environmental","m a.s.l.",0,"both",True,False,False),
 ("THDi","Loading application (THDi)","general","%",1,"both",True,False,False),
 ("Potenza nominale","Rated power","ratings","kVA",0,"both",True,True,False),
 ("Potenza reg. forzato","Forced-cooling power","ratings","kVA",0,"olio",True,False,True),
 ("Potenza reattiva","Reactive power","ratings","kVAr",0,"both",False,False,True),
 ("Avvolg. stabilizzatore? Per trafo EAR","Stabilizing winding","_internal",None,None,"both",False,False,False),
 ("Avvolg. ausiliario? Per trafo EAR","Auxiliary winding","_internal",None,None,"both",False,False,False),
 ("Regolazione? Per trafo EAR","Regulation (EAR)","_internal",None,None,"both",False,False,False),
 ("Corrente guasto omopolare","Zero-sequence fault current","_internal",None,None,"both",False,False,False),
 ("Durata guasto omopolare","Zero-sequence fault duration","_internal",None,None,"both",False,False,False),
 ("Corrente guasto permanente","Permanent fault current","_internal",None,None,"both",False,False,False),
 ("Numero fasi","Number of phases","electrical",None,0,"both",True,True,False),
 ("Frequenza","Frequency","electrical","Hz",0,"both",True,True,False),
 ("Tensione MT1","HV voltage","ratings","V",0,"both",True,True,False),
 ("Tensione MT2","HV voltage (2nd)","ratings","V",0,"both",True,False,True),
 ("Tipo commutatore","Tap changer","ratings",None,None,"both",True,False,False),
 ("Posizioni + rif. MT1","Tap positions (+)","ratings",None,0,"both",True,False,False),
 ("Posizioni - rif. MT1","Tap positions (-)","ratings",None,0,"both",True,False,False),
 ("% gradino rif. MT1","Step per tap","ratings","%",2,"both",True,False,False,100),
 ("Tensione BT","LV voltage","ratings","V",0,"both",True,True,False),
 ("Collegamento MT","HV connection","electrical",None,None,"both",False,False,False),
 ("Collegamento BT","LV connection","electrical",None,None,"both",False,False,False),
 ("Collegamento TER","TER connection","electrical",None,None,"both",False,False,True),
 ("Indice orario","Clock index","electrical",None,0,"both",False,False,False),
 ("Gruppo vettoriale","Vector group","electrical",None,None,"both",True,True,False),
 ("Perdite a vuoto","No-load losses","electrical","W",0,"both",True,True,False),
 ("Perdite a carico 75°C","Load losses 75°C","electrical","W",0,"olio",True,True,False),
 ("Perdite a carico 120°C","Load losses 120°C","electrical","W",0,"resina",True,True,False),
 ("Corrente a vuoto %","No-load current","electrical","%",1,"both",True,False,False),
 ("Impedenza di cortocircuito %","Short-circuit impedance","electrical","%",2,"both",True,True,False),
 ("PEI / MEPS / HEPS in AN/ONAN","Efficiency index (PEI)","electrical","%",2,"both",True,False,False),
 ("Classe termica MT","HV thermal class","ratings",None,None,"both",True,False,False),
 ("Classe termica BT","LV thermal class","ratings",None,None,"both",True,False,False),
 ("Classe isolamento MT","HV insulation level","ratings","kV",None,"both",True,True,False),
 ("Classe isolamento BT","LV insulation level","ratings","kV",None,"both",True,True,False),
 ("Materiale MT","HV winding material","ratings",None,None,"both",True,False,False),
 ("Materiale BT","LV winding material","ratings",None,None,"both",True,False,False),
 ("Materiale ST","TER winding material","ratings",None,None,"both",True,False,True),
 ("Tipo avvolg. MT","HV winding type","ratings",None,None,"both",True,False,False),
 ("Tipo avvolg. BT","LV winding type","ratings",None,None,"both",True,False,False),
 ("Tipo avvolg. ST","TER winding type","ratings",None,None,"both",True,False,True),
 ("Sovratemperatura olio","Oil temperature rise","environmental","°C",0,"olio",True,False,False),
 ("Sovratemperatura avvolg. MT","HV winding temp. rise","environmental","°C",0,"both",True,False,False),
 ("Sovratemperatura avvolg. BT","LV winding temp. rise","environmental","°C",0,"both",True,False,False),
 ("LpA","Sound pressure level (LpA)","environmental","dBA",0,"both",True,False,False),
 ("LwA","Sound power level (LwA)","environmental","dBA",0,"both",True,False,False),
 ("Zo","Zero-seq. impedance Zo","_internal",None,None,"both",False,False,False),
 ("Ro","Zero-seq. resistance Ro","_internal",None,None,"both",False,False,False),
 ("Xo","Zero-seq. reactance Xo","_internal",None,None,"both",False,False,False),
 ("Classe amb / clim / fuoco","Environmental / climatic / fire class","general",None,None,"resina",True,True,False),
 ("Lunghezza trafo","Length","dimensions","mm",0,"both",True,False,False),
 ("Larghezza trafo","Width","dimensions","mm",0,"both",True,False,False),
 ("Altezza trafo","Height","dimensions","mm",0,"both",True,False,False),
 ("Peso totale maggiorato","Total weight","dimensions","kg",0,"both",True,False,False),
 ("Peso olio maggiorato","Oil weight","dimensions","kg",0,"olio",True,False,False),
 ("Tipologia olio","Oil type","general",None,None,"olio",True,True,False),
 ("Interasse ruote","Wheel gauge","dimensions","mm",0,"both",True,False,False),
 ("Ø ruote","Wheel diameter","dimensions","mm",0,"both",True,False,False),
 ("Cabina","Enclosure","_internal",None,None,"resina",False,False,False),
 ("K inserzione","Inrush factor","_internal",None,None,"both",False,False,False),
 ("Costante di tempo (sec)","Time constant","_internal","s",None,"both",False,False,False),
 ("Tempo emivalore (msec)","Half-value time","_internal","ms",None,"both",False,False,False),
 ("Induzione nucleo","Core induction","_internal","T",None,"both",False,False,False),
 ("Densità corrente MT (Sn/Pos Nom)","HV current density","_internal",None,None,"both",False,False,False),
 ("Densità corrente BT (Sn/Pos Nom)","LV current density","_internal",None,None,"both",False,False,False),
]
fields=[]
for row in F:
    it,en,sec,unit,dec,fam,inc,req,sz = row[:9]
    scale = row[9] if len(row)>9 else None
    d={"it":it,"en":en,"section":sec,"family":fam,"include_in_sheet":inc,"required":req}
    if unit is not None: d["unit"]=unit
    if dec is not None: d["decimals"]=dec
    if sz: d["suppress_if_zero"]=True
    if scale: d["scale"]=scale
    fields.append(d)

winding={"Strati":"Layer","Dischi":"Disc","Continuo":"Continuous disc","Elicoidale":"Helical","Bobine":"Coil"}
material={"Al":"Aluminium","Cu":"Copper"}
value_map={
 "Tipo casa":{"Conservatore":"Conservator","Ermetico":"Hermetically sealed"},
 "Tipo sistema raffreddamento":{"Radiatori":"Radiators","Onde":"Corrugated walls"},
 "Tipo commutatore":{"Reg. sottocarico":"On-load tap changer (OLTC)","A vuoto":"Off-circuit tap changer (DETC)","Reg. a vuoto":"Off-circuit tap changer (DETC)","Nessuno":"None"},
 "Tipo avvolg. MT":winding,"Tipo avvolg. BT":winding,"Tipo avvolg. ST":winding,
 "Materiale MT":material,"Materiale BT":material,"Materiale ST":material,
}
schema={
 "meta":{"sheet":"Foglio1","label_col":1,"value_col":2,"first_row":2,"last_row":73,
         "datasheet_language":"en","number_format":{"decimal":",","thousands":"."}},
 "family_rules":{"_doc":"Order: oil if 'Tipologia olio' present; resin if Raffreddamento in air set; else oil.",
   "air_cooling":["AN","AF","ANAF"],"oil_prefix":["O","K"]},
 "image_rules":{"_doc":"family+casa+cooling -> image key",
   "resina":{"cabina_true":"resin_enclosure","cabina_false":"resin_open"},
   "olio":{"Ermetico":"oil_hermetic",
           "Conservatore":{"Radiatori":"oil_conservator_radiators","Onde":"oil_conservator_corrugated"}}},
 "suspicious":{"date_coerced_fields":["Classe isolamento MT","Classe isolamento BT"]},
 "value_map":value_map,
 "fields":fields,
}
with open("parser/schema.yaml","w") as f:
    yaml.safe_dump(schema,f,sort_keys=False,allow_unicode=True,width=100)
print("schema.yaml written:", len(fields), "fields")

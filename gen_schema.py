import yaml

# Campi "semplici" (non avvolgimento): General / Electrical / Working Conditions / Dimensions.
# Avvolgimenti, tap, impedenze, efficiency, norme, accessori -> gestiti in extract.py.
# (it, en, section, unit, decimals, family, include, required, suppress_if_zero)
F = [
 ("Serie","Series","general",None,None,"both",True,True,False),
 ("Tipo casa","Construction","general",None,None,"olio",True,False,False),
 ("Tipo sistema raffreddamento","Cooling system","general",None,None,"both",True,False,False),
 ("Raffreddamento","Cooling","general",None,None,"both",True,True,False),
 ("Applicazione","Application","general",None,None,"both",True,False,False),
 ("Tipologia olio","Oil type","general",None,None,"olio",True,False,False),
 ("Classe amb / clim / fuoco","Environmental / climatic / fire class","general",None,None,"resina",True,False,False),
 ("Installazione","Installation","general",None,None,"both",True,False,False),
 ("Protezione superficiale","Surface protection","general",None,None,"both",True,False,False),
 ("Colore","Paint colour","general",None,None,"both",True,False,False),
 ("THDi","THDi","general","%",1,"both",True,False,False),

 ("Numero fasi","Number of phases","electrical",None,0,"both",True,True,False),
 ("Frequenza","Frequency","electrical","Hz",0,"both",True,True,False),
 ("Gruppo vettoriale","Vector group","electrical",None,None,"both",True,True,False),
 ("Perdite a vuoto","No-load losses","electrical","W",0,"both",True,True,False),
 ("Perdite a carico 75°C","Load losses 75°C","electrical","W",0,"olio",True,False,False),
 ("Perdite a carico 120°C","Load losses 120°C","electrical","W",0,"resina",True,False,False),
 ("Corrente a vuoto %","No-load current","electrical","%",1,"both",True,False,False),

 ("Temp. Ambiente MIN","Min ambient temperature","environmental","°C",0,"both",True,False,False),
 ("Temp. Ambiente MAX","Max ambient temperature","environmental","°C",0,"both",True,False,False),
 ("Altitudine","Max installation altitude","environmental","m a.s.l.",0,"both",True,False,False),
 ("Sovratemperatura olio","Oil temperature rise","environmental","°C",0,"olio",True,False,False),
 ("LpA","Sound pressure level (LpA)","environmental","dBA",0,"both",True,False,False),
 ("LwA","Sound power level (LwA)","environmental","dBA",0,"both",True,False,False),

 ("Lunghezza trafo","Length","dimensions","mm",0,"both",True,False,False),
 ("Larghezza trafo","Width","dimensions","mm",0,"both",True,False,False),
 ("Altezza trafo","Height","dimensions","mm",0,"both",True,False,False),
 ("Peso totale maggiorato","Total weight","dimensions","kg",0,"both",True,False,False),
 ("Peso olio maggiorato","Oil weight","dimensions","kg",0,"olio",True,False,False),
 ("Interasse ruote","Distance between wheels","dimensions","mm",0,"both",True,False,False),
 ("Ø ruote","Wheel diameter","dimensions","mm",0,"both",True,False,False),
]
fields=[]
for it,en,sec,unit,dec,fam,inc,req,sz in F:
    d={"it":it,"en":en,"section":sec,"family":fam,"include_in_sheet":inc,"required":req}
    if unit is not None: d["unit"]=unit
    if dec is not None: d["decimals"]=dec
    if sz: d["suppress_if_zero"]=True
    fields.append(d)

winding={"Strati":"Layer","Dischi":"Disc","Disco":"Disc","Continuo":"Continuous disc","Elicoidale":"Helical",
         "Bobine":"Coil","Inglobata":"Cast (encapsulated)","Impregnata":"Impregnated"}
material={"Al":"Aluminium","Cu":"Copper"}
value_map={
 "Tipo casa":{"Conservatore":"Conservator","Ermetico":"Hermetically sealed"},
 "Tipo sistema raffreddamento":{"Radiatori":"Radiators","Onde":"Corrugated walls","Scambiatore":"Heat exchanger"},
 "Tipo commutatore":{"Reg. sottocarico":"On-load tap changer (OLTC)","A vuoto":"Off-circuit tap changer (DETC)",
                     "Reg. a vuoto":"Off-circuit tap changer (DETC)","Nessuno":"None"},
 "Applicazione":{"Distribuzione":"Distribution","Conversione":"Conversion","Trasmissione":"Transmission",
                 "Trazione":"Traction","Fotovoltaico":"Photovoltaic","Eolico":"Wind"},
 "Installazione":{"Interno / Esterno":"Indoor / Outdoor","Interno":"Indoor","Esterno":"Outdoor"},
 "Materiale MT":material,"Materiale BT1":material,"Materiale BT2":material,
 "Tipo avvolg. MT":winding,"Tipo avvolg. BT1":winding,"Tipo avvolg. BT2":winding,
}
schema={
 "meta":{"sheet":"Foglio1","label_col":1,"value_col":2,"first_row":2,"last_row":94,
         "accessories_marker":"Accessori:","datasheet_language":"en",
         "number_format":{"decimal":",","thousands":"."}},
 "family_rules":{"_doc":"oil if 'Tipologia olio' present; resin if air cooling; else oil.",
   "air_cooling":["AN","AF","ANAF"],"oil_prefix":["O","K"]},
 "image_rules":{"resina":{"cabina_true":"resin_enclosure","cabina_false":"resin_open"},
   "olio":{"Ermetico":"oil_hermetic",
           "Conservatore":{"Radiatori":"oil_conservator_radiators","Onde":"oil_conservator_corrugated"}}},
 "suspicious":{"date_coerced_fields":["Classe isolamento MT","Classe isolamento BT1","Classe isolamento BT2"]},
 "voltage_class":{"lv_max":3600,"mv_max":52000},   # LV<=3600, MV<=52000, HV>52000
 "efficiency":{"pei_when":["548"],"meps_when":["AS 60076","AS60076"]},
 "value_map":value_map,
 "fields":fields,
}
with open("parser/schema.yaml","w") as f:
    yaml.safe_dump(schema,f,sort_keys=False,allow_unicode=True,width=100)
print("schema.yaml:",len(fields),"campi semplici")

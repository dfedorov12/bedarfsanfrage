'use strict';

// ── CONFIG ──────────────────────────────────────────────────────────────────
const CLIENT_ID = '37e2a0cc-37b9-4eb5-b2a1-2ca0d9a62db0';
const TENANT_ID = 'fdb70646-023a-403b-a4b9-1f474a935123';
const SCOPES    = ['User.Read', 'Sites.ReadWrite.All'];
const SP_SITE   = 'dihag.sharepoint.com:/sites/gruppe_shb';
const SP_LIST   = 'Bedarfsanfrage';
const API       = 'https://graph.microsoft.com/v1.0';
const SP_BASE   = 'https://' + SP_SITE.split(':/')[0] + '/' + SP_SITE.split(':/')[1]; // split(':/')[1] = 'sites/...' (no leading slash)

// ── BPMN-KONFORME FORMFELD-DEFINITION ───────────────────────────────────────
// key        = interner SP-Spaltenname (wird beim POST verwendet)
// alsoTry    = alternative interne Namen falls key nicht gefunden
// step       = Wizard-Schritt (1-3)
const FORM_FIELDS = [
  // Step 1: Bedarf
  { key:'Title',             label:'Bezeichnung',                  step:1, required:true  },
  { key:'Beschreibung',      label:'Beschreibung',                 step:1, alsoTry:['Description','Beschreibung_x002f_Begruendung','Grund'] },
  { key:'Warengruppe',       label:'Warengruppe',                  step:1, required:true, alsoTry:['ProductCategory'] },
  { key:'Prioritaet',        label:'Priorität',                    step:1, alsoTry:['Priority','Priorit_x00e4_t'] },
  // Step 2: Menge
  { key:'Menge',             label:'Menge',                        step:2, required:true, alsoTry:['Quantity','Amount'] },
  { key:'Mengeneinheit',     label:'Mengeneinheit',                step:2, required:true, alsoTry:['Unit','UnitOfMeasure'] },
  { key:'Mindestlagermenge', label:'Mindestlagermenge',            step:2, alsoTry:['MinStock','MinLager'] },
  { key:'Termin',            label:'Benötigt bis',                 step:2, required:true, alsoTry:['Deadline','DueDate','Ben_x00f6_tigtBis','Ben_x00f6_tigtbis'] },
  // Step 3: Beschaffung
  { key:'Artikelnummer',     label:'Artikelnummer',                 step:1, alsoTry:['MaterialNumber','ItemNumber','Artikelnummer_x002f_Nummernangab'] },
  { key:'Beschaffungslogik', label:'Beschaffungsart',              step:3, required:true, alsoTry:['Materialtyp','ProcurementType'] },
  { key:'Lieferant',         label:'Lieferant 1',                  step:3, alsoTry:['Vendor','Supplier'] },
  { key:'Lieferant2',        label:'Lieferant 2 (Alternative)',    step:3, alsoTry:['Vendor2','Supplier2','Lieferant_2'] },
  { key:'Lieferant3',        label:'Lieferant 3 (Alternative)',    step:3, alsoTry:['Vendor3','Supplier3','Lieferant_3'] },
  { key:'Lieferant4',        label:'Lieferant 4 (Alternative)',    step:3, alsoTry:['Vendor4','Supplier4','Lieferant_4'] },
  { key:'GeschaetzterPreis',    label:'Geschätzter Preis',            step:3, alsoTry:['EstimatedPrice','Preis','Price','Gesch_x00e4_tzterPreisnetto_x002'] },
  { key:'Kostenstelle',         label:'Kostenstelle',                 step:3, alsoTry:['CostCenter'] },
  { key:'LeadBuyerAbschluss',   label:'Lead-Buyer-Abschluss',         step:3, alsoTry:['LeadBuyer','LeadBuyerAbschlus'] },
];

// Felder die Einkauf nach der Einreichung befüllt
const EINKAUF_FIELDS = [
  { key:'Bestellnummer',  label:'Bestellnummer',        alsoTry:['OrderNumber','PO_Number'] },
  { key:'Lieferdatum',    label:'Lieferdatum',          alsoTry:['DeliveryDate'] },
  { key:'TatsaechlicherPreis', label:'Tatsächlicher Preis (€)', alsoTry:['ActualPrice','FinalPrice'] },
];

// Status-Werte → Darstellung (kommt von Power Automate)
// Reihenfolge wichtig: spezifischere (exakte) Werte zuerst, Fallback-Partials am Ende
const STATUS_STYLES = {
  // ── Exakte SP-Auswahlwerte ──
  'eingereicht':                        { bg:'#fce7f3', color:'#be185d' },
  'in prüfung (einkauf)':               { bg:'#ccfbf1', color:'#0f766e' },
  'in prüfung (werkleitung)':           { bg:'#ffe4e6', color:'#be123c' },
  'in prüfung (strategischer einkauf)': { bg:'#ccfbf1', color:'#0f766e' },
  'in prüfung (controlling)':           { bg:'#ccfbf1', color:'#0f766e' },
  'freigegeben':                        { bg:'#f3e8ff', color:'#7e22ce' },
  'abgelehnt':                          { bg:'#f3f4f6', color:'#374151' },
  'bestellt':                           { bg:'#dbeafe', color:'#1d4ed8' },
  'erledigt':                           { bg:'#f3f4f6', color:'#374151' },
  // ── Generische Fallbacks (Teilübereinstimmung) ──
  'in prüfung':                         { bg:'#ccfbf1', color:'#0f766e' },
  'angefragt':                          { bg:'#f0f9ff', color:'#0369a1' },
  'in bearbeitung':                     { bg:'#fffbeb', color:'#b45309' },
  'offen':                              { bg:'#eff6ff', color:'#1d4ed8' },
};

const SYSTEM_FIELDS = new Set([
  '@odata.etag','@odata.id','id','ContentType','Modified','Created',
  'AuthorLookupId','EditorLookupId','Attachments','Edit','LinkTitleNoMenu','LinkTitle',
  'ItemChildCount','FolderChildCount','_UIVersionString','ComplianceAssetId',
  'OData__ColorTag','AppAuthorLookupId','AppEditorLookupId'
]);

// ── ARTIKEL-STAMMDATEN (TID → Bezeichnung / Warengruppe) ─────────────────────
const TID_MAP = {
  '4001-00010': {b:'Grobblech/Kupolofenschrott unleg', w:'Schrott', g:'Rohstoffe'},
  '4001-00015': {b:'Schrott / unlegiert /Stanz- und', w:'Schrott', g:'Rohstoffe'},
  '4001-00020': {b:'Schrott / Sonder', w:'Schrott', g:'Rohstoffe'},
  '4001-00021': {b:'Schrott legiert WN 1.8983', w:'Schrott', g:'Rohstoffe'},
  '4001-00022': {b:'Schrott legiert WN 1.2740', w:'Schrott', g:'Rohstoffe'},
  '4001-00023': {b:'Schrott legiert WN 1.2767', w:'Schrott', g:'Rohstoffe'},
  '4001-00024': {b:'Schrott legiert WN 1.2713/1.2714', w:'Schrott', g:'Rohstoffe'},
  '4001-00040': {b:'TZ-Stanzabfälle/Pakete. 0.25 Mn', w:'Schrott', g:'Rohstoffe'},
  '4001-00090': {b:'Kreislauf / Späne unleg. / Eigen', w:'Schrott', g:'Rohstoffe'},
  '4001-00100': {b:'Kreislauf / unlegiert', w:'Schrott', g:'Rohstoffe'},
  '4001-00103': {b:'Kreislauf / niedriglegiert', w:'Schrott', g:'Rohstoffe'},
  '4001-00109': {b:'Kreislauf Grauguss (GJL)', w:'Schrott', g:'Rohstoffe'},
  '4001-00119': {b:'Kreislauf Sphäroguss (GJS)', w:'Schrott', g:'Rohstoffe'},
  '4001-00120': {b:'Gussbruch 2A', w:'Schrott', g:'Rohstoffe'},
  '4002-00010': {b:'Stahlwerkserz  6-25 mm (oder grö', w:'Legierungen', g:'Rohstoffe'},
  '4002-00020': {b:'Reinsteisen', w:'Legierungen', g:'Rohstoffe'},
  '4002-00030': {b:'Roheisen (Stahlroheisen)', w:'Legierungen', g:'Rohstoffe'},
  '4002-00031': {b:'Hämatitroheisen (GJL)', w:'Legierungen', g:'Rohstoffe'},
  '4002-00032': {b:'Brasilianisches Roheisen T5', w:'Legierungen', g:'Rohstoffe'},
  '4002-00110': {b:'Aluminium Zweiteiler min. 97 % A', w:'Legierungen', g:'Rohstoffe'},
  '4002-00111': {b:'Aluminium Zehnteiler min. 97 % A', w:'Legierungen', g:'Rohstoffe'},
  '4002-00170': {b:'Chrom-Metall / aluminotherm.', w:'Legierungen', g:'Rohstoffe'},
  '4002-00180': {b:'Kupfer', w:'Legierungen', g:'Rohstoffe'},
  '4002-00190': {b:'Mangan-Metall min. 99 % Mn', w:'Legierungen', g:'Rohstoffe'},
  '4002-00200': {b:'Molybdän-Metall', w:'Legierungen', g:'Rohstoffe'},
  '4002-00210': {b:'Nickel 99 % Ni, als Kathoden 4x4', w:'Legierungen', g:'Rohstoffe'},
  '4002-00250': {b:'Silico-Mangan 65 % Mn, 15 % Si,', w:'Legierungen', g:'Rohstoffe'},
  '4002-00510': {b:'Ferro Chrom aff. ca 68 % Cr, Kör', w:'Legierungen', g:'Rohstoffe'},
  '4002-00520': {b:'Ferro Chrom carburé, ca. 65 % Cr', w:'Legierungen', g:'Rohstoffe'},
  '4002-00530': {b:'Ferro Chrom suraff. / C max. 0,0', w:'Legierungen', g:'Rohstoffe'},
  '4002-00540': {b:'Ferro Chrom suraff. N-haltig', w:'Legierungen', g:'Rohstoffe'},
  '4002-00550': {b:'Ferro Mangan aff. ca. 78 % Mn, K', w:'Legierungen', g:'Rohstoffe'},
  '4002-00560': {b:'Ferro Mangan carb. ca. 75 % Mn,', w:'Legierungen', g:'Rohstoffe'},
  '4002-00570': {b:'Ferro Molybdän 65 % Mo, Körnung', w:'Legierungen', g:'Rohstoffe'},
  '4002-00590': {b:'Ferro Silizium 75 %, fein 0-10 m', w:'Legierungen', g:'Rohstoffe'},
  '4002-00600': {b:'Ferro Silizium 75 stückig 10 - 8', w:'Legierungen', g:'Rohstoffe'},
  '4002-00620': {b:'Ferro-Titan', w:'Legierungen', g:'Rohstoffe'},
  '4002-00630': {b:'Ferro-Vanadium, ca. 80 % V, Körn', w:'Legierungen', g:'Rohstoffe'},
  '4002-00640': {b:'Ferro-Wolfram', w:'Legierungen', g:'Rohstoffe'},
  '4002-00650': {b:'Ferro Zirkon  ca. 75 %', w:'Legierungen', g:'Rohstoffe'},
  '4116-00028': {b:'Buchse 465-8391 Ø100xØ120xØ80,4x', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00029': {b:'Bolzen CAT 619-3910', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00031': {b:'Sechskantschraube MJ30x3,5 101-1', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00033': {b:'Unterlegscheibe 198-4772, 33x60x', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00036': {b:'Zwischenstück 521-2995, 31,5x60x', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00037': {b:'Sechskantmutter M30x3,5 8T-1583', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00054': {b:'Laufbuchse 340/250,5x256', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00082': {b:'Buchse 200r6/180,3xF7x125+/-0,5', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00134': {b:'Bolzen RD 65x452+/-0,5 mm 30CrNi', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00171': {b:'Blech für Tragarm lt. Skizze,Mat', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00174': {b:'RUD Anschlagpunkt ABA 3,2 t', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00175': {b:'RUD Anschlagpunkt ABA 5,0 t', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00186': {b:'Anhängeöse für Umschmelzblock au', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00187': {b:'Anhängeöse für Umschmelzblock au', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00188': {b:'Anhängeöse für Hubzylinderlageru', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00200': {b:'Buchse 477-7939 CAT Kettensegmen', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00201': {b:'Buchse 477-9504 CAT Kettensegmen', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00202': {b:'Bolzen 465-8403 CAT Kettensegmen', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00203': {b:'Teller 465-8401 CAT Kettensegmen', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00204': {b:'Ring 7U-3538 CAT Kettensegmentmo', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00205': {b:'Platte 1 Fertigungsformat 29x20x', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00206': {b:'Platte 2 Fertigungsformat 4x17x2', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00207': {b:'Bolzen 465-8489', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00208': {b:'Bolzen 465-8480', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00210': {b:'Buchse 503-0041', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00211': {b:'Buchse 493-9690', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00212': {b:'Buchse 493-9691', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00213': {b:'Platte 465-8490, 149,80x149,80x2', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00214': {b:'Sechskantschraube 493-1477 931 S', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00215': {b:'Sicherheitsring 109-0119', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00216': {b:'Unterlegscheibe 493-1472', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00217': {b:'Zwischenstück 465-8481', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00218': {b:'Platte 4 Fertigungsformat 20x160', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00219': {b:'Buchse 503-0040 für Kettenglied', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00220': {b:'Buchse 503-0041 für Kettenglied', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00221': {b:'Buchse 477-7939 für Kettenglied', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00222': {b:'Buchse 477-9504 für Kettenglied', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00223': {b:'Sicherungsring DNH-150 611-4017', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00224': {b:'Bolzen CAT 626-7857', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00225': {b:'Platte CAT 616-7594', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00300': {b:'Body 120 B Fertigbearbeitet', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00301': {b:'Cup', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4116-00303': {b:'Body 152 B Rohguss mit Kopfspeis', w:'Kaufteile', g:'Hilfs- und Betriebsstoffe'},
  '4211-00220': {b:'Isopropanol, kosmet. Qualität (i', w:'Kaufteile Formerei (Kühlm. ..)', g:'Hilfs- und Betriebsstoffe'},
  '4211-00325': {b:'Teno Coat. 26AC  CONC 2 (fr. Mol', w:'Kaufteile Formerei (Kühlm. ..)', g:'Hilfs- und Betriebsstoffe'},
  '4211-00327': {b:'Hardcote 8310 / 35-kg-Hob.', w:'Kaufteile Formerei (Kühlm. ..)', g:'Hilfs- und Betriebsstoffe'},
  '4211-00332': {b:'Teno coating ZKPX', w:'Kaufteile Formerei (Kühlm. ..)', g:'Hilfs- und Betriebsstoffe'},
  '4211-00333': {b:'Teno Sil 1320 E', w:'Kaufteile Formerei (Kühlm. ..)', g:'Hilfs- und Betriebsstoffe'},
  '4211-00335': {b:'Semco Tec 78020 ML im 65 kg-Hobb', w:'Kaufteile Formerei (Kühlm. ..)', g:'Hilfs- und Betriebsstoffe'},
  '4211-00336': {b:'Semco Perm  M 70  /  75-kg-Fass', w:'Kaufteile Formerei (Kühlm. ..)', g:'Hilfs- und Betriebsstoffe'},
  '4211-00337': {b:'Semco Zir 4300 E 50 kg/DR', w:'Kaufteile Formerei (Kühlm. ..)', g:'Hilfs- und Betriebsstoffe'},
  '4211-00338': {b:'Semco Sir 6330 C', w:'Kaufteile Formerei (Kühlm. ..)', g:'Hilfs- und Betriebsstoffe'},
  '4211-00362': {b:'Ecopart 56 L / 8-kg-Kan.', w:'Kaufteile Formerei (Kühlm. ..)', g:'Hilfs- und Betriebsstoffe'},
  '4211-00363': {b:'Ecopart 56 D / 500 ml-Spraydose', w:'Kaufteile Formerei (Kühlm. ..)', g:'Hilfs- und Betriebsstoffe'},
  '4211-00365': {b:'Bentogliss 121 / 170-kg-Faß', w:'Kaufteile Formerei (Kühlm. ..)', g:'Hilfs- und Betriebsstoffe'},
  '4211-00366': {b:'Askopaste GF 2 / 2-kg-Btl.', w:'Kaufteile Formerei (Kühlm. ..)', g:'Hilfs- und Betriebsstoffe'},
  '4211-00406': {b:'Kernkleber / á 1 kg-Corfix', w:'Kaufteile Formerei (Kühlm. ..)', g:'Hilfs- und Betriebsstoffe'},
  '4211-00412': {b:'Formlack 1572', w:'Kaufteile Formerei (Kühlm. ..)', g:'Hilfs- und Betriebsstoffe'},
  '4211-00415': {b:'Speseal asbestfreie Abdichtmasse', w:'Kaufteile Formerei (Kühlm. ..)', g:'Hilfs- und Betriebsstoffe'},
  '4211-00500': {b:'Teno Coating  ZBBP-16', w:'Kaufteile Formerei (Kühlm. ..)', g:'Hilfs- und Betriebsstoffe'},
  '4304-00005': {b:'Ferrosad-Gemisch F 13 X im Big b', w:'Strahlmittel', g:'Hilfs- und Betriebsstoffe'},
  '4304-00010': {b:'Strahlmittel F16', w:'Strahlmittel', g:'Hilfs- und Betriebsstoffe'},
  '4304-00015': {b:'Strahlmittel F 34  (0,8mm)', w:'Strahlmittel', g:'Hilfs- und Betriebsstoffe'},
  '4306-00016': {b:'Quarzsand  WF 322/331', w:'Formstoffkomponenten', g:'Hilfs- und Betriebsstoffe'},
  '4306-00017': {b:'Quedlinburger Quarzsand QQs 37', w:'Formstoffkomponenten', g:'Hilfs- und Betriebsstoffe'},
  '4306-00028': {b:'Quarzsand, getrocknet  GS 13', w:'Formstoffkomponenten', g:'Hilfs- und Betriebsstoffe'},
  '4306-00050': {b:'Chromerzsand AFS 45-55 im Big Ba', w:'Formstoffkomponenten', g:'Hilfs- und Betriebsstoffe'},
  '4306-00051': {b:'Chromerzsand AFS 45-55 im Silo', w:'Formstoffkomponenten', g:'Hilfs- und Betriebsstoffe'},
  '4306-00055': {b:'Olivinsand  AFS 50 (TS 24)   in', w:'Formstoffkomponenten', g:'Hilfs- und Betriebsstoffe'},
  '4306-00060': {b:'Zirkonsand / 25-kg-Säcke', w:'Formstoffkomponenten', g:'Hilfs- und Betriebsstoffe'},
  '4306-00175': {b:'Spezialsand M-Sand / in 25-kg-Sä', w:'Formstoffkomponenten', g:'Hilfs- und Betriebsstoffe'},
  '4306-00211': {b:'Ecosil D04 PV 4842 +0,5% Amylon', w:'Formstoffkomponenten', g:'Hilfs- und Betriebsstoffe'},
  '4306-00212': {b:'Ecosil D04 R90 0,5A', w:'Formstoffkomponenten', g:'Hilfs- und Betriebsstoffe'},
  '4306-00213': {b:'Thorbent 100 C11G - SHB lose im', w:'Formstoffkomponenten', g:'Hilfs- und Betriebsstoffe'},
  '4306-00216': {b:'Ceratec 50 (S-K) in 25 kg-Säcken', w:'Formstoffkomponenten', g:'Hilfs- und Betriebsstoffe'},
  '4306-00217': {b:'Bauxitsand (Ceratec 50) im Big B', w:'Formstoffkomponenten', g:'Hilfs- und Betriebsstoffe'},
  '4306-00226': {b:'Eisenoxyd FE 60', w:'Formstoffkomponenten', g:'Hilfs- und Betriebsstoffe'},
  '4306-00350': {b:'Novanol 165 im 1000l Container', w:'Formstoffkomponenten', g:'Hilfs- und Betriebsstoffe'},
  '4306-00355': {b:'PEP Set Blue EP 5168 Part 1, Sta', w:'Formstoffkomponenten', g:'Hilfs- und Betriebsstoffe'},
  '4306-00356': {b:'PEP SET Blue EP 5188 Part 2, Sta', w:'Formstoffkomponenten', g:'Hilfs- und Betriebsstoffe'},
  '4306-00360': {b:'Askocure 388 Part 1 im 1000l Con', w:'Formstoffkomponenten', g:'Hilfs- und Betriebsstoffe'},
  '4306-00361': {b:'Askocure 699 Part 2 im 1000l Con', w:'Formstoffkomponenten', g:'Hilfs- und Betriebsstoffe'},
  '4306-00362': {b:'Katalysator 704 / 140kg Faß', w:'Formstoffkomponenten', g:'Hilfs- und Betriebsstoffe'},
  '4306-00502': {b:'Alphaset ACE 1075 Härter', w:'Formstoffkomponenten', g:'Hilfs- und Betriebsstoffe'},
  '4306-00503': {b:'Alphaset ACE 1535/10 S Härter', w:'Formstoffkomponenten', g:'Hilfs- und Betriebsstoffe'},
  '4306-00508': {b:'Alphaset-Harz TPA 180', w:'Formstoffkomponenten', g:'Hilfs- und Betriebsstoffe'},
  '4306-00531': {b:'Novaset - Harz NB 720', w:'Formstoffkomponenten', g:'Hilfs- und Betriebsstoffe'},
  '4306-00533': {b:'Katalysator 7040 (f. Novaset NB', w:'Formstoffkomponenten', g:'Hilfs- und Betriebsstoffe'},
  '4306-00534': {b:'Katalysator 7090 (f. Novaset NB', w:'Formstoffkomponenten', g:'Hilfs- und Betriebsstoffe'},
  '4307-00102': {b:'Gießtrichter GT 40/150 mF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00103': {b:'Trichter ET 50/100/95-1 F 36', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00105': {b:'Gießtrichter GT 50/150 mF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00110': {b:'Gießtrichter  GT 70/250 mF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00112': {b:'Gießtrichter  GT 80/250 mF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00114': {b:'Trichterrohr  TR 100/250/30', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00116': {b:'Trichterrohr  TR 120/250/30', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00122': {b:'L-Stück  LKR 40 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00125': {b:'L-Stück  LKR 50 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00128': {b:'L-Stück  LKR 60 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00130': {b:'L-Stück  LKR 70 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00132': {b:'L-Stück  LKR 80 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00133': {b:'L-Stück  L100/100  mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00134': {b:'L-Stück  LKR 100/100  mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00136': {b:'L-Stück  LKR 120 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00141': {b:'Rohr   R 40/ 35 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00145': {b:'Rohr   R 40/ 50 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00148': {b:'Rohr   R 40/100 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00150': {b:'Rohr   R 40/150 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00153': {b:'Rohr   R 40/200 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00155': {b:'Rohr    R 40/250 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00160': {b:'Rohr   R 50/ 50 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00163': {b:'Rohr   R 50/100 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00166': {b:'Rohr   R 50/150 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00170': {b:'Rohr   R 50/200 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00173': {b:'Rohr   R 50/250 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00176': {b:'Rohr   R 50/300 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00180': {b:'Rohr   R 60/ 50 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00182': {b:'Rohr   R 60/100 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00184': {b:'Rohr   R 60/150 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00185': {b:'Rohr   R 60/200 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00187': {b:'Rohr   R 60/250 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00189': {b:'Rohr   R 60/300 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00195': {b:'Rohr   R 70/ 50 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00196': {b:'Rohr   R 70/100 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00197': {b:'Rohr  R 70/150 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00198': {b:'Rohr   R 70/200 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00200': {b:'Rohr   R 70/250 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00203': {b:'Rohr   R 70/300 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00208': {b:'Rohr   R 80/ 50 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00210': {b:'Rohr   R 80/100 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00212': {b:'Rohr   R 80/150 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00213': {b:'Rohr   R 80/200 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00215': {b:'Rohr   R 80/250 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00220': {b:'Rohr  R 80/300 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00221': {b:'Rohr   R 120/ 50 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00222': {b:'Rohr   R 120/150 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00223': {b:'Rohr   R 120/200 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00224': {b:'Rohr   R 120/250 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00225': {b:'Rohr   R 120/300 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00226': {b:'Rohr  R120/100', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00227': {b:'ROHR R 150/50', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00228': {b:'Rohr R 150/200', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00230': {b:'Rohr   R 100/ 50 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00232': {b:'Rohr   R 100/100 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00234': {b:'Rohr   R 100/150 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00235': {b:'Rohr   R 100/200 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00236': {b:'Rohr   R 100/250 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00238': {b:'Rohr   R 100/300 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00239': {b:'Rohr 150/300', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00240': {b:'Rohr   RNN 40', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00241': {b:'Rohr   RNN 50', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00242': {b:'Rohr   RNN 60', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00243': {b:'Rohr   RNN 70', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00244': {b:'Rohr   RNN 80', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00245': {b:'Rohr   RNN 100', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00246': {b:'Rohr   RFF 40', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00247': {b:'Rohr   RFF 50', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00248': {b:'Rohr   RFF 60', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00249': {b:'Rohr   RFF 70', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00250': {b:'Rohr   RFF 80', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00251': {b:'Rohr RFF 150', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00252': {b:'Rohr   RFF 100', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00253': {b:'Rohr   RFF 120', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00254': {b:'T-Stück    T 40/40/40 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00255': {b:'T-Stück    T 40/50/40 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00258': {b:'T-Stück   T 40/60/40 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00265': {b:'T-Stück   T 50/60/50 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00266': {b:'T-Stück   T 50/50/50 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00267': {b:'T-Stück   T 60/60/60 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00268': {b:'T-Stück   T 50/70/50 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00269': {b:'T-Stück   T 60/80/60 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00270': {b:'T-Stück   T 60/70/60 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00271': {b:'T-Stück    T 70/70/70 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00272': {b:'T-Stück   T 70/50/70 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00273': {b:'T-Stück    T 60/50/60 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00274': {b:'T-Stück  T 70/60/70', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00275': {b:'T-Stück  T 70/80/70 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00277': {b:'T-Stück   T 80/80/80 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00280': {b:'T-Stück   T 80/100/80 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00281': {b:'T-Stück   T 100/80/100 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00282': {b:'T-Stück   T 80/120/80 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00283': {b:'T-Stück   T 120/100/120 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00284': {b:'T-Stück   T 100/120/100 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00285': {b:'T Stück 150/150/150', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00300': {b:'Verteiler KV 60', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00301': {b:'Kreuz Verteiler KV 40', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00309': {b:'VSTA 80/4x50', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00312': {b:'Verteilerstein    VTSA 60/ 4x40', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00313': {b:'Verteilerstein    VTSA 80/4x60', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00314': {b:'Verteilerstein VTSA 70/4x50', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00315': {b:'Verteilerstein    VTSA 70/3x50', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00316': {b:'Verteiler, gekl.  100/6x60', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00317': {b:'Verteilerstein    VTSA 60/ 3x50', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00318': {b:'Kreuz-Verteiler  KV 80 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00319': {b:'Kreuz-Verteiler  KV 120 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00321': {b:'Halbschale   HS 7518/ mit einem', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00324': {b:'Halbschale   HS 94042 n.Zg. 1230', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00326': {b:'Halbschale    HS 7839', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00327': {b:'Halbschale    HS 7517', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00328': {b:'Halbschale    HS 6539', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00329': {b:'Aufschlagst. / Mittelstein / Zg.', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00331': {b:'Aufschlagst. / Seitenstein / Zg.', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00333': {b:'Verteiler, gekl.  120/6x60', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00334': {b:'Verteiler, gekl. 80/6x60', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00338': {b:'Schräganschnittrohr SAR 40/45°', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00341': {b:'Eingußrohr n. Zg. 124280', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00342': {b:'Eingußrohr Nr. 124341 / mit 2 se', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00343': {b:'Schräganschnittrohr SAR 70/45°', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00344': {b:'Schräganschnittrohr SAR 60/45°', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00345': {b:'Schräganschnittrohr SAR 50/45°', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00346': {b:'Reduzierstück       RE 70/60', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00347': {b:'Reduzierstück       RE 60/50', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00348': {b:'Reduzierstück  RE 80/70', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00349': {b:'Schamotteplättchen   80 x 110 x', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00352': {b:'Schräganschnittrohr SAR 80/45°', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00360': {b:'Reduzierstück    RE 120/100', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00361': {b:'Reduzierstück RE 100/80/50', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00362': {b:'Reduzierstück  RE 150/120', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00363': {b:'Reduzierstück RE 50/40', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00401': {b:'Reduzierstück RE 50/40 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00402': {b:'Anschnittrohr AR 80/50 mF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00403': {b:'Reduzierstück RE 80/60 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00404': {b:'Reduzierstück RE 100/80/150', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00405': {b:'Anschnittstein ARB 80/60 mF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00406': {b:'Anschnittstein ARB 80/60 mN', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00408': {b:'Anschnittstein AM 40/25 mit Nut', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00410': {b:'Anschnittrohr   AR 40/50  mN', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00412': {b:'Anschnittrohr   AR 50/50 mN', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00414': {b:'Anschnittrohr   AR 60/50 mN', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00415': {b:'Anschnittrohr AR 70/50 mN', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00416': {b:'Schräganschnittrohr SAR 100/45°', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00420': {b:'Messeranschnittsteine  AMH 40 mN', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00422': {b:'Messeranschnittsteine  AMH 50 mN', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00424': {b:'Messeranschnittstein  AMH 60 mN', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00425': {b:'Messeranschnittstein AMH 70 mN', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00426': {b:'Messeranschnittstein AMH 80 mN', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00440': {b:'AER 50/60 mF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00441': {b:'AER 60/70 mF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00444': {b:'ER  40/50 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00445': {b:'ER 60/80 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00446': {b:'ER  50/60 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00447': {b:'ER  60/70 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00448': {b:'ER 70/80 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00456': {b:'ER 100/120 mNuF', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00457': {b:'Erweiterungsstück 120/150', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00730': {b:'Filterkreisel 6/3Q rechts, inkl.', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00733': {b:'Filterkreisel 6/4Q links, inkl.', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00740': {b:'Filterkreisel 6/8Q links, inkl.', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-00768': {b:'Filterhalterset f. 150er Filter', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-01113': {b:'Aufschlagstein  AK 60 A / 1 Abga', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-01114': {b:'Aufschlagstein AK 60 A / 2 Abgän', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-96197': {b:'K-Lager Hagenburger', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4307-97910': {b:'K-Lager Krause / Keramik', w:'Eingußsysteme', g:'Hilfs- und Betriebsstoffe'},
  '4308-00005': {b:'TELE-FEEDER 100-18 (35)', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00007': {b:'Kalminex FF2000 Einsatz FZF 7/10', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00102': {b:'Speiser X 3', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00103': {b:'Speiser SD 3', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00104': {b:'Speiser X 4', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00105': {b:'Speiser SD 4', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00106': {b:'Speiser X 5', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00107': {b:'Speiser SD 5', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00110': {b:'Speiser X 6', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00111': {b:'Speiser SD 6', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00115': {b:'Speiser X 7', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00116': {b:'Speiser SD 7', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00117': {b:'Speiser SD 8', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00118': {b:'Speiser SD 9', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00120': {b:'Speiser X 8', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00122': {b:'Speiser X 9', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00123': {b:'Speiser X 9,5', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00124': {b:'Speiser X 10', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00126': {b:'Speiser X 11', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00130': {b:'Speiser X 12', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00132': {b:'Speiser X 13', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00134': {b:'Speiser X 14', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00136': {b:'Speiser X 15', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00137': {b:'Speiser SD 10', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00138': {b:'Speiser X 16', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00150': {b:'Speiser X 17', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00152': {b:'Speiser X 18', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00220': {b:'Speiser OX 4', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00222': {b:'Speiser OSD 4', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00223': {b:'Speiser OSD  4 / 75', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00224': {b:'Speiser OX 5', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00225': {b:'Speiser OSD 5', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00226': {b:'Speiser OSD 6', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00228': {b:'Speiser OX 6', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00231': {b:'Speiser OSD 8/11 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00232': {b:'Speiser OX 7', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00234': {b:'Speiser OSD 8', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00236': {b:'Speiser OX 8', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00240': {b:'Speiser OX 9', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00244': {b:'Speiser OX 10', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00246': {b:'Speiser OX 10-15', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00250': {b:'Speiser OX 11', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00252': {b:'Speiser OX 11-15', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00254': {b:'Speiser OX 12-20', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00255': {b:'Speiser OX 13-20', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00257': {b:'Speiser OX 14-20', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00260': {b:'Speiser OX 16-20', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00263': {b:'Speiser OX 18-20', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00265': {b:'Speiser OX 20-20', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00267': {b:'Speiser OX 22-20', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00270': {b:'Speiser OX 25-20', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00271': {b:'Speiser OX 27-20', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00272': {b:'Speiser OX 30-20', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00308': {b:'Speiser ZTA 0', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00312': {b:'Speiser ZTA 1', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00316': {b:'Speiser ZTA 2', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00318': {b:'Speiser ZTA 2R', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00320': {b:'Speiser ZTA 3', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00324': {b:'Speiser ZTA 4', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00328': {b:'Speiser ZTA 4-5', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00332': {b:'Speiser ZTA 5', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00334': {b:'Speiser XTA 6', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00335': {b:'Speiser ZTA 5-6', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00336': {b:'Speiser XTA 7 GT', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00338': {b:'Speiser XTA 7A/10CC', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00339': {b:'Speiser XTA 6P', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00340': {b:'Speiser XTA 8 GT', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00342': {b:'Speiser XTA 8A/10CC', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00344': {b:'Speiser XTA 9 GT', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00346': {b:'Speiser XTA 10 GT', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00347': {b:'Speiser XTA 10', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00350': {b:'Speiser AXT 8/300', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00353': {b:'Speiser AXT 10', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00354': {b:'Speiser SDTA 6', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00355': {b:'Speiser SDTA 8', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00356': {b:'Speiser SDTA 9', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00360': {b:'Speiser XTA 11 GT', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00362': {b:'Speiser MD 400', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00402': {b:'Speiser ZPF 3,5/5 K /12QC', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00403': {b:'Speiser ZPF 3,5/5 K /31QC', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00404': {b:'Speiser ZP 3,5/5 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00406': {b:'Speiser ZP 4/5 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00407': {b:'Speiser ZP 4/5 K /31QC', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00408': {b:'Speiser ZP 4/7 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00412': {b:'Speiser ZP 4/7 K /11QC', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00413': {b:'Speiser ZP 4/7 K /31QC', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00414': {b:'Speiser ZP 4/7 KE /31Q', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00415': {b:'Speiser ZP 4/95 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00416': {b:'Speiser ZP 4/95 K /31Q', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00417': {b:'Speiser ZP 4/95 K /11Q', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00418': {b:'Speiser ZP 5/8 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00424': {b:'Speiser ZP 5/8 K /11Q', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00428': {b:'Speiser ZP 5/8 K /31Q', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00434': {b:'Speiser ZP 6/9 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00438': {b:'Speiser ZP 6/9 K /11Q', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00439': {b:'Speiser ZP 6/9 K /31Q', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00444': {b:'Speiser ZP 6/12 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00448': {b:'Speiser ZP 6/12 K /11Q', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00450': {b:'Speiser ZP 6/12 K /31Q', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00454': {b:'Speiser ZP 7/10 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00458': {b:'Speiser ZP 7/10 K /11Q', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00462': {b:'Speiser ZP 7/10 K /31Q', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00468': {b:'Speiser ZP 8/11 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00472': {b:'Speiser ZP 8/11 K /11Q', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00478': {b:'Speiser ZP 9/12 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00480': {b:'Speiser ZP 9/12 K /11Cr', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00482': {b:'Speiser ZP 9/12 K /11Q', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00488': {b:'Speiser ZP 10/13 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00492': {b:'Speiser ZP 10/13 K /11Cr', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00496': {b:'Speiser ZP 10/13 K /11Q', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00504': {b:'Speiser ZP 12/15 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00505': {b:'Speiser ZP 12/15 KL', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00508': {b:'Speiser ZP 12/15 K /11Cr', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00511': {b:'Speiser ZP 12/15 KL /11QC', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00512': {b:'Speiser ZP 12/15 K /11Q', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00513': {b:'Speiser ZP 14/15 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00526': {b:'Speiser SDP 4/5 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00528': {b:'Speiser SDP 4/95 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00531': {b:'Speiser SDP 5/8 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00536': {b:'Speiser SDP 6/9 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00537': {b:'Speiser SDP 6/9 K /11QC', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00538': {b:'Speiser SDP 6/12 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00541': {b:'Speiser SDP 7/10 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00542': {b:'Speiser SDP 7/10 K /11QC', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00551': {b:'Speiser SDP 8/11 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00552': {b:'Speiser SDP 8/11 K /11QC', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00561': {b:'Speiser SDP 9/12 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00562': {b:'Speiser SDP 9/12 K /11QC', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00571': {b:'Speiser SDP 10/13 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00572': {b:'Speiser SDP 10/17 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00574': {b:'Speiser SDP 12/15 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00575': {b:'Speiser SDP 12/15 K /11QC', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00576': {b:'Speiser SDP 14/15 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00578': {b:'Speiser SDP 16/15 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00618': {b:'Speiser OZF 5/8', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00619': {b:'Auflagekern FAKN1 11/250', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00620': {b:'Speiser OZF 5/8 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00624': {b:'Speiser OZF 5/8 K /20QC', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00630': {b:'Speiser OZF 6/9', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00632': {b:'Speiser OZF 6/9 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00633': {b:'Speiser OZF 6/9 K /20Q', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00639': {b:'Speiser OZF 7/10', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00640': {b:'Speiser OZF 7/10 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00646': {b:'Speiser OZF 7/10 K /20Q', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00649': {b:'Speiser OZF 8/11', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00650': {b:'Speiser OZF 8/11 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00652': {b:'Speiser OZF 8/11 K /20Q', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00662': {b:'Speiser OSD 5/8', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00663': {b:'Speiser OSD 5/8 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00665': {b:'Speiser OSD 6/9', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00666': {b:'Speiser OSD 6/9 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00669': {b:'Speiser OSD 7/10', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00670': {b:'Speiser OSD 7/10 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00717': {b:'Speiser ZF 4/95 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00718': {b:'Speiser ZF 4/7 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00719': {b:'Speiser ZF 5/8', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00720': {b:'Speiser ZF 5/8 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00730': {b:'Speiser ZF 6/9', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00731': {b:'Speiser ZF 6/9 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00740': {b:'Speiser ZF 7/10', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00743': {b:'Speiser ZF 7/10 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00749': {b:'Speiser ZF 8/11', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00750': {b:'Speiser ZF 8/11 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00755': {b:'Speiser ZF 9/12', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00760': {b:'Speiser ZF 9/12 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00769': {b:'Speiser ZF 10/13', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00770': {b:'Speiser ZF 10/13 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00773': {b:'Speiser SD 4/95 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00774': {b:'Speiser SD 5/8', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00779': {b:'Speiser SD 8/11', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00780': {b:'Speiser SD 8/11 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00782': {b:'Speiser SD 9/12', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00784': {b:'Speiser SD 10/13 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00790': {b:'Speiser KSP 8/11 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00799': {b:'Speiser ZTAE 18/20', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00800': {b:'Speiser ZTAE 15/18', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00820': {b:'HD1 P  5/ 8KW/31Q', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00830': {b:'HD1 G  6/ 9KW/31HD', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00836': {b:'HD1 P  6/12K/31QH', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00870': {b:'Feedex HD1 V 267/10 QH', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00871': {b:'Feedex HD1 V 36 / 10 / QH', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00872': {b:'Feedex HD1 V56/10/QH', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-00873': {b:'Feedex FEF VS 276', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01030': {b:'Tele Feeder 230-25(40)B3 CB 21/2', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01031': {b:'TELE-FEEDER 40-15 (26) BO', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01032': {b:'TELE-FEEDER 80-17 (28) B 2,5', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01033': {b:'TELE FEEDER 220 - 25 (40)', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01034': {b:'TELE FEEDER 330 - 30 (52)', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01035': {b:'TELE FEEDER 370 - 25 (40)', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01036': {b:'TELE FEEDER 300 - 30 (52)', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01037': {b:'TELE FEEDER CB 21/21 390-25 (40)', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01038': {b:'TELE FEEDER 500-40 (62) B3', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01039': {b:'TELE FEEDER 820-30 (52) B5', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01040': {b:'TELE FEEDER SH 21/21 390-25 (40)', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01102': {b:'Deckel KD 1', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01104': {b:'Deckel KD 2', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01106': {b:'Deckel KD 3', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01108': {b:'Deckel KD 4', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01110': {b:'Deckel KD 5', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01112': {b:'Deckel KD 6', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01114': {b:'Deckel KD 7', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01116': {b:'Deckel KD 8', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01118': {b:'Deckel KD 9', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01120': {b:'Deckel KD 10', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01122': {b:'Deckel KD 11', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01160': {b:'Deckel OKD 4', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01162': {b:'Deckel OKD 5', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01164': {b:'Deckel OKD 6', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01166': {b:'Deckel OKD 7', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01168': {b:'Deckel OKD 8', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01170': {b:'Deckel OKD 9', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01172': {b:'Deckel OKD 10', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01173': {b:'Deckel OKD 11', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01205': {b:'Brechkern BKQ 3,5-5 /11', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01208': {b:'Brechkern BKQ 4/1', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01209': {b:'Brechkern BKQ 4-7/10', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01210': {b:'Brechkern BKQ 4/2', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01212': {b:'Brechkern BKQ 4-7/11', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01218': {b:'Brechkern BKQ 5/1', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01220': {b:'Brechkern BKQ 5/2', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01222': {b:'Brechkern BKQ 5-8/10', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01223': {b:'Brechkern BKQ 5-8/11', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01224': {b:'Brechkern BKQ 5-8/20', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01226': {b:'Brechkern BKQ 5-8/31', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01229': {b:'F-Brechkern Quarzsand BKF 6/1', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01230': {b:'Brechkern BKQ 6/2', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01232': {b:'Brechkern BKQ 6-9/10', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01233': {b:'Brechkern BKQ 6-9/11 (= 6-12/11)', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01234': {b:'Brechkern BKQ 6-9/20', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01235': {b:'Brechkern BKQ 7/2', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01242': {b:'Brechkern BKQ 7-10/11', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01243': {b:'Brechkern BKQ 7-10/10', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01244': {b:'Brechkern BKQ 7-10/20', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01246': {b:'Brechkern BKQ 7-10/31', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01252': {b:'Brechkern BKQ 8-11/11', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01254': {b:'F-Brechkern Quarzsand BKF 8-11/2', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01256': {b:'Brechkern BKQ 8-11/31', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01259': {b:'Brechkern BKQ 9-12/10 (= 3/1)', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01262': {b:'F-Brechkern Quarzsand BKF 9-12/1', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01274': {b:'Brechkern BKQ 10-13/10', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01275': {b:'Brechkern BKQ 10-13/11', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01276': {b:'Brechkern BKQ 10-13/31', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01280': {b:'F-Brechkern Quarzsand BKF 12-15/', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01281': {b:'Brechkern BKQ 12-15/31', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01313': {b:'Brechkern BKCr 4/1', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01315': {b:'Brechkern BKCr 4/2', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01318': {b:'Brechkern BKCr 5/1', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01320': {b:'Brechkern BKCr 5/2', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01321': {b:'Brechkern BKCr 6-9/20', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01325': {b:'Brechkern BKCr 6/2', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01326': {b:'Brechkern BKCr 5-8/20', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01330': {b:'Brechkern BKCr 7/2', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01334': {b:'Brechkern BKCr 8-11/11', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01335': {b:'Brechkern BKCr 8/2', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01338': {b:'Brechkern BKCr 9/1', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01340': {b:'Brechkern BKCr 9/2', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01344': {b:'Brechkern BKCr 10/1', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01346': {b:'Brechkern BKCr 9-12/10 (= 3/1)', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01348': {b:'Brechkern BKCr 12-15/11', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01349': {b:'Brechkern BKCr 12/1 (SHB)', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01351': {b:'Brechkern BKCr 11/1 (SHB)', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01355': {b:'Brechkern BKCr 14/1', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01356': {b:'Brechkern BKCr 15/1', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01360': {b:'Brechkern BKCr 16-20/2', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01370': {b:'Brechkern BKCr 20-25/2', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01438': {b:'BKN 3  0207  (200 mm)', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01439': {b:'BKN 3  0197  (300 mm)', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01611': {b:'Auflagekern FAKN2 5/90', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01613': {b:'Auflagekern FAKN1 6/125', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01614': {b:'Auflagekern FSCP 9/200', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01615': {b:'Auflagekern FAKN1 7/150', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01616': {b:'Auflagekern FSCN1 9/200', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01617': {b:'Auflagekern FSCN1 10/250', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01618': {b:'Auflagekern FAKN1 11/200', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01626': {b:'Auflagekern FSCP 60/150', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01628': {b:'Auflagekern FSCP 10/250', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-01700': {b:'Feedex HD1R-Brechkerne BKHD1R 03', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02002': {b:'Kalpad-Platte  1001', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02003': {b:'Kalpad-Platte 1002', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02103': {b:'Speiser FSD 3', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02105': {b:'Speiser FSD 4', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02107': {b:'Speiser FSD 5', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02111': {b:'Speiser FSD 6', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02116': {b:'Speiser FSD 7', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02119': {b:'Speiser FSD 8', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02122': {b:'Speiser FSD 9', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02124': {b:'Speiser FSD 10', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02222': {b:'Speiser FOSD 4', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02223': {b:'Speiser FOSD 4 / 75', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02225': {b:'Speiser FOSD 5', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02226': {b:'Speiser FOSD 6', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02231': {b:'Speiser FOSD 8/11 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02234': {b:'Speiser FOSD 8', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02334': {b:'Speiser FSDTA 6', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02340': {b:'Speiser FSDTA 8', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02344': {b:'Speiser FSDTA 9', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02438': {b:'Speiser FSDP 6/9 K /11 QC', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02458': {b:'Speiser FSDP 7/10 K /11 QC', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02472': {b:'Speiser FSDP 8/11 K /11 QC', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02482': {b:'Speiser FSDP 9/12 K /11 QC', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02505': {b:'Speiser FSDP 12/15 KL', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02511': {b:'Speiser FSDP 12/15 KL /11 Q', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02512': {b:'Speiser FSDP 12/15 K /11 QC', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02513': {b:'Speiser FSDP 14/15 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02528': {b:'Speiser FSDP 4/95 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02531': {b:'Speiser FSDP 5/8 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02536': {b:'Speiser FSDP 6/9 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02538': {b:'Speiser FSDP 6/12 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02541': {b:'Speiser FSDP 7/10 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02551': {b:'Speiser FSDP 8/11  K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02561': {b:'Speiser FSDP 9/12/K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02571': {b:'Speiser FSDP 10/13 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02572': {b:'Speiser FSDP 10/17 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02573': {b:'Speiser FSDP 12/15 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02575': {b:'Speiser FSDP 16/15 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02630': {b:'Speiser FOSD 6/9', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02632': {b:'Speiser FOSD 6/9 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02663': {b:'Speiser FOSD 5/8 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02669': {b:'Speiser FOSD 7/10', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02670': {b:'Speiser FOSD 7/10 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02719': {b:'Speiser FSD 5/8', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02740': {b:'Speiser FSD 7/10', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02773': {b:'Speiser FSD 4/95 K', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-02950': {b:'Speiser FSD 8/11 K / 0376 HDC', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-85005': {b:'Speiser K-Lag. > 6 Monate', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4308-95005': {b:'Speiser K-Lager / Foseco', w:'exotherme Kappen (Speiser)', g:'Hilfs- und Betriebsstoffe'},
  '4309-00117': {b:'Filter STELEX ZR 70R x 25/10ppi', w:'Gießfilter', g:'Hilfs- und Betriebsstoffe'},
  '4309-00120': {b:'Filter STELEX ZR 90R x 25/10ppi', w:'Gießfilter', g:'Hilfs- und Betriebsstoffe'},
  '4309-00121': {b:'Filter STELEX ZR 100R x 25/10ppi', w:'Gießfilter', g:'Hilfs- und Betriebsstoffe'},
  '4309-00122': {b:'Filter STELEX ZR 175R x 35/10ppi', w:'Gießfilter', g:'Hilfs- und Betriebsstoffe'},
  '4309-00123': {b:'Filter STELEX ZR 125R x 30/10ppi', w:'Gießfilter', g:'Hilfs- und Betriebsstoffe'},
  '4309-00124': {b:'Filter STELEX ZR 200R x 30/10ppi', w:'Gießfilter', g:'Hilfs- und Betriebsstoffe'},
  '4309-00125': {b:'Filter STELEX ZR 150R x 30/10ppi', w:'Gießfilter', g:'Hilfs- und Betriebsstoffe'},
  '4309-00128': {b:'Filter STELEX ZR 150x150x30/10pp', w:'Gießfilter', g:'Hilfs- und Betriebsstoffe'},
  '4309-00129': {b:'Filter STELEX ZR 125x125x30/10pp', w:'Gießfilter', g:'Hilfs- und Betriebsstoffe'},
  '4309-00131': {b:'Filter STELEX ZR 100x100x25/10pp', w:'Gießfilter', g:'Hilfs- und Betriebsstoffe'},
  '4309-00133': {b:'Filter STELEX ZR 75x75x25/10ppi', w:'Gießfilter', g:'Hilfs- und Betriebsstoffe'},
  '4309-00134': {b:'Filter STELEX ZR 55x55x25/10ppi', w:'Gießfilter', g:'Hilfs- und Betriebsstoffe'},
  '4309-00140': {b:'Filter SEDEX* SIC 50x50x22/10ppi', w:'Gießfilter', g:'Hilfs- und Betriebsstoffe'},
  '4309-00148': {b:'Filter STELEX PrO 90R x 25/10ppi', w:'Gießfilter', g:'Hilfs- und Betriebsstoffe'},
  '4309-00150': {b:'Filter STELEX PrO 100R x 25/10pp', w:'Gießfilter', g:'Hilfs- und Betriebsstoffe'},
  '4309-00152': {b:'Filter STELEX PrO 125R x 30/10pp', w:'Gießfilter', g:'Hilfs- und Betriebsstoffe'},
  '4309-00153': {b:'Filter STELEX PrO 150R x 30/10pp', w:'Gießfilter', g:'Hilfs- und Betriebsstoffe'},
  '4309-00154': {b:'Filter STELEX PrO 175R x 35/10pp', w:'Gießfilter', g:'Hilfs- und Betriebsstoffe'},
  '4309-00155': {b:'Filter STELEX PrO 200R x 35/10pp', w:'Gießfilter', g:'Hilfs- und Betriebsstoffe'},
  '4309-00156': {b:'Filter STELEX PrO 150x150x30/10p', w:'Gießfilter', g:'Hilfs- und Betriebsstoffe'},
  '4309-00157': {b:'Filter STELEX PrO 250R x 40/10pp', w:'Gießfilter', g:'Hilfs- und Betriebsstoffe'},
  '4309-00158': {b:'Filter STELEX PrO 125x125x30/10p', w:'Gießfilter', g:'Hilfs- und Betriebsstoffe'},
  '4309-00160': {b:'Filter STELEX PrO 100x100x25/10p', w:'Gießfilter', g:'Hilfs- und Betriebsstoffe'},
  '4309-00162': {b:'Filter STELEX PrO 55x55x25/10ppi', w:'Gießfilter', g:'Hilfs- und Betriebsstoffe'},
  '4309-00163': {b:'Filter STELEX PrO 75x75x25/10ppi', w:'Gießfilter', g:'Hilfs- und Betriebsstoffe'},
  '4309-00181': {b:'Filter STELEX PrO 150x200x30/10p', w:'Gießfilter', g:'Hilfs- und Betriebsstoffe'},
  '4309-00182': {b:'Filter STELEX PrO 190x90x30/10pp', w:'Gießfilter', g:'Hilfs- und Betriebsstoffe'},
  '4309-00250': {b:'Filter SEDEX 50x50x22/10ppi', w:'Gießfilter', g:'Hilfs- und Betriebsstoffe'},
  '4309-00251': {b:'Filter SEDEX 75x75x22/10ppi', w:'Gießfilter', g:'Hilfs- und Betriebsstoffe'},
  '4309-00252': {b:'Filter SEDEX 100x100x22/10ppi', w:'Gießfilter', g:'Hilfs- und Betriebsstoffe'},
  '4309-00253': {b:'Filter SEDEX 60x60x22/10ppi', w:'Gießfilter', g:'Hilfs- und Betriebsstoffe'},
  '4309-00300': {b:'Filter SEDEX 100x150x22/10ppi', w:'Gießfilter', g:'Hilfs- und Betriebsstoffe'},
  '4309-00400': {b:'FCF - 1 Ø 70 - 60 x 25 10 PPI', w:'Gießfilter', g:'Hilfs- und Betriebsstoffe'},
  '4309-00732': {b:'Filter f. Kreisel 6/4Q  (4x4x1,2', w:'Gießfilter', g:'Hilfs- und Betriebsstoffe'},
  '4310-00010': {b:'Williamskern WKQ 0350', w:'Kerne (bez. Kerne f. Gießerei)', g:'Hilfs- und Betriebsstoffe'},
  '4310-00015': {b:'Williamskern WKQ 0349', w:'Kerne (bez. Kerne f. Gießerei)', g:'Hilfs- und Betriebsstoffe'},
  '4310-00018': {b:'Abdeckkern  K 3', w:'Kerne (bez. Kerne f. Gießerei)', g:'Hilfs- und Betriebsstoffe'},
  '4310-00019': {b:'Abdeckkern K4', w:'Kerne (bez. Kerne f. Gießerei)', g:'Hilfs- und Betriebsstoffe'},
  '4310-00368': {b:'Innenkern, li. (1) / Bod.-pl. 12', w:'Kerne (bez. Kerne f. Gießerei)', g:'Hilfs- und Betriebsstoffe'},
  '4310-00369': {b:'Innenkern, re. (2) / Bod.-pl. 12', w:'Kerne (bez. Kerne f. Gießerei)', g:'Hilfs- und Betriebsstoffe'},
  '4310-00416': {b:'Kern KG 1000 K2+K3 teilgeschl.', w:'Kerne (bez. Kerne f. Gießerei)', g:'Hilfs- und Betriebsstoffe'},
  '4310-00417': {b:'Kern KG 1000 K4+2xK5 ungeschl.', w:'Kerne (bez. Kerne f. Gießerei)', g:'Hilfs- und Betriebsstoffe'},
  '4310-00418': {b:'Kern KG 1000 K1', w:'Kerne (bez. Kerne f. Gießerei)', g:'Hilfs- und Betriebsstoffe'},
  '4310-01010': {b:'Bod.-pl. 1200/230 Kom. / K 3.1+2', w:'Kerne (bez. Kerne f. Gießerei)', g:'Hilfs- und Betriebsstoffe'},
  '4310-01011': {b:'Bod.-pl. 1200/230 Kom. / K 5.1+2', w:'Kerne (bez. Kerne f. Gießerei)', g:'Hilfs- und Betriebsstoffe'},
  '4310-01012': {b:'Bod.-pl. 1200/230 Kom. / K 6.1+2', w:'Kerne (bez. Kerne f. Gießerei)', g:'Hilfs- und Betriebsstoffe'},
  '4310-01013': {b:'Bod.-pl. 1200/230 Kom. / K 10', w:'Kerne (bez. Kerne f. Gießerei)', g:'Hilfs- und Betriebsstoffe'},
  '4310-01014': {b:'Bod.-pl. 1200/230 Kom. / K 7', w:'Kerne (bez. Kerne f. Gießerei)', g:'Hilfs- und Betriebsstoffe'},
  '4310-01030': {b:'Bod.-pl. 1500/320 Kom. / K 2u+o', w:'Kerne (bez. Kerne f. Gießerei)', g:'Hilfs- und Betriebsstoffe'},
  '4310-01031': {b:'Bod.-pl. 1500/320 Kom. / K 4u+o', w:'Kerne (bez. Kerne f. Gießerei)', g:'Hilfs- und Betriebsstoffe'},
  '4310-01032': {b:'Bod.-pl. 1500/320 Kom. / K 6u+o', w:'Kerne (bez. Kerne f. Gießerei)', g:'Hilfs- und Betriebsstoffe'},
  '4310-01033': {b:'Bod.-pl. 1500/320 Kom. / K 2au+o', w:'Kerne (bez. Kerne f. Gießerei)', g:'Hilfs- und Betriebsstoffe'},
  '4310-01148': {b:'Brechkern 1 (A) / Bod.-pl. 1200', w:'Kerne (bez. Kerne f. Gießerei)', g:'Hilfs- und Betriebsstoffe'},
  '4310-01149': {b:'Brechkern 2 (B) / Bod.-pl. 1200', w:'Kerne (bez. Kerne f. Gießerei)', g:'Hilfs- und Betriebsstoffe'},
  '4310-01150': {b:'Brechkern 3 (C) / Bod.-pl. 1200', w:'Kerne (bez. Kerne f. Gießerei)', g:'Hilfs- und Betriebsstoffe'},
  '4310-01151': {b:'Brechkern 4 (D) / Bod.-pl. 1200', w:'Kerne (bez. Kerne f. Gießerei)', g:'Hilfs- und Betriebsstoffe'},
  '4310-01152': {b:'Brechkern 5 (F) / Bod.-pl. 1200', w:'Kerne (bez. Kerne f. Gießerei)', g:'Hilfs- und Betriebsstoffe'},
  '4311-00001': {b:'GERMALLOY K  110', w:'metallurgische Zusatzstoffe', g:'Hilfs- und Betriebsstoffe'},
  '4311-00002': {b:'GERMALLOY K 150', w:'metallurgische Zusatzstoffe', g:'Hilfs- und Betriebsstoffe'},
  '4311-00003': {b:'GERMALLOY K 200', w:'metallurgische Zusatzstoffe', g:'Hilfs- und Betriebsstoffe'},
  '4311-00004': {b:'GERMALLOY  K 110', w:'metallurgische Zusatzstoffe', g:'Hilfs- und Betriebsstoffe'},
  '4311-00005': {b:'GERMALLOY K 150', w:'metallurgische Zusatzstoffe', g:'Hilfs- und Betriebsstoffe'},
  '4311-00006': {b:'GERMALLOY  K 200', w:'metallurgische Zusatzstoffe', g:'Hilfs- und Betriebsstoffe'},
  '4311-00009': {b:'GERMALLOY P 300', w:'metallurgische Zusatzstoffe', g:'Hilfs- und Betriebsstoffe'},
  '4311-00010': {b:'GERMALLOY P 500', w:'metallurgische Zusatzstoffe', g:'Hilfs- und Betriebsstoffe'},
  '4312-00020': {b:'Dolomit', w:'Hilfsstoffe Gießerei', g:'Hilfs- und Betriebsstoffe'},
  '4312-00030': {b:'Weißstückkalk 6 - 30 mm in Big B', w:'Hilfsstoffe Gießerei', g:'Hilfs- und Betriebsstoffe'},
  '4312-00031': {b:'Rohkalk / Steinkalk, lose', w:'Hilfsstoffe Gießerei', g:'Hilfs- und Betriebsstoffe'},
  '4312-00032': {b:'Weißstückkalk 2 - 6 mm  in Big B', w:'Hilfsstoffe Gießerei', g:'Hilfs- und Betriebsstoffe'},
  '4312-00040': {b:'Flußspat', w:'Hilfsstoffe Gießerei', g:'Hilfs- und Betriebsstoffe'},
  '4312-00050': {b:'Aluminium Granulat mind. 97 % Al', w:'Hilfsstoffe Gießerei', g:'Hilfs- und Betriebsstoffe'},
  '4312-00115': {b:'Ferrux 740', w:'Hilfsstoffe Gießerei', g:'Hilfs- und Betriebsstoffe'},
  '4312-00118': {b:'Glutin T / 10kg Beutel', w:'Hilfsstoffe Gießerei', g:'Hilfs- und Betriebsstoffe'},
  '4312-00200': {b:'Sauerstoffrohr 1/2",  schwarz', w:'Hilfsstoffe Gießerei', g:'Hilfs- und Betriebsstoffe'},
  '4312-00300': {b:'Magnesium-Behandlungsdraht W Mg', w:'Hilfsstoffe Gießerei', g:'Hilfs- und Betriebsstoffe'},
  '4312-00301': {b:'Ni-MG VL 4 (M) 800 g', w:'Hilfsstoffe Gießerei', g:'Hilfs- und Betriebsstoffe'},
  '4312-00350': {b:'Impflegierung SMZ25 0,7-2 mm', w:'Hilfsstoffe Gießerei', g:'Hilfs- und Betriebsstoffe'},
  '4312-00351': {b:'Impfmittel GJL', w:'Hilfsstoffe Gießerei', g:'Hilfs- und Betriebsstoffe'},
  '4312-00352': {b:'Pyrit -FeS2', w:'Hilfsstoffe Gießerei', g:'Hilfs- und Betriebsstoffe'},
  '4312-00500': {b:'Meßspitzen PT 18 / 400 mm / ohne', w:'Hilfsstoffe Gießerei', g:'Hilfs- und Betriebsstoffe'},
  '4312-00502': {b:'Tauchmesskopf Typ TC36032WE', w:'Hilfsstoffe Gießerei', g:'Hilfs- und Betriebsstoffe'},
  '4312-00503': {b:'CeloxAL Typ CE36911290CT, Pt18%', w:'Hilfsstoffe Gießerei', g:'Hilfs- und Betriebsstoffe'},
  '4313-00050': {b:'Ofenstein Rubinal MNH NF 2 / M 8', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00051': {b:'Ofenstein Radex E12 M88 NF 2 / C', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00052': {b:'Ofenstein Rubinal MNH 2 Q 28 /', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00054': {b:'Ofenstein Radex E12 M 88  2 Q 28', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00056': {b:'Ofenstein  2 H 6 / CM 23', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00057': {b:'Ofenstein  2 H 16  / CM 23', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00059': {b:'Ofenstein  2 H 26  / CM 23', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00062': {b:'Türpfeiler-Eckstein  V2L-65', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00064': {b:'Abstichstein Radex E 12  /  60/0', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00066': {b:'Ofenstein 2-32 /30-er Plättchen', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00070': {b:'K-M88  SL 564 NF 2 / Dauerfutter', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00071': {b:'K-M88  NF 1 / Dauerfutter', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00100': {b:'SYNCARBON C  F3T14X 65/20T20A', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00220': {b:'Deckelstein         KR 30     BX', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00222': {b:'Deckelstein         R 30      BX', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00224': {b:'Deckelstein         2 H 10', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00226': {b:'Deckelstein         2 H 38', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00227': {b:'Spülerhülse H12/K16 SHB', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00446': {b:'Pfannenstein  NF 2  /  Schamotte', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00448': {b:'Pfannenstein  1 K 11  /  Schamot', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00449': {b:'Pfannenstein  1 K 40  /  Schamot', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00450': {b:'Pfannenstein  B 1  /  S 36 C', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00455': {b:'Pfannenstein  2 P 24  /  B 85 C', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00457': {b:'Pfannenstein  NF 2  /  B 85 C', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00459': {b:'Pfannenstein  1 P 18  /  B 85 C', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00460': {b:'Schamotteplättchen  30 mm', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00461': {b:'Schamotteplättchen  40-er', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00464': {b:'Lochstein  L 1      S 63 C', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00465': {b:'Schamotteplatten A 30t 400 x 200', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00466': {b:'Pfannenstein  1 P 26  /  B 85 C', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00469': {b:'Pfannenstein  0 P 96/80  /  B 85', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00470': {b:'Pfannenstein  0 P 7   B 85 C', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00472': {b:'Pfannenstein  4 P 22/54  Urex B', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00473': {b:'Pfannenstein  4 P 22/58  Urex 85', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00608': {b:'Resimur 70/90 0-0,7   a 1200 kg', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00623': {b:'Anker Indux MB 14  0-6, für MF-O', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00625': {b:'Feuerbeton   GA 70', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00626': {b:'Reparaturmasse   RA 70', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00635': {b:'Schamotte-Mörtel   S 36', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00640': {b:'Diram A86 CRP-3 (eh. Shamrock)', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00650': {b:'Perramit  MA 20   /  in 1,25-t-B', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00655': {b:'Ankerjet  NW 22     1,2-t-Pal. (', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00658': {b:'Comprit B80 G-30', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00659': {b:'Refracast LC C-84 CR-7 SP', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00660': {b:'Refraselfcast LC C-84 CR-7 SP', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00701': {b:'Ankerfix NS 60 / Magn.-Mörtel, c', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00703': {b:'Diram B88P-6  /  1,2 t/Pal.', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00705': {b:'Ankoflo Max - 15', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00706': {b:'Ankoflo Max - 6 / Gießmasse [f.St](http://f.St)', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00707': {b:'Ankermix NS13 in 25 kg-Säcken, 1', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00709': {b:'Hinterfüllmasse SM 90/0-1, 90 %', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00853': {b:'Stopfenstangenrohre  SR 2/120', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00854': {b:'Stopfenstangenrohre  SR 3/330', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00855': {b:'Stopfenstangenrohre  SR 2/330', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00856': {b:'Ringe aus Keramikfaser, Dmr. 120', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00857': {b:'Stopfenstangenrohre  SR 3/100', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00858': {b:'Stopfenstangenrohr SR3/330 B70', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00910': {b:'Ausgüsse  1 A 40', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00915': {b:'Ausgüsse  1 A 50', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00917': {b:'Ausgüsse (Kreuz)  1 A 50 K', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00920': {b:'Ausgüsse  1 A 60', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00930': {b:'Ausgüsse  1 A 70', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00940': {b:'Ausgüsse  1 A 80', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-00950': {b:'Stopfen Tundish Nozzle', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4313-97910': {b:'Feuerfest K-Lager / Krause', w:'Feuerfestmaterial Ofen/Pfannen', g:'Hilfs- und Betriebsstoffe'},
  '4314-00010': {b:'Aufkohlungsmittel / fein  0,5-3m', w:'Kohleprodukte (Grafitelektr., Aufk', g:'Hilfs- und Betriebsstoffe'},
  '4314-00020': {b:'Ranco 9904 Elektrodengrafit 0,2', w:'Kohleprodukte (Grafitelektr., Aufk', g:'Hilfs- und Betriebsstoffe'},
  '4314-00021': {b:'Aufkohlungsmittel GJL', w:'Kohleprodukte (Grafitelektr., Aufk', g:'Hilfs- und Betriebsstoffe'},
  '4314-00050': {b:'Graphitelektroden / 250 x 1800', w:'Kohleprodukte (Grafitelektr., Aufk', g:'Hilfs- und Betriebsstoffe'},
  '4314-00108': {b:'Rotolok-Stopfen, get. Einsatz /', w:'Kohleprodukte (Grafitelektr., Aufk', g:'Hilfs- und Betriebsstoffe'},
  '4314-00205': {b:'Monoblockstopfen  BP 00755 / für', w:'Kohleprodukte (Grafitelektr., Aufk', g:'Hilfs- und Betriebsstoffe'},
  '4318-00206': {b:'Gewindestange M10x725 +/-1,0 mm', w:'Befestigungsteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00001': {b:'ZULAUFROHR U 105X500 RBPK', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00002': {b:'EINLAUFSTÜCK U105x500 RBPK', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00003': {b:'DICHTUNG FÜR ZULAUFROHR U105x500', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00004': {b:'VERTEILER U105X500 RBPK', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00005': {b:'MITNEHMERSCHEIBE U105x500 RBPK', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00006': {b:'KUPPLUNGSSTÜCK U105X500 RBPK', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00007': {b:'WURFSCHAUFEL U105x500 RBPK', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00008': {b:'STIRNSCHUTZPLATTE  U105', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00009': {b:'DECKENSCHUTZPLATTE U105', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00041': {b:'ZUTEILROHR SR', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00042': {b:'SCHLEUDERSCHAUFEL SR', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00043': {b:'ZUTEILRAD MIT SCHEIBE', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00044': {b:'ZUTEILHÜLSE SR', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00045': {b:'KOPFSCHLEIßPLATTE  SR', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00046': {b:'DECKELSCHLEIßPLATTE SR', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00047': {b:'LABYRINTHRING AUßEN  SR', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00048': {b:'LABYRINTHRING INNEN  SR', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00050': {b:'GEHÄUSESCHUTZKLOBEN', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00051': {b:'DECKELSCHUTZKLOBEN', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00052': {b:'WURFSCHAUFELN TSM IV aus Werkzeu', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00053': {b:'BEFESTIGUNGSFEDER', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00054': {b:'VERTEILERRAD KRS 477/42.2.077d', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00055': {b:'VERTEILERHÜLSE 30 mm breit KRS 5', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00056': {b:'SCHUTZSCHEIBE TSM IV', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00057': {b:'STRAHLMITTELZULAUF', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00074': {b:'Schneiddüse K6 Messing 10 bar, 2', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00075': {b:'HEIZDÜSE A2 DMR 28,0 mm f. 100-3', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00392': {b:'1-Stranghakenkette H1 10x1024 m.', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00393': {b:'1-Stranghakenkette H1 10x574 m.', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00394': {b:'1 Stranghakenkette H1 6x1000 m.', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00395': {b:'1-Stranghakenkette H1 10x500 m.Ö', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00396': {b:'1-Stranghakenkette H1 8 x 1500 m', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00397': {b:'1-Stranghakenkette H1-10x2000 m.', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00398': {b:'1-Stranghakenkette H1-10x1500 m.', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00399': {b:'1-Stranghakenkette H1-6x1500 m.', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00400': {b:'1-Stranghakenkette H1-6x1000 m.Ö', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00401': {b:'1-Stranghakenkette H1 10x1000 m.', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00402': {b:'1-Stranghakenkette H1 6x2000 m.Ö', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00600': {b:'Schleuderschaufel', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00601': {b:'Schenkelfeder', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00602': {b:'Zuteilrad', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00603': {b:'Zuteilhülse', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00605': {b:'Schutzstück links mit Überlappun', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00606': {b:'Schutzstück rechts ohne Überlapp', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00607': {b:'Zuteilrohr B-BR11-218', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4320-00608': {b:'Mitnehmerscheibe', w:'Ersatzteile', g:'Hilfs- und Betriebsstoffe'},
  '4321-00005': {b:'Kleinschleifkörper 52ZY-75x10x08', w:'Schleif- und Fräsmittel', g:'Hilfs- und Betriebsstoffe'},
  '4321-00006': {b:'Kleinschleifkörper 18-75x35xM12x', w:'Schleif- und Fräsmittel', g:'Hilfs- und Betriebsstoffe'},
  '4321-00052': {b:'Trennscheibe 115x3x22,2', w:'Schleif- und Fräsmittel', g:'Hilfs- und Betriebsstoffe'},
  '4321-00055': {b:'SCHRUPPSCHEIBE 115x6,5x22,2', w:'Schleif- und Fräsmittel', g:'Hilfs- und Betriebsstoffe'},
  '4321-00060': {b:'Schleiflamellenteller SLTR 125 K', w:'Schleif- und Fräsmittel', g:'Hilfs- und Betriebsstoffe'},
  '4321-00066': {b:'TRENNSCHEIBE 125x3x22,2', w:'Schleif- und Fräsmittel', g:'Hilfs- und Betriebsstoffe'},
  '4321-00068': {b:'SCHRUPPSCHEIBE 125X7x22,2 ZA30R-', w:'Schleif- und Fräsmittel', g:'Hilfs- und Betriebsstoffe'},
  '4321-00074': {b:'Schruppscheibe 178X7X22,2', w:'Schleif- und Fräsmittel', g:'Hilfs- und Betriebsstoffe'},
  '4321-00075': {b:'TRENNSCHEIBE 180X3X22,2', w:'Schleif- und Fräsmittel', g:'Hilfs- und Betriebsstoffe'},
  '4321-00084': {b:'SCHRUPPSCHEIBE 180X6X22,2', w:'Schleif- und Fräsmittel', g:'Hilfs- und Betriebsstoffe'},
  '4321-00100': {b:'SCHRUPPSCHEIBE 230X8x22,2', w:'Schleif- und Fräsmittel', g:'Hilfs- und Betriebsstoffe'},
  '4321-00101': {b:'Trennscheibe 230x3x22,2', w:'Schleif- und Fräsmittel', g:'Hilfs- und Betriebsstoffe'},
  '4321-00138': {b:'SCHLEIFSCHEIBE 350x60x127', w:'Schleif- und Fräsmittel', g:'Hilfs- und Betriebsstoffe'},
  '4321-00155': {b:'SCHLEIFSCHEIBE 400x63x127 63 m/s', w:'Schleif- und Fräsmittel', g:'Hilfs- und Betriebsstoffe'},
  '4321-00156': {b:'Schleifscheibe 400x63x127 80m/s', w:'Schleif- und Fräsmittel', g:'Hilfs- und Betriebsstoffe'},
  '4321-00211': {b:'Trennscheibe 800x9x100 mm A20-BF', w:'Schleif- und Fräsmittel', g:'Hilfs- und Betriebsstoffe'},
  '4321-00215': {b:'SCHLEIFSTIFT 620 16/6x45.06 NK 2', w:'Schleif- und Fräsmittel', g:'Hilfs- und Betriebsstoffe'},
  '4321-00216': {b:'SCHLEIFSTIFT 620 16/6x45 S6x40 H', w:'Schleif- und Fräsmittel', g:'Hilfs- und Betriebsstoffe'},
  '4321-00220': {b:'SCHLEIFSTIFT ZY2040.06 HK30NV13', w:'Schleif- und Fräsmittel', g:'Hilfs- und Betriebsstoffe'},
  '4321-00221': {b:'Schleifstift 520 20x40 S6x40 HK', w:'Schleif- und Fräsmittel', g:'Hilfs- und Betriebsstoffe'},
  '4321-00230': {b:'SCHLEIFSTIFT 510 32x50 S8x40 NK', w:'Schleif- und Fräsmittel', g:'Hilfs- und Betriebsstoffe'},
  '4321-00231': {b:'SCHLEIFSTIFTE 32x50 Form 620 S8x', w:'Schleif- und Fräsmittel', g:'Hilfs- und Betriebsstoffe'},
  '4321-00233': {b:'Schleifstift KE 3550 8 CU 24 R5V', w:'Schleif- und Fräsmittel', g:'Hilfs- und Betriebsstoffe'},
  '4321-00234': {b:'WKG 42/26x63 GEW.B. M10 NSB 20 O', w:'Schleif- und Fräsmittel', g:'Hilfs- und Betriebsstoffe'},
  '4321-00315': {b:'Fräser HFF 0820.06 BASE-X ZX (sp', w:'Schleif- und Fräsmittel', g:'Hilfs- und Betriebsstoffe'},
  '4321-00346': {b:'Fräser HFF 1225.06 (spitz, groß)', w:'Schleif- und Fräsmittel', g:'Hilfs- und Betriebsstoffe'},
  '4321-00350': {b:'FRÄSER HFC 0820.06 BASE-X ZX (ru', w:'Schleif- und Fräsmittel', g:'Hilfs- und Betriebsstoffe'},
  '4321-00355': {b:'FRÄSER WRC 1225 06 S Z 3 + (rund', w:'Schleif- und Fräsmittel', g:'Hilfs- und Betriebsstoffe'},
  '4321-01326': {b:'BKCR 5-8/20', w:'Schleif- und Fräsmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00001': {b:'HALBSCHUH S3   GR.39', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00002': {b:'HALBSCHUH S3  GR. 40', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00003': {b:'HALBSCHUH S3  GR.41', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00004': {b:'HALBSCHUH S3  GR.42', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00005': {b:'HALBSCHUH S3  GR.43', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00006': {b:'HALBSCHUH S3  GR.44', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00007': {b:'HALBSCHUH S3  GR.45', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00008': {b:'HALBSCHUH S3  GR.46', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00009': {b:'HALBSCHUH S3  GR.47', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00010': {b:'Hautschutz Multi Tec vor der Arb', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00011': {b:'Hautreinigungslotion Ivraxo Soft', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00012': {b:'Hautreiniger Ivraxo Soft U', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00013': {b:'Hautpflegecreme Greven Creme C,', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00050': {b:'SCHNÜRSCHUH S3 GR.40', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00051': {b:'SCHNÜRSCHUH S3 GR.41', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00052': {b:'SCHNÜRSCHUH S3 GR.42', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00053': {b:'SCHNÜRSCHUH S3 GR.43', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00054': {b:'SCHNÜRSCHUH S3 GR.44', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00055': {b:'SCHNÜRSCHUH S3 GR.45', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00056': {b:'SCHNÜRSCHUH S3 GR.46', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00057': {b:'SCHNÜRSCHUHE S3 GR.47', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00080': {b:'GIEßERSCHUH/SCHWEISSERSTIEFEL S3', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00081': {b:'GIEßERSCHUH/SCHWEISSERSTIEFEL S3', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00084': {b:'GIEßERSCHUH/SCHWEISSERSTIEFEL S3', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00085': {b:'GIEßERSCHUH/SCHWEISSERSTIEFEL S3', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00086': {b:'GIEßERSCHUH/SCHWEISSERSTIEFEL S3', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00087': {b:'GIEßERSCHUH/SCHWEISSERSTIEFEL S3', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00088': {b:'GIEßERSCHUH/SCHWEISSERSTIEFEL S3', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00089': {b:'GIEßERSCHUH/SCHWEISSERSTIEFEL S3', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00090': {b:'GIEßERSCHUH/SCHWEISSERSTIEFEL S3', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00095': {b:'SCHWEIßERGAMASCHE AUS SPALTLEDER', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00096': {b:'Armschoner Baumwolle', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00097': {b:'Anstoßkappe', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00100': {b:'5-FINGERHANDSCHUH KURZ', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00101': {b:'5-FINGERHANDSCHUH 35 cm lang', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00105': {b:'3-FINGERHANDSCHUH 35 cm lang', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00120': {b:'Latexhandschuhe auf Baumwolle m.', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00130': {b:'Nitril-Montagehandschuhe grau, G', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00131': {b:'Nitril-Montagehandschuhe grau, G', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00132': {b:'Nitril-Montagehandschuhe grau Gr', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00133': {b:'Nitril-Montagehandschuhe grau Gr', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00134': {b:'Nitril-Montagehandschuhe grau Gr', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00141': {b:'HITZEFAUSTHANDSCHUH  SILBER  400', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00142': {b:'5-Finger-Hitzeschutzhandschuhe', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00158': {b:'Feinstaubmaske 4255', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00159': {b:'FEINSTAUBMASKE 9332 FFP3', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00160': {b:'FEINSTAUBMASKE 9322 FFP2', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00161': {b:'FEINSTAUBMASKE 8810', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00163': {b:'Adflo 3M Partikelfilter 20-tlg.', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00164': {b:'Adflo 3M Vorfilter 5-tlg.', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00180': {b:'Lederschürze 80 x 100', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00182': {b:'Spaltlederhüftschurz 60 x 70', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00183': {b:'SCHWEIßERSCHÜRZE 100x30, Chromna', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00184': {b:'Brustlatzschürze', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00185': {b:'Gießermantel Tempex Treme Heat A', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00186': {b:'Tempex-Hitzeschutz-Gamasche', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00188': {b:'Gießermantel Tempex Treme Heat A', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00189': {b:'Komfort-Schutzbrille uvex pheos', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00190': {b:'Gießermantel Tempex Treme Heat A', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00191': {b:'Gießermantel Tempex Treme Heat A', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00201': {b:'Schutzbrille Terminator schwarz', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00202': {b:'BOCHUMER BRILLE KLEIN 200x135 mm', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00206': {b:'Schweißerbrille XC grün', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00209': {b:'Ersatzscheibe BIONIC', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00216': {b:'OFENSCHAUGLAS BLAU DMR.50 STUFE', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00220': {b:'Helm Bionic', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00225': {b:'Gehörschutz Nachfüllpackung 303L', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00226': {b:'Gehörschutz mit Bügel', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4322-00245': {b:'RUNDUMVISIER', w:'Arbeitsschutzmittel', g:'Hilfs- und Betriebsstoffe'},
  '4323-00001': {b:'SCHWEIßKOHLEN 6 mm', w:'Schweißbedarf', g:'Hilfs- und Betriebsstoffe'},
  '4323-00002': {b:'SCHWEIßKOHLEN 8 mm', w:'Schweißbedarf', g:'Hilfs- und Betriebsstoffe'},
  '4323-00003': {b:'SCHWEIßKOHLEN 10 mm', w:'Schweißbedarf', g:'Hilfs- und Betriebsstoffe'},
  '4323-00004': {b:'Kohleelektrode 13,0x430 mm steck', w:'Schweißbedarf', g:'Hilfs- und Betriebsstoffe'},
  '4323-00005': {b:'SCHWEIßKOHLEN 16 mm', w:'Schweißbedarf', g:'Hilfs- und Betriebsstoffe'},
  '4323-00006': {b:'SCHWEIßKOHLEN 19 mm', w:'Schweißbedarf', g:'Hilfs- und Betriebsstoffe'},
  '4323-00040': {b:'ELEKTRODE Garant 2,5x350 mm', w:'Schweißbedarf', g:'Hilfs- und Betriebsstoffe'},
  '4323-00041': {b:'ELEKTRODE GARANT 3,25/350', w:'Schweißbedarf', g:'Hilfs- und Betriebsstoffe'},
  '4323-00042': {b:'ELEKTRODE GARANT 4,0/350', w:'Schweißbedarf', g:'Hilfs- und Betriebsstoffe'},
  '4323-00043': {b:'ELEKTRODE GARANT 5,0/450', w:'Schweißbedarf', g:'Hilfs- und Betriebsstoffe'},
  '4323-00077': {b:'Gusselektrode Recept Cast 31 2,5', w:'Schweißbedarf', g:'Hilfs- und Betriebsstoffe'},
  '4323-00078': {b:'Gusselektrode Recept Cast 31 4,0', w:'Schweißbedarf', g:'Hilfs- und Betriebsstoffe'},
  '4323-00089': {b:'ELEKTRODE SHNI 2K 90  2,5 MM', w:'Schweißbedarf', g:'Hilfs- und Betriebsstoffe'},
  '4323-00100': {b:'SCHW.DR.  S-MO 1,2MM', w:'Schweißbedarf', g:'Hilfs- und Betriebsstoffe'},
  '4323-00108': {b:'ELEKTRODE SHNI 2K100 5,0 mm', w:'Schweißbedarf', g:'Hilfs- und Betriebsstoffe'},
  '4323-00120': {b:'Schweißdraht S-Cromo 1 1,2 mm', w:'Schweißbedarf', g:'Hilfs- und Betriebsstoffe'},
  '4323-00130': {b:'SCHWEIßDRAHT  S-CRMO2 1,2mm', w:'Schweißbedarf', g:'Hilfs- und Betriebsstoffe'},
  '4323-00135': {b:'SCHWEIßDRAHT   SG2  1,0MM', w:'Schweißbedarf', g:'Hilfs- und Betriebsstoffe'},
  '4323-00140': {b:'SCHWEIßDRAHT ESAB OK AristoRod 1', w:'Schweißbedarf', g:'Hilfs- und Betriebsstoffe'},
  '4323-00141': {b:'Schweißdraht SG3 1,6 mm', w:'Schweißbedarf', g:'Hilfs- und Betriebsstoffe'},
  '4323-00142': {b:'Schweißdraht ESAB OK Aristo Rod', w:'Schweißbedarf', g:'Hilfs- und Betriebsstoffe'},
  '4323-00146': {b:'Schweißdraht ESAB OK AristoRod 8', w:'Schweißbedarf', g:'Hilfs- und Betriebsstoffe'},
  '4323-00222': {b:'Drahtelektrode Esab Coreweld 1,2', w:'Schweißbedarf', g:'Hilfs- und Betriebsstoffe'},
  '4323-00240': {b:'Schweißdraht Megafil 745B  1,2 m', w:'Schweißbedarf', g:'Hilfs- und Betriebsstoffe'},
  '4323-00241': {b:'Schweißdraht Megafil 807 1,2 mm', w:'Schweißbedarf', g:'Hilfs- und Betriebsstoffe'},
  '4324-00008': {b:'OSNALKYD 1K HS Reinweiß', w:'Gemeinkostenmaterial', g:'Hilfs- und Betriebsstoffe'},
  '4324-00010': {b:'OSNAPOX Z 1K Tiefschwarz', w:'Gemeinkostenmaterial', g:'Hilfs- und Betriebsstoffe'},
  '4324-00012': {b:'OSNACRYL-AQUA. Rotbraun', w:'Gemeinkostenmaterial', g:'Hilfs- und Betriebsstoffe'},
  '4324-00018': {b:'Farbe, lichtgrau', w:'Gemeinkostenmaterial', g:'Hilfs- und Betriebsstoffe'},
  '4324-00020': {b:'Verdünnung für schwarze Farbe (e', w:'Gemeinkostenmaterial', g:'Hilfs- und Betriebsstoffe'},
  '4324-00021': {b:'Verdünnung für lichtgraue Farbe', w:'Gemeinkostenmaterial', g:'Hilfs- und Betriebsstoffe'},
  '4324-00023': {b:'OSNAPOX Z 1K Staubgrau', w:'Gemeinkostenmaterial', g:'Hilfs- und Betriebsstoffe'},
  '4324-00026': {b:'OSNAPOX Z 1K Steingrau', w:'Gemeinkostenmaterial', g:'Hilfs- und Betriebsstoffe'},
  '4324-00045': {b:'Lankwitzer 2K-HS-ACRYL Beige', w:'Gemeinkostenmaterial', g:'Hilfs- und Betriebsstoffe'},
  '4324-00051': {b:'Brillux Spritzverdünnung 5121', w:'Gemeinkostenmaterial', g:'Hilfs- und Betriebsstoffe'},
  '4324-00052': {b:'Brillux EP-Ester 5206 Rotbraun', w:'Gemeinkostenmaterial', g:'Hilfs- und Betriebsstoffe'},
  '4324-00053': {b:'OSNAPOX Z 1K Verkehrsrot', w:'Gemeinkostenmaterial', g:'Hilfs- und Betriebsstoffe'},
  '4324-00060': {b:'Teknozinc 3233 grey, 6 kg', w:'Gemeinkostenmaterial', g:'Hilfs- und Betriebsstoffe'},
  '4324-00061': {b:'Teknodur Combi 3560-78 Tinted 10', w:'Gemeinkostenmaterial', g:'Hilfs- und Betriebsstoffe'},
  '4324-00070': {b:'Weilburger 2K-EP-LM Beigerot', w:'Gemeinkostenmaterial', g:'Hilfs- und Betriebsstoffe'},
  '4324-00071': {b:'Weilburger Härter-EP-LM (5kg)', w:'Gemeinkostenmaterial', g:'Hilfs- und Betriebsstoffe'},
  '4324-00072': {b:'Weilburger 2K-EP-LM-DS Basaltgra', w:'Gemeinkostenmaterial', g:'Hilfs- und Betriebsstoffe'},
  '4324-00073': {b:'Härter-EP-LM Weilburger Rapidein', w:'Gemeinkostenmaterial', g:'Hilfs- und Betriebsstoffe'},
  '4324-00074': {b:'Verdünnung-EP-LM Weilburger', w:'Gemeinkostenmaterial', g:'Hilfs- und Betriebsstoffe'},
  '4324-00075': {b:'Weilburger 2K-EP-LM-DS Tiefschwa', w:'Gemeinkostenmaterial', g:'Hilfs- und Betriebsstoffe'},
  '4324-00076': {b:'Weilburger Härter-EP-LM (3,75kg)', w:'Gemeinkostenmaterial', g:'Hilfs- und Betriebsstoffe'},
  '4324-00080': {b:'Krönadal-A-KH-DS-Grundierung, ba', w:'Gemeinkostenmaterial', g:'Hilfs- und Betriebsstoffe'},
  '4324-00081': {b:'Krönadal-A-KH-Verdünnung farblos', w:'Gemeinkostenmaterial', g:'Hilfs- und Betriebsstoffe'},
  '4324-00090': {b:'Bergolin Grundierung 3PH27-S, Sc', w:'Gemeinkostenmaterial', g:'Hilfs- und Betriebsstoffe'},
  '4324-00091': {b:'Bergolin Verdünnung 5PH24', w:'Gemeinkostenmaterial', g:'Hilfs- und Betriebsstoffe'},
  '4324-00092': {b:'KRÖNADUR-EP-2K-ZP-Grundierung', w:'Gemeinkostenmaterial', g:'Hilfs- und Betriebsstoffe'},
  '4324-00093': {b:'KRÖNADUR-Spezial-Härter farblos', w:'Gemeinkostenmaterial', g:'Hilfs- und Betriebsstoffe'},
  '4324-00094': {b:'Verdünnung für Grundierung EP-2K', w:'Gemeinkostenmaterial', g:'Hilfs- und Betriebsstoffe'},
  '4324-00120': {b:'Epoxinver/E Primer Extra N.80 sc', w:'Gemeinkostenmaterial', g:'Hilfs- und Betriebsstoffe'},
  '4324-00121': {b:'EP-Härter zu Epoxinver/E Primer', w:'Gemeinkostenmaterial', g:'Hilfs- und Betriebsstoffe'},
  '4324-00122': {b:'Spezialverdünnung Gross & Perthu', w:'Gemeinkostenmaterial', g:'Hilfs- und Betriebsstoffe'},
  '4324-00130': {b:'Farbe Grossol 1901473', w:'Gemeinkostenmaterial', g:'Hilfs- und Betriebsstoffe'},
  '4324-00131': {b:'Verdünnung Grossol', w:'Gemeinkostenmaterial', g:'Hilfs- und Betriebsstoffe'},
  '4324-00200': {b:'Destilliertes Wasser', w:'Gemeinkostenmaterial', g:'Hilfs- und Betriebsstoffe'},
  '4325-00110': {b:'ACETYLEN 7,2 KG (für 6er-Bündel)', w:'Technische Gase', g:'Hilfs- und Betriebsstoffe'},
  '4325-00120': {b:'ACETHYLEN 9,0 KG (für 16er-Bünde', w:'Technische Gase', g:'Hilfs- und Betriebsstoffe'},
  '4325-00130': {b:'ACETHYLEN 8,0 KG   (Typ 50)', w:'Technische Gase', g:'Hilfs- und Betriebsstoffe'},
  '4325-00150': {b:'ARGON 10,7 CBM', w:'Technische Gase', g:'Hilfs- und Betriebsstoffe'},
  '4325-00161': {b:'ARGON 5.0 (f. Spectro.)  10,7 cb', w:'Technische Gase', g:'Hilfs- und Betriebsstoffe'},
  '4325-00170': {b:'CORGON 11,8 CBM', w:'Technische Gase', g:'Hilfs- und Betriebsstoffe'},
  '4325-00171': {b:'CRONIGON® 2 50l 200bar', w:'Technische Gase', g:'Hilfs- und Betriebsstoffe'},
  '4325-00181': {b:'KOHLENSÄURE mit Steigrohr', w:'Technische Gase', g:'Hilfs- und Betriebsstoffe'},
  '4325-00182': {b:'Kohlendioxid 12x37,5 kg Bündel', w:'Technische Gase', g:'Hilfs- und Betriebsstoffe'},
  '4325-00190': {b:'Sauerstoff 10 CBM', w:'Technische Gase', g:'Hilfs- und Betriebsstoffe'},
  '4325-00200': {b:'Sauerstoff 10 cbm (12er-Bündel)', w:'Technische Gase', g:'Hilfs- und Betriebsstoffe'},
  '4325-00210': {b:'Stickstoff 200 bar', w:'Technische Gase', g:'Hilfs- und Betriebsstoffe'},
  '4325-00300': {b:'KOHLENSÄURE FLÜSSIG', w:'Technische Gase', g:'Hilfs- und Betriebsstoffe'},
  '4325-00351': {b:'Stickstoff, flüssig', w:'Technische Gase', g:'Hilfs- und Betriebsstoffe'},
  '4325-00363': {b:'Argon 4.6 , flüssig', w:'Technische Gase', g:'Hilfs- und Betriebsstoffe'},
  '4325-00401': {b:'Sauerstoff, flüssig', w:'Technische Gase', g:'Hilfs- und Betriebsstoffe'},
  '4325-00500': {b:'PROPANGAS / 11-kg', w:'Technische Gase', g:'Hilfs- und Betriebsstoffe'},
  '4325-00506': {b:'Heizgas DIN 51622 verst.', w:'Technische Gase', g:'Hilfs- und Betriebsstoffe'},
  '4325-00515': {b:'TROCKENEIS CO2-Pellets', w:'Technische Gase', g:'Hilfs- und Betriebsstoffe'},
  '4326-00005': {b:'DIESEL', w:'Treibstoffe, Öle, Fette', g:'Hilfs- und Betriebsstoffe'},
  '4326-00085': {b:'Feroquench 2000', w:'Treibstoffe, Öle, Fette', g:'Hilfs- und Betriebsstoffe'},
  '4326-00110': {b:'Kühlschmierstoff HYCUT ET 46', w:'Treibstoffe, Öle, Fette', g:'Hilfs- und Betriebsstoffe'},
  '4326-00115': {b:'Kühlschmierstoff ADDITIV BX', w:'Treibstoffe, Öle, Fette', g:'Hilfs- und Betriebsstoffe'},
  '4328-00330': {b:'Hammerbohrer SDS 10x550/600', w:'Werkzeug', g:'Hilfs- und Betriebsstoffe'},
  '4328-00331': {b:'Hammerbohrer SDS  8x550/600', w:'Werkzeug', g:'Hilfs- und Betriebsstoffe'},
  '4329-00005': {b:'Sisalseil 10,0 mm DIN EN ISO 118', w:'sonst. Hilfs- u. Betriebsstoffe', g:'Hilfs- und Betriebsstoffe'},
  '4329-00010': {b:'PUTZLUMPEN bunte Trikotputzlappe', w:'sonst. Hilfs- u. Betriebsstoffe', g:'Hilfs- und Betriebsstoffe'},
  '4329-00025': {b:'Arecal Schweißhelfer Weldspray', w:'sonst. Hilfs- u. Betriebsstoffe', g:'Hilfs- und Betriebsstoffe'},
  '4329-00035': {b:'TESA KREPP 4319  ( 50 mm Breite', w:'sonst. Hilfs- u. Betriebsstoffe', g:'Hilfs- und Betriebsstoffe'},
  '4329-00040': {b:'Schweißspray', w:'sonst. Hilfs- u. Betriebsstoffe', g:'Hilfs- und Betriebsstoffe'},
  '4329-00047': {b:'GLEITMO 805K-1', w:'sonst. Hilfs- u. Betriebsstoffe', g:'Hilfs- und Betriebsstoffe'},
  '4329-00085': {b:'VORHÄNGESCHLOß', w:'sonst. Hilfs- u. Betriebsstoffe', g:'Hilfs- und Betriebsstoffe'},
  '4329-00090': {b:'GASANZÜNDER', w:'sonst. Hilfs- u. Betriebsstoffe', g:'Hilfs- und Betriebsstoffe'},
  '4329-00110': {b:'Profimarker weiß temperaturfest', w:'sonst. Hilfs- u. Betriebsstoffe', g:'Hilfs- und Betriebsstoffe'},
  '4329-00111': {b:'Profimarker gelb temperaturfest', w:'sonst. Hilfs- u. Betriebsstoffe', g:'Hilfs- und Betriebsstoffe'},
  '4329-00200': {b:'Fluxa-Konzentrat HKS mit RS', w:'sonst. Hilfs- u. Betriebsstoffe', g:'Hilfs- und Betriebsstoffe'},
  '4329-00201': {b:'Fluxa-Konzentrat HKS ohne RS', w:'sonst. Hilfs- u. Betriebsstoffe', g:'Hilfs- und Betriebsstoffe'},
  '4329-00202': {b:'Magnetpulverspray, Fluxa HS-O', w:'sonst. Hilfs- u. Betriebsstoffe', g:'Hilfs- und Betriebsstoffe'},
  '4329-00204': {b:'Reiniger DIFFU-THERM / BRE - 2', w:'sonst. Hilfs- u. Betriebsstoffe', g:'Hilfs- und Betriebsstoffe'},
  '4329-00230': {b:'Penetrant rot 313 DL / 300 ml Sp', w:'sonst. Hilfs- u. Betriebsstoffe', g:'Hilfs- und Betriebsstoffe'},
  '4329-00231': {b:'Entwickler weiß 70 3W / 500 ml S', w:'sonst. Hilfs- u. Betriebsstoffe', g:'Hilfs- und Betriebsstoffe'},
  '4329-00308': {b:'Lampen  HQL 250W/E40', w:'sonst. Hilfs- u. Betriebsstoffe', g:'Hilfs- und Betriebsstoffe'},
  '4329-00310': {b:'Lampen  HRL 400W/230/E40', w:'sonst. Hilfs- u. Betriebsstoffe', g:'Hilfs- und Betriebsstoffe'},
};


// ── STATE ───────────────────────────────────────────────────────────────────
let msalApp, account;
let siteId = null, listId = null;
let allItems = [];
let colByKey   = {};  // internal name → column definition (from SP)
let spUserMap  = {};  // SP user id (string) → display name (for Person-column LookupId resolution)
let resolvedFields = {};  // FORM_FIELDS key → actual SP internal name (null if not found)
let currentView = 'dashboard';
let prevView    = 'dashboard';
let wizardData  = {};
let wizardFilesArr = []; // Drag-and-drop file store for wizard attachments
let panelItemId = null;
// SP column may be misspelled "Stauts" in some tenants → always try both
const getStatusVal = item => getField(item,'Status') || getField(item,'Stauts') || '';

// ── USER SETTINGS ────────────────────────────────────────────────────────────
const SETTINGS_KEY = 'bedarfsanfrage_settings_v1';
const ADMIN_EMAIL  = 'administrator@dihag.com';

function getSettings(email) {
  const all = JSON.parse(localStorage.getItem(SETTINGS_KEY) || '{}');
  return Object.assign({ autoRefresh: false, pageSize: 100 }, all[(email||'').toLowerCase()] || {});
}
function saveUserSettings(email, patch, _skipSP = false) {
  const all = JSON.parse(localStorage.getItem(SETTINGS_KEY) || '{}');
  const key = (email||'').toLowerCase();
  all[key] = Object.assign(getSettings(email), patch);
  localStorage.setItem(SETTINGS_KEY, JSON.stringify(all));
  // Persist to SP asynchronously so other users/devices pick up the change.
  if (!_skipSP) persistSpSettings().catch(()=>{});
  return all[key];
}
function getAllUserSettings() {
  return JSON.parse(localStorage.getItem(SETTINGS_KEY) || '{}');
}

// ── SP-BASED SETTINGS STORE ───────────────────────────────────────────────────
// Settings are persisted as a JSON file in the site drive so admin-granted
// permissions are shared across all users and devices (not just local browser).
// localStorage stays as a fast in-memory cache; SP is the source of truth.
const SP_CONFIG_NAME = 'bedarfsanfrage-config.json';

async function loadSpSettings() {
  if (!siteId) return;
  try {
    const tok = await getToken();
    const url = `${API}/sites/${siteId}/drive/root:/${SP_CONFIG_NAME}:/content`;
    const r   = await fetch(url, {
      headers: { Authorization: 'Bearer ' + tok, 'Cache-Control': 'no-cache', Pragma: 'no-cache' }
    });
    if (r.status === 404) return; // file doesn't exist yet — first run
    if (!r.ok) return;
    const remote = await r.json();
    // Merge remote into localStorage: remote wins for granted flags so admin
    // grants propagate to users on all devices.
    const local = JSON.parse(localStorage.getItem(SETTINGS_KEY) || '{}');
    for (const [em, cfg] of Object.entries(remote)) {
      local[em] = Object.assign(local[em] || {}, cfg);
    }
    localStorage.setItem(SETTINGS_KEY, JSON.stringify(local));
    console.log('[loadSpSettings] synced', Object.keys(remote).length, 'user(s) from SP');
  } catch(e) {
    console.warn('[loadSpSettings]', e.message);
  }
}

async function persistSpSettings() {
  if (!siteId) return;
  try {
    const tok  = await getToken();
    const url  = `${API}/sites/${siteId}/drive/root:/${SP_CONFIG_NAME}:/content`;
    const body = localStorage.getItem(SETTINGS_KEY) || '{}';
    const r = await fetch(url, {
      method: 'PUT',
      headers: { Authorization: 'Bearer ' + tok, 'Content-Type': 'application/json' },
      body
    });
    if (!r.ok) {
      const txt = await r.text().catch(() => '');
      console.error('[persistSpSettings] HTTP', r.status, txt);
      toast(`Einstellungen konnten nicht in SharePoint gespeichert werden (${r.status})`, 'error');
    }
  } catch(e) {
    console.warn('[persistSpSettings]', e.message);
    toast('Einstellungen konnten nicht in SharePoint gespeichert werden', 'error');
  }
}

// ── AUTO-REFRESH ─────────────────────────────────────────────────────────────
// autoRefreshTimer: 1-second tick; arCountdown: seconds until next refresh
let autoRefreshTimer = null;
let arCountdown = 30;
let arPaused = false;   // user can pause without admin losing the feature-enable

function startAutoRefresh() {
  stopAutoRefresh();
  if (!account) return;
  if (!getSettings(account.username).autoRefresh) { updateARBtn(); return; }
  arCountdown = 30;
  arPaused    = false;
  autoRefreshTimer = setInterval(() => {
    if (arPaused) return;
    arCountdown--;
    if (arCountdown <= 0) {
      arCountdown = 30;
      loadItems(false);
    }
    updateARBtn();
  }, 1000);
  updateARBtn();
}
function stopAutoRefresh() {
  if (autoRefreshTimer) { clearInterval(autoRefreshTimer); autoRefreshTimer = null; }
  updateARBtn();
}
function updateARBtn() {
  const btn = $id('btn-autorefresh');
  if (!btn) return;
  // Feature visibility: only show button when admin has enabled autoRefresh for this user
  const featureOn = !!(account && getSettings(account.username).autoRefresh);
  btn.style.display = featureOn ? '' : 'none';
  if (!featureOn) return;
  const running = !!autoRefreshTimer && !arPaused;
  btn.classList.toggle('ar-on', running);
  if (running) {
    btn.title       = `Auto-Aktualisierung AN – nächste in ${arCountdown}s – klicken zum Pausieren`;
    btn.textContent = `⏱ ${arCountdown}s`;
  } else {
    btn.title       = 'Auto-Aktualisierung pausiert – klicken zum Fortsetzen';
    btn.textContent = '⏱ Pause';
  }
}
function toggleAutoRefresh() {
  // Toggle pause/resume only — does NOT change the admin-controlled autoRefresh setting
  if (!account) return;
  if (!getSettings(account.username).autoRefresh) return;
  if (!autoRefreshTimer) { startAutoRefresh(); return; }
  arPaused = !arPaused;
  updateARBtn();
}

// ── APPROVER DISPLAY ─────────────────────────────────────────────────────────
// Confirmed SP internal name: Entscheider_x002a_in (display: "Genehmiger")
// Also checks colByKey display names for future-proofing.
const APPROVER_DIRECT_KEYS = ['Entscheider_x002a_in', 'Genehmiger', 'CurrentApprover'];
const APPROVER_COL_RE = /\bgenehmiger\b|\bentscheider\b|\bbearbeiter\b|aktuelle[rs]?\s*(bearbeiter|zugewiesen|genehmiger)|current.?approver/i;

function getApproverVal(item) {
  // Helper: extract a printable string from a field value (handles Person/Lookup objects).
  function extractStr(raw) {
    if (raw == null) return null;
    if (typeof raw === 'object') {
      const name = raw.displayName || raw.LookupValue || raw.Title || raw.title || raw.text;
      return (name && name !== '[object Object]') ? String(name).trim() || null : null;
    }
    const s = String(raw).trim();
    return (s && s !== '[object Object]') ? s : null;
  }

  if (!item?.fields) return null;

  // 1. Direct known keys (fastest path — confirmed SP internal names)
  for (const k of APPROVER_DIRECT_KEYS) {
    // Text/object value
    const v = extractStr(item.fields[k]);
    if (v) return v;
    // Person-column LookupId variant (e.g. "Entscheider_x002a_inLookupId" = 42)
    const lid = item.fields[k + 'LookupId'];
    if (lid != null) {
      const name = spUserMap[String(lid)];
      if (name) return name;
    }
  }

  // 2. Match via colByKey display names (catches renamed/future columns)
  for (const [k, c] of Object.entries(colByKey)) {
    if (SYSTEM_FIELDS.has(k)) continue;
    if (APPROVER_COL_RE.test(c.displayName || k)) {
      const v = extractStr(getField(item, k));
      if (v) return v;
      // Also try LookupId variant for Person columns
      const lid = item.fields[k + 'LookupId'];
      if (lid != null) {
        const name = spUserMap[String(lid)];
        if (name) return name;
      }
    }
  }

  // 3. Scan item.fields keys directly as last resort
  for (const [k, raw] of Object.entries(item.fields)) {
    if (SYSTEM_FIELDS.has(k)) continue;
    if (APPROVER_COL_RE.test(k)) {
      const v = extractStr(raw);
      if (v) return v;
    }
  }

  return null;
}

// ── AUTH ────────────────────────────────────────────────────────────────────
async function initAuth() {
  const redirectUri = location.href.split('?')[0].split('#')[0];
  msalApp = new msal.PublicClientApplication({
    auth: { clientId:CLIENT_ID, authority:`https://login.microsoftonline.com/${TENANT_ID}`, redirectUri },
    cache: { cacheLocation:'localStorage', storeAuthStateInCookie:true }
  });
  await msalApp.initialize();
  const r = await msalApp.handleRedirectPromise();
  if (r) { account = r.account; msalApp.setActiveAccount(r.account); return true; }
  const accounts = msalApp.getAllAccounts();
  if (accounts.length) { account = accounts[0]; msalApp.setActiveAccount(accounts[0]); return true; }
  return false;
}

async function doLogin() {
  $id('boot-btn').style.display = 'none';
  $id('boot-sub').textContent = 'Weiterleitung zur Anmeldung…';
  $id('boot-spinner').style.display = 'block';
  try {
    await msalApp.loginRedirect({ scopes: SCOPES });
  } catch(e) {
    $id('boot-err').textContent = e.message;
    $id('boot-spinner').style.display = 'none';
    $id('boot-btn').style.display = 'block';
    $id('boot-btn').textContent = 'Erneut versuchen';
  }
}

async function bootDone() {
  $id('boot-sub').textContent = 'Daten werden geladen…';
  try {
    await discoverSP();
    // Load shared settings from SP (source of truth for admin grants).
    // Must happen before migration so granted flags are already in localStorage.
    await loadSpSettings();
    await loadItems(false);
    $id('boot').style.display = 'none';
    $id('app').style.display  = 'flex';
    // Auto-refresh is admin-controlled. Clear any self-set value from old localStorage
    // if admin hasn't explicitly granted the feature (autoRefreshGranted flag).
    if (account) {
      const em = account.username.toLowerCase();
      const s  = getSettings(em);
      if (em !== ADMIN_EMAIL && s.autoRefresh && !s.autoRefreshGranted) {
        saveUserSettings(em, { autoRefresh: false }, true); // _skipSP: already synced
      }
    }
    startAutoRefresh();
    applyDashboardVisibility();
    // For non-admins, rename the "Dashboard" nav label to "Meine Anfragen"
    // since the view only shows their own items.
    if (!isAdmin()) {
      const dashNav = document.querySelector('.nav-item[data-view="dashboard"]');
      if (dashNav) {
        for (const node of dashNav.childNodes) {
          if (node.nodeType === 3 /* TEXT_NODE */ && node.textContent.trim()) {
            node.textContent = node.textContent.replace('Dashboard', 'Meine Anfragen');
            break;
          }
        }
      }
    }
    // Set user info in sidebar
    const name = account?.name || account?.username || '?';
    $id('hdr-av').textContent   = name.split(' ').map(p=>p[0]||'').join('').substring(0,2).toUpperCase();
    $id('hdr-name').textContent = name;
    $id('hdr-mail').textContent = account?.username || '';
    navigate('dashboard');
  } catch(e) {
    $id('boot-sub').textContent = 'Fehler beim Laden: ' + e.message;
    $id('boot-spinner').style.display = 'none';
    $id('boot-err').textContent = e.message;
    $id('boot-btn').style.display = 'block';
    $id('boot-btn').textContent = 'Erneut versuchen';
    $id('boot-btn').onclick = bootDone;
  }
}

function doLogout() {
  msalApp?.logoutPopup({ account }).catch(()=>{});
  location.reload();
}

async function getToken() {
  if (!account) throw new Error('Nicht angemeldet');
  try { return (await msalApp.acquireTokenSilent({scopes:SCOPES, account})).accessToken; }
  catch(e) {
    if (e instanceof msal.InteractionRequiredAuthError) {
      await msalApp.acquireTokenRedirect({scopes:SCOPES});
    }
    throw e;
  }
}

// SharePoint REST API needs a token for the SP resource (different from Graph)
const SP_SCOPES = [`https://${SP_SITE.split(':/')[0]}/Sites.ReadWrite.All`];
async function getSpToken() {
  if (!account) throw new Error('Nicht angemeldet');
  try { return (await msalApp.acquireTokenSilent({ scopes: SP_SCOPES, account })).accessToken; }
  catch(e) {
    if (e instanceof msal.InteractionRequiredAuthError) {
      await msalApp.acquireTokenRedirect({ scopes: SP_SCOPES });
    }
    throw e;
  }
}

// ── GRAPH API ────────────────────────────────────────────────────────────────
async function gGet(path) {
  const tok = await getToken();
  const r   = await fetch(API + path, {
    headers: { Authorization: 'Bearer ' + tok, 'Cache-Control': 'no-cache', Pragma: 'no-cache' }
  });
  if (!r.ok) throw new Error(`Graph GET ${r.status}: ${await r.text().catch(()=>'')}`);
  return r.json();
}
async function gPost(path, body) {
  const tok = await getToken();
  const url = API + path;
  console.log('[gPost] POST', url, JSON.stringify(body).slice(0, 300));
  const r   = await fetch(url, {
    method:'POST', headers:{ Authorization:'Bearer '+tok, 'Content-Type':'application/json' },
    body: JSON.stringify(body)
  });
  if (!r.ok) {
    const txt = await r.text().catch(()=>'');
    console.error('[gPost] error', r.status, txt);
    throw new Error(`Graph POST ${r.status}: ${txt}`);
  }
  return r.json();
}
// Retry helper: SP sometimes reports 404 immediately after item creation (propagation delay).
// Wait delayMs and retry up to `retries` times when itemNotFound.
async function retryOn404(fn, retries = 4, delayMs = 1200) {
  for (let i = 0; i <= retries; i++) {
    try { return await fn(); }
    catch(e) {
      const is404 = (e.message.includes('404') || e.message.includes('itemNotFound')) && !e._noRetry;
      if (is404 && i < retries) {
        console.warn(`[retryOn404] Versuch ${i+1} – warte ${delayMs}ms…`, e.message);
        await new Promise(r => setTimeout(r, delayMs));
        continue;
      }
      throw e;
    }
  }
}

async function gPatch(path, body) {
  const tok = await getToken();
  const r   = await fetch(API + path, {
    method:'PATCH', headers:{ Authorization:'Bearer '+tok, 'Content-Type':'application/json' },
    body: JSON.stringify(body)
  });
  if (r.status === 204) return {};
  if (!r.ok) throw new Error(`Graph PATCH ${r.status}: ${await r.text().catch(()=>'')}`);
  return r.json().catch(()=>({}));
}

// ── SP DISCOVERY ─────────────────────────────────────────────────────────────
async function discoverSP() {
  // 1. Site ID
  const site = await gGet(`/sites/${SP_SITE}`);
  siteId = site.id;

  // 2. List ID
  const list = await gGet(`/sites/${siteId}/lists/${encodeURIComponent(SP_LIST)}`);
  listId = list.id;

  // 3. Columns
  const colsRes = await gGet(`/sites/${siteId}/lists/${listId}/columns`);
  colByKey = {};
  for (const c of colsRes.value) {
    colByKey[c.name] = c;  // by internal name
  }

  // 4. Resolve each form field to actual SP internal name
  resolvedFields = {};
  const missing = [];
  for (const fd of [...FORM_FIELDS, ...EINKAUF_FIELDS]) {
    const found = resolveColName(fd);
    resolvedFields[fd.key] = found;
    if (!found && fd.key !== 'Title') missing.push(fd.label);
  }
  // Title always exists
  resolvedFields['Title'] = 'Title';

  // Show warning for missing columns
  if (missing.length) {
    const w = $id('col-warning');
    w.innerHTML = `<strong>⚠ Folgende Spalten fehlen noch in der SharePoint-Liste</strong> – bitte in der SP-Liste anlegen, damit alle Felder gespeichert werden können:<br>
      <em>${missing.join(' · ')}</em>`;
    w.style.display = 'block';
  }

  // Populate Choice selects from actual SP column definitions
  for (const fd of FORM_FIELDS) {
    const colName = resolvedFields[fd.key];
    if (!colName) continue;
    const col = colByKey[colName];
    if (!col?.choice?.choices?.length) continue;
    const el = document.getElementById('f-' + fd.key);
    if (!el || el.tagName !== 'SELECT') continue;
    el.innerHTML = fd.required ? '' : '<option value="">– bitte wählen –</option>';
    for (const c of col.choice.choices) el.add(new Option(c, c));
  }

  // ── Status-Spalte mit SP-Liste synchronisieren ────────────────────────────
  // Findet die Status-Spalte (auch "Stauts" als Tippfehler-Variante), liest die
  // Auswahlwerte und baut STATUS_STYLES sowie WORKFLOW_STAGES automatisch auf.
  const statusCol = Object.values(colByKey).find(c =>
    /^status$/i.test(c.name) || /^stauts$/i.test(c.name)
  );
  if (statusCol?.choice?.choices?.length) {
    const choices = statusCol.choice.choices;
    // Rebuild STATUS_STYLES
    for (const v of choices) {
      const key = v.toLowerCase().trim();
      if (!STATUS_STYLES[key]) STATUS_STYLES[key] = statusColorFor(v);
    }
    // Rebuild WORKFLOW_STAGES: preserve order from SP list, skip "Abgelehnt"/"Erledigt" (end states)
    const endStates = /abgelehnt|erledigt|abgebrochen/i;
    const stages = choices.filter(c => !endStates.test(c));
    if (stages.length) {
      WORKFLOW_STAGES.length = 0;
      WORKFLOW_STAGES.push(...stages);
    }
    console.log('[discoverSP] Status-Werte synchronisiert:', choices);
  }

  // 5. Build SP user map (id → display name) for Person-column LookupId resolution.
  //    Person columns via Graph API return only a numeric LookupId; we need this map
  //    to show the approver's name instead of a raw ID.
  try {
    const usersRes = await gGet(
      `/sites/${siteId}/lists('User Information List')/items?$expand=fields($select=Title,EMail,Name)&$top=500`
    );
    spUserMap = {};
    for (const u of (usersRes.value || [])) {
      const name = u.fields?.Title || u.fields?.Name || u.fields?.EMail;
      if (name) spUserMap[String(u.id)] = name;
    }
    console.log('[discoverSP] spUserMap loaded:', Object.keys(spUserMap).length, 'users');
  } catch(e) {
    console.warn('[discoverSP] Could not load User Information List:', e.message);
  }
}

// Decode SP hex-encoded field names like "Gesch_x00e4_tzter" → "Geschätzter"
function decodeSpFieldName(name) {
  return (name || '').replace(/_x([0-9a-fA-F]{4})_/g, (_, hex) =>
    String.fromCharCode(parseInt(hex, 16))
  );
}

function resolveColName(fd) {
  if (fd.key === 'Title') return 'Title';
  // Only consider writable, non-sealed columns
  const ok = c => c && !c.readOnly && !c.sealed;
  if (ok(colByKey[fd.key])) return fd.key;
  for (const alt of (fd.alsoTry || [])) {
    if (ok(colByKey[alt])) return alt;
  }
  // Case-insensitive display-name match (writable only)
  const labelNorm = (fd.label || '').toLowerCase().replace(/[^a-z0-9äöüß]/g,'');
  for (const [k, c] of Object.entries(colByKey)) {
    if (!ok(c)) continue;
    const dn = (c.displayName||'').toLowerCase().replace(/[^a-z0-9äöüß]/g,'');
    if (dn === labelNorm) return k;
  }
  // Fuzzy hex-decode fallback: decode SP internal name and compare normalised
  // e.g. "Gesch_x00e4_tzterPreisnetto_x002" decoded ≈ "GeschätzterPreisnetto..."
  const keyNorm = fd.key.toLowerCase().replace(/[^a-z0-9äöüß]/g,'');
  for (const [k, c] of Object.entries(colByKey)) {
    if (!ok(c)) continue;
    const decoded = decodeSpFieldName(k).toLowerCase().replace(/[^a-z0-9äöüß]/g,'');
    if (decoded === keyNorm) return k;
    // Prefix match (SP truncates names at 32 chars)
    if (decoded.length >= 10 && (decoded.startsWith(keyNorm) || keyNorm.startsWith(decoded))) return k;
  }
  return null;
}

// ── LOAD ITEMS ───────────────────────────────────────────────────────────────
async function loadItems(showToast = true) {
  const btn = $id('btn-reload');
  if (btn) btn.disabled = true;
  try {
    const pageSize = account ? getSettings(account.username).pageSize : 100;
    const data = await gGet(
      `/sites/${siteId}/lists/${listId}/items?$expand=fields($select=*)&$top=${pageSize}&$orderby=createdDateTime desc`
    );
    allItems = data.value || [];
    if (showToast) toast('Daten aktualisiert', 'success');
    // Re-render current view
    if (currentView === 'dashboard') renderDashboard();
    else if (currentView === 'mine')  renderList('mine');
    else if (currentView === 'all')   renderList('all');
    // Refresh open panel
    if (panelItemId && ['mine','all','dashboard'].includes(currentView)) {
      const pi = allItems.find(i => String(i.id) === panelItemId);
      if (pi) { $id(`panel-${currentView}-content`).innerHTML = renderPanel(pi); bindPanelEvents(panelItemId); }
    }
  } catch(e) {
    toast('Fehler beim Laden: ' + e.message, 'error');
  } finally {
    if (btn) btn.disabled = false;
  }
}

// ── ROUTING ──────────────────────────────────────────────────────────────────
// Dashboard title differs by role: admins see "Dashboard (Alle)", others see "Meine Anfragen"
function VIEW_TITLES(view) {
  const map = { new:'Neue Bedarfsanfrage', mine:'Meine Anfragen', all:'Alle Anfragen', detail:'Anfrage Details' };
  if (view === 'dashboard') return isAdmin() ? 'Dashboard (Alle Anfragen)' : 'Meine Anfragen';
  return map[view] || view;
}

// Dashboard is visible to all logged-in users.
// Non-admins see only their own items; admins see all items.
function canSeeDashboard() { return !!account; }
function isAdmin() { return account?.username?.toLowerCase() === ADMIN_EMAIL; }

// No-op: dashboard is always visible — kept for call-site compatibility.
function applyDashboardVisibility() {
  const navItem = document.querySelector('.nav-item[data-view="dashboard"]');
  if (navItem) navItem.style.display = '';
}

function navigate(view, id) {
  if (!canSeeDashboard() && view === 'dashboard') view = 'mine'; // should never happen now
  // Always close panels of all split views when navigating
  ['mine','all','dashboard'].forEach(v => {
    $id('panel-' + v)?.classList.add('hidden');
    $id('split-' + v)?.classList.remove('has-panel');
  });
  panelItemId = null;

  document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
  document.querySelectorAll('.nav-item[data-view]').forEach(n => n.classList.remove('active'));

  const el = $id('view-' + view);
  if (!el) return;
  el.classList.add('active');

  const nav = document.querySelector(`.nav-item[data-view="${view}"]`);
  if (nav) nav.classList.add('active');

  $id('page-title').textContent = VIEW_TITLES(view);
  prevView = currentView;
  currentView = view;

  if      (view === 'dashboard') renderDashboard();
  else if (view === 'mine')      renderList('mine');
  else if (view === 'all')       renderList('all');
  else if (view === 'new')       initWizard();
  else if (view === 'detail' && id) renderDetail(id);
}

document.addEventListener('DOMContentLoaded', async () => {
  // Nav clicks
  document.querySelectorAll('.nav-item[data-view]').forEach(a => {
    a.addEventListener('click', e => { e.preventDefault(); navigate(a.dataset.view); });
  });
  // Dashboard "Neue Anfrage" button
  document.querySelectorAll('[data-view="new"]').forEach(b => {
    b.addEventListener('click', e => { e.preventDefault(); navigate('new'); });
  });
  // Detail back button
  $id('detail-back').addEventListener('click', () => navigate(prevView));
  // Menu toggle
  $id('menu-toggle').addEventListener('click', () => $id('sidebar').classList.toggle('open'));

  // Boot
  $id('boot-spinner').style.display = 'block';
  try {
    const loggedIn = await initAuth();
    if (loggedIn) await bootDone();
    else {
      $id('boot-sub').textContent = 'Bitte melden Sie sich an.';
      $id('boot-spinner').style.display = 'none';
      $id('boot-btn').style.display = 'block';
    }
  } catch(e) {
    $id('boot-err').textContent = e.message;
    $id('boot-spinner').style.display = 'none';
    $id('boot-btn').style.display = 'block';
  }
});

// ── DASHBOARD ─────────────────────────────────────────────────────────────────
let dashStatusFilter = '';
let dashSortOrder   = 'date-desc';

function renderDashboard() {
  renderStatusChips();
  filterDashboard();
}

function setDashSort(val) {
  dashSortOrder = val;
  filterDashboard();
}

function renderStatusChips() {
  const el = $id('status-chips');
  if (!el) return;

  // Non-admins see only their own items in the Dashboard.
  const baseItems = isAdmin() ? allItems : myItems();

  // Count per status
  const counts = {};
  let total = 0;
  for (const item of baseItems) {
    const s = getStatusVal(item) || '';
    counts[s] = (counts[s] || 0) + 1;
    total++;
  }

  // Volume from filtered items only
  const filteredForVol = dashStatusFilter
    ? baseItems.filter(i => (getStatusVal(i) || '') === dashStatusFilter)
    : baseItems;
  const volume = filteredForVol.reduce((sum, i) =>
    sum + (parseFloat(getField(i, resolvedFields['GeschaetzterPreis'] || 'GeschaetzterPreis')) || 0), 0);

  // Build chips: "Alle" + one per status
  const allChip = `<button class="sc-chip${dashStatusFilter==='' ? ' active' : ''}" onclick="setDashFilter('')">
    <span class="sc-label">Alle</span>
    <span class="sc-count">${total}</span>
  </button>`;

  const statusChips = Object.entries(counts).sort((a,b) => b[1]-a[1]).map(([s, n]) => {
    const sl = s.toLowerCase().trim();
    let st = STATUS_STYLES[sl];
    if (!st) {
      for (const [k, style] of Object.entries(STATUS_STYLES)) {
        if (sl.includes(k)) { st = style; break; }
      }
    }
    st = st || { bg:'#f3f4f6', color:'#374151' };
    let icon = '';
    if      (sl.includes('abgelehnt') || sl.includes('rejected'))                icon = '✗';
    else if (sl.includes('freigegeben') || sl.includes('bestellt') || sl.includes('erledigt')) icon = '✓';
    else if (sl.includes('prüfung') || sl.includes('bearbeitung'))                icon = '⏳';
    else if (sl.includes('eingereicht') || sl.includes('angefragt'))              icon = '📋';
    else if (!s)                                                                   icon = '📋';
    const isActive = dashStatusFilter === s;
    const labelText = s || 'Eingereicht';
    return `<button class="sc-chip${isActive ? ' active' : ''}"
      style="--sc-bg:${st.bg};--sc-color:${st.color}"
      onclick="setDashFilter('${s.replace(/'/g,"\\'")}')">
      ${icon ? `<span class="sc-icon">${icon}</span>` : ''}
      <span class="sc-label">${esc(labelText)}</span>
      <span class="sc-count" style="background:${st.bg};color:${st.color}">${n}</span>
    </button>`;
  }).join('');

  const vol = `<span class="sc-volume">${fmtEuro(volume)}<small>Volumen</small></span>`;
  el.innerHTML = allChip + statusChips + vol;
}

function setDashFilter(status) {
  dashStatusFilter = (dashStatusFilter === status) ? '' : status;
  renderStatusChips();
  filterDashboard();
}

function filterDashboard() {
  const search = ($id('search-dashboard')?.value || '').toLowerCase();
  // Non-admins see only their own items in the Dashboard.
  let items = isAdmin() ? [...allItems] : [...myItems()];
  if (search) items = items.filter(i =>
    (getField(i,'Title')||'').toLowerCase().includes(search) || String(i.id||'').includes(search)
  );
  if (dashStatusFilter !== '') items = items.filter(i => (getStatusVal(i)||'') === dashStatusFilter);

  const priceKey = resolvedFields['GeschaetzterPreis'] || 'GeschaetzterPreis';
  switch (dashSortOrder) {
    case 'date-asc':
      items.sort((a,b) => new Date(a.createdDateTime) - new Date(b.createdDateTime)); break;
    case 'price-desc':
      items.sort((a,b) => (parseFloat(getField(b,priceKey))||0) - (parseFloat(getField(a,priceKey))||0)); break;
    case 'price-asc':
      items.sort((a,b) => (parseFloat(getField(a,priceKey))||0) - (parseFloat(getField(b,priceKey))||0)); break;
    case 'status':
      items.sort((a,b) => (getStatusVal(a)||'').localeCompare(getStatusVal(b)||'')); break;
    default: // date-desc
      items.sort((a,b) => new Date(b.createdDateTime) - new Date(a.createdDateTime));
  }

  const container = $id('list-dashboard');
  if (container) container.innerHTML = items.length
    ? items.map(i => itemCard(i)).join('')
    : emptyState('Keine Anfragen für diesen Status.');
}

function isOpenStatus(s) {
  if (!s) return true;
  const sl = s.toLowerCase();
  return sl.includes('eingereicht') || sl.includes('prüfung') || sl.includes('offen');
}

const KANBAN_COLS = [
  { id:'eingereicht', label:'Eingereicht',  icon:'📋', match: s => !s || /eingereicht|angefragt|offen/.test(s.toLowerCase()),  bg:'#eff6ff', border:'#bfdbfe', dot:'#1d4ed8' },
  { id:'pruefung',    label:'In Prüfung',   icon:'⏳', match: s => /prüfung|bearbeitung/.test((s||'').toLowerCase()),           bg:'#fffbeb', border:'#fde68a', dot:'#b45309' },
  { id:'freigegeben', label:'Freigegeben',  icon:'✓',  match: s => /freigegeben/.test((s||'').toLowerCase()),                  bg:'#f0fdf4', border:'#bbf7d0', dot:'#15803d' },
  { id:'bestellt',    label:'Bestellt',     icon:'📦', match: s => /bestellt|erledigt/.test((s||'').toLowerCase()),             bg:'#faf5ff', border:'#e9d5ff', dot:'#7e22ce' },
  { id:'abgelehnt',   label:'Abgelehnt',   icon:'✗',  match: s => /abgelehnt|rejected/.test((s||'').toLowerCase()),            bg:'#fef2f2', border:'#fecaca', dot:'#b91c1c' },
];

function renderKanban(items) {
  if (!items.length) return `<div style="padding:24px 18px;color:#9ca3af;font-size:.88rem">Noch keine Anfragen vorhanden.</div>`;
  const cols = KANBAN_COLS.map(col => ({
    ...col,
    items: items.filter(i => col.match(getStatusVal(i)))
  }));
  return `<div class="kb-board">` +
    cols.map(col => {
      const cards = col.items.slice(0, 15).map(item => {
        const title  = getField(item,'Title') || '–';
        const preis  = parseFloat(getField(item, resolvedFields['GeschaetzterPreis'] || 'GeschaetzterPreis')) || null;
        const wg     = getField(item, resolvedFields['Warengruppe'] || 'Warengruppe') || '';
        const created = item.createdDateTime ? fmtDate(item.createdDateTime) : '';
        return `<div class="kb-mini" onclick="navigate('detail','${item.id}')">
          <div class="kb-mini-title">${esc(title)}</div>
          ${wg ? `<div class="kb-mini-wg">${esc(wg)}</div>` : ''}
          <div class="kb-mini-foot">#${item.id}${preis ? ` · ${fmtEuro(preis)}` : ''}${created ? ` · ${created}` : ''}</div>
        </div>`;
      }).join('');
      const moreCount = col.items.length - 15;
      return `
        <div class="kb-col" style="--kb-border:${col.border};--kb-bg:${col.bg}">
          <div class="kb-col-hdr">
            <span class="kb-col-icon">${col.icon}</span>
            <span class="kb-col-label">${col.label}</span>
            <span class="kb-col-count" style="background:${col.bg};color:${col.dot};border-color:${col.border}">${col.items.length}</span>
          </div>
          <div class="kb-col-body">
            ${cards || `<div class="kb-empty">–</div>`}
            ${moreCount > 0 ? `<div class="kb-more">+${moreCount} weitere</div>` : ''}
          </div>
        </div>`;
    }).join('') +
  `</div>`;
}

// ── LIST VIEWS ────────────────────────────────────────────────────────────────
let mineStatusFilter = '';
let mineSortOrder    = 'date-desc';

function myItems() {
  const myEmail = (account?.username || '').toLowerCase();
  return allItems.filter(i =>
    (i.createdBy?.user?.email || '').toLowerCase() === myEmail ||
    (i.createdBy?.user?.displayName || '').toLowerCase() === (account?.name || '').toLowerCase()
  );
}

function renderList(type) {
  if (type === 'mine') { renderStatusChipsMine(); filterView('mine'); return; }
  const sel = $id(`filter-${type}-status`);
  if (sel && sel.options.length <= 1) {
    const statuses = [...new Set(allItems.map(i => getStatusVal(i)).filter(Boolean))];
    statuses.forEach(s => sel.add(new Option(s, s)));
  }
  filterView(type);
}

function renderStatusChipsMine() {
  const el = $id('status-chips-mine');
  if (!el) return;
  const mine = myItems();
  const counts = {};
  for (const item of mine) {
    const s = getStatusVal(item) || '';
    counts[s] = (counts[s] || 0) + 1;
  }
  const filteredVol = (mineStatusFilter
    ? mine.filter(i => (getStatusVal(i)||'') === mineStatusFilter)
    : mine
  ).reduce((sum, i) => sum + (parseFloat(getField(i, resolvedFields['GeschaetzterPreis'] || 'GeschaetzterPreis')) || 0), 0);

  const allChip = `<button class="sc-chip${mineStatusFilter==='' ? ' active' : ''}" onclick="setMineFilter('')">
    <span class="sc-label">Alle</span><span class="sc-count">${mine.length}</span></button>`;

  const statusChips = Object.entries(counts).sort((a,b) => b[1]-a[1]).map(([s, n]) => {
    const sl = (s||'').toLowerCase().trim();
    let st = STATUS_STYLES[sl];
    if (!st) for (const [k, style] of Object.entries(STATUS_STYLES)) { if (sl.includes(k)) { st = style; break; } }
    st = st || { bg:'#f3f4f6', color:'#374151' };
    let icon = sl.includes('abgelehnt') ? '✗' : sl.includes('freigegeben') || sl.includes('bestellt') ? '✓'
             : sl.includes('prüfung') || sl.includes('bearbeitung') ? '⏳' : '📋';
    const isActive = mineStatusFilter === s;
    return `<button class="sc-chip${isActive ? ' active' : ''}" style="--sc-bg:${st.bg};--sc-color:${st.color}"
      onclick="setMineFilter('${s.replace(/'/g,"\\'")}')">
      <span class="sc-icon">${icon}</span>
      <span class="sc-label">${esc(s||'Eingereicht')}</span>
      <span class="sc-count" style="background:${st.bg};color:${st.color}">${n}</span>
    </button>`;
  }).join('');

  const vol = `<span class="sc-volume">${fmtEuro(filteredVol)}<small>Volumen</small></span>`;
  el.innerHTML = allChip + statusChips + vol;
}

function setMineFilter(status) {
  mineStatusFilter = (mineStatusFilter === status) ? '' : status;
  renderStatusChipsMine();
  filterView('mine');
}

function setMineSort(val) {
  mineSortOrder = val;
  filterView('mine');
}

function filterView(type) {
  const search = ($id(`search-${type}`)?.value || '').toLowerCase();
  const priceKey = resolvedFields['GeschaetzterPreis'] || 'GeschaetzterPreis';

  let items = type === 'mine' ? myItems() : [...allItems];

  if (search) items = items.filter(i =>
    (getField(i,'Title')||'').toLowerCase().includes(search) ||
    String(i.id||'').includes(search)
  );

  if (type === 'mine' && mineStatusFilter)
    items = items.filter(i => (getStatusVal(i)||'') === mineStatusFilter);
  else if (type !== 'mine') {
    const status = $id(`filter-${type}-status`)?.value || '';
    if (status) items = items.filter(i => (getStatusVal(i)||'') === status);
  }

  if (type === 'mine') {
    switch (mineSortOrder) {
      case 'date-asc':   items.sort((a,b) => new Date(a.createdDateTime)-new Date(b.createdDateTime)); break;
      case 'price-desc': items.sort((a,b) => (parseFloat(getField(b,priceKey))||0)-(parseFloat(getField(a,priceKey))||0)); break;
      case 'price-asc':  items.sort((a,b) => (parseFloat(getField(a,priceKey))||0)-(parseFloat(getField(b,priceKey))||0)); break;
      case 'status':     items.sort((a,b) => (getStatusVal(a)||'').localeCompare(getStatusVal(b)||'')); break;
      default:           items.sort((a,b) => new Date(b.createdDateTime)-new Date(a.createdDateTime));
    }
  }

  const container = $id(`list-${type}`);
  if (container) container.innerHTML = items.length
    ? items.map(i => itemCard(i)).join('')
    : emptyState(type === 'mine' ? 'Sie haben noch keine Anfragen erstellt.' : 'Keine Anfragen gefunden.');
}

// ── DETAIL VIEW ───────────────────────────────────────────────────────────────
function renderDetail(id) {
  const item = allItems.find(i => String(i.id) === String(id));
  if (!item) { $id('detail-content').innerHTML = '<p>Eintrag nicht gefunden.</p>'; return; }
  const f = item.fields || {};

  // Group known fields
  const bedarfKeys   = FORM_FIELDS.filter(fd => fd.step === 1).map(fd => resolvedFields[fd.key] || fd.key);
  const mengeKeys    = FORM_FIELDS.filter(fd => fd.step === 2).map(fd => resolvedFields[fd.key] || fd.key);
  const beschKeys    = FORM_FIELDS.filter(fd => fd.step === 3).map(fd => resolvedFields[fd.key] || fd.key);
  const einkaufKeys  = EINKAUF_FIELDS.map(fd => resolvedFields[fd.key] || fd.key);
  const knownKeys    = new Set([...bedarfKeys,...mengeKeys,...beschKeys,...einkaufKeys,'Status','Title']);

  // Remaining SP fields not in our definition
  const extraFields  = Object.entries(f)
    .filter(([k]) => !SYSTEM_FIELDS.has(k) && !knownKeys.has(k))
    .filter(([,v]) => v !== null && v !== undefined && v !== '');

  const createdBy    = item.createdBy?.user?.displayName || item.createdBy?.user?.email || getField(item,'Author') || '–';
  const createdAt    = item.createdDateTime ? fmtDateTime(item.createdDateTime) : '–';
  const statusVal    = getStatusVal(item) || 'Eingereicht';
  const preis        = parseFloat(getField(item,'GeschaetzterPreis') || getField(item, resolvedFields['GeschaetzterPreis'])) || null;
  const menge        = getField(item,'Menge') || getField(item, resolvedFields['Menge']);
  const isEinkauf    = true; // for now, all logged-in users can add order info (adjust if needed)

  $id('detail-content').innerHTML = `
    <div class="detail-header">
      <h2>${esc(getField(item,'Title') || '–')}</h2>
      <div class="detail-meta">
        <span class="item-id">ID ${item.id}</span>
        ${statusBadge(statusVal)}
        ${prioTag(getField(item,'Prioritaet') || getField(item, resolvedFields['Prioritaet']))}
      </div>
      <div class="detail-byline">Erstellt von <strong>${esc(createdBy)}</strong> am ${createdAt}</div>
    </div>

    <div class="detail-grid">
      <div class="detail-left">
        ${detailSection('Bedarf', [
          ...FORM_FIELDS.filter(fd=>fd.step===1).map(fd => detailRow(fd.label, getField(item, resolvedFields[fd.key] || fd.key)))
        ])}
        ${detailSection('Mengengaben', [
          ...FORM_FIELDS.filter(fd=>fd.step===2).map(fd => detailRow(fd.label, formatFieldValue(fd, getField(item, resolvedFields[fd.key] || fd.key))))
        ])}
        ${detailSection('Beschaffungsdetails', [
          ...FORM_FIELDS.filter(fd=>fd.step===3).map(fd => detailRow(fd.label, formatFieldValue(fd, getField(item, resolvedFields[fd.key] || fd.key))))
        ])}
        ${extraFields.length ? detailSection('Weitere Felder aus SharePoint', extraFields.map(([k,v]) => detailRow(k, v))) : ''}
      </div>

      <div class="detail-right">
        ${detailSidebar(item, isEinkauf)}
      </div>
    </div>`;

  // Bind update button
  const btnOrder = document.getElementById('btn-add-order');
  if (btnOrder) btnOrder.onclick = () => openOrderModal(item.id);

  // Load attachments async — cache-busting ensures freshly uploaded files appear immediately
  const attachEl = document.getElementById('detail-attachments');
  if (attachEl) {
    getSpToken().then(tok =>
      fetch(`${SP_BASE}/_api/web/lists/getByTitle('${SP_LIST}')/items(${id})/AttachmentFiles?_=${Date.now()}`, {
        headers: { Authorization: 'Bearer ' + tok, Accept: 'application/json;odata=nometadata',
          'Cache-Control': 'no-cache', Pragma: 'no-cache' }
      })
    ).then(r => r.ok ? r.json() : { value: [] })
     .then(data => {
       const files = data.value || [];
       attachEl.innerHTML = files.length
         ? files.map(attachmentLink).join('')
         : '<span class="no-order">Keine Anhänge.</span>';
     })
     .catch(() => { attachEl.innerHTML = ''; });
  }
}

function detailSection(title, rows) {
  const content = rows.filter(r => r).join('');
  if (!content) return '';
  return `<div class="detail-card"><div class="detail-card-header">${title}</div><div class="detail-card-body">${content}</div></div>`;
}
function detailRow(label, value) {
  if (!value && value !== 0) return '';
  return `<div class="detail-row"><span class="detail-label">${esc(String(label))}</span><span class="detail-value">${esc(String(value))}</span></div>`;
}
function formatFieldValue(fd, v) {
  if (!v && v !== 0) return '';
  if (fd.key === 'GeschaetzterPreis' || fd.key === 'TatsaechlicherPreis') return fmtEuro(v);
  if (fd.key === 'Termin' || fd.key === 'Lieferdatum') return v ? fmtDate(v) : '';
  return v;
}

function detailSidebar(item, isEinkauf) {
  const orderNr   = getField(item,'Bestellnummer')      || getField(item, resolvedFields['Bestellnummer'])      || '';
  const lieferd   = getField(item,'Lieferdatum')        || getField(item, resolvedFields['Lieferdatum'])        || '';
  const tatPreis  = getField(item,'TatsaechlicherPreis')|| getField(item, resolvedFields['TatsaechlicherPreis'])|| '';

  return `
    ${renderApprovalCard(item)}

    <div class="detail-card">
      <div class="detail-card-header">Bestellung (Einkauf)
        ${isEinkauf && /^freigegeben$/i.test((getStatusVal(item)||'').trim()) ? `<button class="btn btn-sm btn-outline" id="btn-add-order">Bearbeiten</button>` : ''}
      </div>
      <div class="detail-card-body">
        ${orderNr  ? detailRow('Bestellnummer', orderNr) : '<p class="no-order">Noch keine Bestellnummer eingetragen.</p>'}
        ${lieferd  ? detailRow('Lieferdatum', fmtDate(lieferd)) : ''}
        ${tatPreis ? detailRow('Tatsächlicher Preis', fmtEuro(tatPreis)) : ''}
        <div id="detail-attachments" class="detail-attachments">Lade Anhänge …</div>
      </div>
    </div>`;
}

// Keywords to detect approval-related SP columns by displayName or internal name
const APPROVAL_RE   = /genehmig|freigab|entscheid|ablehn|kommentar.*genehm|genehm.*kommentar/i;
const STAGE_MAP = [
  { label: 'Einkauf',                    re: /einkauf/i },
  { label: 'Werkleitung',                re: /werkleitung/i },
  { label: 'Strategischer Einkauf',      re: /strategisch/i },
  { label: 'Controlling',                re: /controlling/i },
  { label: 'Geschäftsführung',           re: /\bgf\b|geschäftsführ|geschaeftsfuehr/i },
];

// Workflow-Reihenfolge für die Status-Zeitleiste (wird in discoverSP() aus SP-Spalte befüllt)
const WORKFLOW_STAGES = [
  'Eingereicht',
  'In Prüfung (Einkauf)',
  'In Prüfung (Werkleitung)',
  'In Prüfung (Controlling)',
  'In Prüfung (strategischer Einkauf)',
  'Freigegeben',
  'Bestellt',
];

// Farbe für einen Status anhand von Schlüsselwörtern bestimmen
function statusColorFor(val) {
  const v = (val || '').toLowerCase();
  if (/eingereicht/.test(v))            return { bg:'#fce7f3', color:'#be185d' };
  if (/werkleitung/.test(v))            return { bg:'#ffe4e6', color:'#be123c' };
  if (/prüf|pruef|prüf/.test(v))        return { bg:'#ccfbf1', color:'#0f766e' };
  if (/freigegeben/.test(v))            return { bg:'#f3e8ff', color:'#7e22ce' };
  if (/abgelehnt/.test(v))              return { bg:'#f3f4f6', color:'#374151' };
  if (/bestellt/.test(v))               return { bg:'#dbeafe', color:'#1d4ed8' };
  if (/erledigt|abgeschlossen/.test(v)) return { bg:'#f3f4f6', color:'#374151' };
  return { bg:'#f0f9ff', color:'#0369a1' };
}

// Data-driven status timeline.
// Uses a FIXED semantic display order (not WORKFLOW_STAGES from discoverSP, which can be
// misordered by SP and cause completed stages to appear before earlier ones).
// When item is provided, SP approval column values are read per stage for ✓/✗ dots.
// 'strategischer Einkauf' is optional and only shown when SP data exists or it is current.
function statusTimeline(statusVal, item) {
  const sv    = (statusVal || '').trim();
  const isRej = /abgelehnt/i.test(sv);

  // Fixed display order — smIdx maps to STAGE_MAP entry for column lookup (-1 = no approval col).
  // test(sv): flexible regex match against the raw SP status value (handles SP wording variants,
  //           e.g. "strategische Einkauf" vs "strategischer Einkauf").
  const DISPLAY = [
    { label: 'Eingereicht',                         smIdx: -1, test: v => /^eingereicht$/i.test(v) },
    { label: 'In Prüfung (Einkauf)',                smIdx:  0, test: v => /pr[üu]fung/i.test(v) && /einkauf/i.test(v) && !/strategisch/i.test(v) },
    { label: 'In Prüfung (Werkleitung)',            smIdx:  1, test: v => /werkleitung/i.test(v) },
    { label: 'In Prüfung (strategischer Einkauf)', smIdx:  2, test: v => /strategisch/i.test(v) },
    { label: 'In Prüfung (Controlling)',            smIdx:  3, test: v => /controlling/i.test(v) },
    { label: 'Freigegeben',                         smIdx: -1, test: v => /^freigegeben$/i.test(v) },
    { label: 'Bestellt',                            smIdx: -1, test: v => /^bestellt$/i.test(v) },
  ];

  // Collect all approval-related columns that have a value on this item.
  const approvalCols = item
    ? Object.entries(colByKey)
        .filter(([k, c]) => APPROVAL_RE.test(c.displayName || k) && !SYSTEM_FIELDS.has(k))
        .map(([k, c]) => ({ key: k, label: c.displayName || k, val: getField(item, k) }))
        .filter(c => c.val != null && c.val !== '')
    : [];

  // Return the decision value (genehmigt/abgelehnt/…) for a STAGE_MAP[smIdx] stage.
  // smIdx 0 = Einkauf: exclude any column also matching /strategisch/ to avoid cross-contamination.
  function getDecision(smIdx) {
    if (smIdx < 0 || smIdx >= STAGE_MAP.length) return null;
    const sm = STAGE_MAP[smIdx];
    const isEinkauf = smIdx === 0;
    const cols = approvalCols.filter(c => {
      if (!(sm.re.test(c.label) || sm.re.test(c.key))) return false;
      if (isEinkauf && /strategisch/i.test(c.label + c.key)) return false;
      return /genehmig|entscheid|freigab|ablehn/i.test(c.label);
    });
    return cols.length ? cols[0].val : null;
  }

  const currentIdx = DISPLAY.findIndex(d => d.test(sv));

  return DISPLAY.map((d, i) => {
    const decision  = getDecision(d.smIdx);
    const isCurrent = d.test(sv);
    const hasData   = decision != null;

    let dot, cls;
    if (hasData) {
      const dL = String(decision).toLowerCase();
      if (/freigegeben|genehmigt|approved|ja\b/.test(dL)) { dot = '✓'; cls = 'ap-ok';      }
      else if (/abgelehnt|rejected|nein\b/.test(dL))       { dot = '✗'; cls = 'ap-no';      }
      else                                                  { dot = '○'; cls = 'ap-neutral'; }
    } else if (isCurrent) {
      dot = '●'; cls = 'ap-pending';
    } else if (currentIdx >= 0 && i < currentIdx) {
      dot = '✓'; cls = 'ap-ok';   // we've passed through this stage
    } else if (isRej && i === 0) {
      dot = '✓'; cls = 'ap-ok';   // eingereicht was reached before rejection
    } else {
      dot = '○'; cls = 'ap-neutral';
    }

    const bold = isCurrent ? ' style="font-weight:600"' : '';
    // Show current approver next to the active stage (● dot only)
    let approverHtml = '';
    if (isCurrent && item) {
      const approver = getApproverVal(item);
      if (approver) approverHtml = `<div class="ap-approver">👤 ${esc(approver)}</div>`;
    }
    return `<div class="approval-stage"><div class="ap-dot ${cls}">${dot}</div>`
         + `<div class="ap-body"><div class="ap-stage-label"${bold}>${esc(d.label)}</div>${approverHtml}</div></div>`;
  }).filter(Boolean).join('');
}

function approvalStyle(val) {
  const v = (val || '').toLowerCase();
  if (/freigegeben|genehmigt|approved|ja\b/.test(v)) return { bg:'#f0fdf4', color:'#15803d', dot:'✓', cls:'ap-ok' };
  if (/abgelehnt|rejected|nein\b/.test(v))           return { bg:'#fef2f2', color:'#b91c1c', dot:'✗', cls:'ap-no' };
  return { bg:'#fffbeb', color:'#b45309', dot:'…', cls:'ap-pending' };
}

function renderApprovalCard(item) {
  const statusVal = getStatusVal(item) || 'Eingereicht';

  // Find all approval-related columns that have a value on this item
  const found = Object.entries(colByKey)
    .filter(([k, c]) => APPROVAL_RE.test(c.displayName || k) && !SYSTEM_FIELDS.has(k))
    .map(([k, c]) => ({ key: k, label: c.displayName || k, val: getField(item, k) }))
    .filter(c => c.val !== null && c.val !== undefined && c.val !== '');

  // Group into stages; ungrouped = shown below without stage header
  const stages = STAGE_MAP.map(s => ({
    label: s.label,
    cols:  found.filter(c => s.re.test(c.label) || s.re.test(c.key)),
  })).filter(s => s.cols.length);
  const assigned = new Set(stages.flatMap(s => s.cols.map(c => c.key)));
  const ungrouped = found.filter(c => !assigned.has(c.key));

  const stagesHtml = stages.map(s => {
    // Pick the "decision" field (contains genehmig/entscheid/freigab/ablehn)
    const decisionCol = s.cols.find(c => /genehmig|entscheid|freigab|ablehn/i.test(c.label));
    const otherCols   = s.cols.filter(c => c !== decisionCol);
    const st = decisionCol ? approvalStyle(decisionCol.val) : { bg:'#f3f4f6', color:'#6b7280', dot:'○', cls:'ap-neutral' };
    return `
      <div class="approval-stage">
        <div class="ap-dot ${st.cls}">${st.dot}</div>
        <div class="ap-body">
          <div class="ap-stage-label">${esc(s.label)}</div>
          ${decisionCol ? `<span class="ap-badge" style="background:${st.bg};color:${st.color}">${esc(String(decisionCol.val))}</span>` : ''}
          ${otherCols.map(c => `<div class="ap-meta">${esc(c.label)}: ${esc(String(c.val))}</div>`).join('')}
        </div>
      </div>`;
  }).join('');

  const ungroupedHtml = ungrouped.map(c => {
    const st = approvalStyle(c.val);
    const isComment = /kommentar|ablehn|grund/i.test(c.label);
    if (isComment) return `<div class="ap-comment-box"><strong>${esc(c.label)}:</strong> ${esc(String(c.val))}</div>`;
    return `<div class="approval-stage">
      <div class="ap-dot ${st.cls}">${st.dot}</div>
      <div class="ap-body">
        <div class="ap-stage-label">${esc(c.label)}</div>
        <span class="ap-badge" style="background:${st.bg};color:${st.color}">${esc(String(c.val))}</span>
      </div>
    </div>`;
  }).join('');

  // Always show the status timeline when no stage-grouped approval columns have data.
  // Comments/ungrouped fields (e.g. GenehmigungsKommentar) are shown BELOW the timeline,
  // not instead of it – so the workflow progress stays visible even when a comment is present.
  const mainContent = stages.length ? stagesHtml : statusTimeline(statusVal, item);

  return `
    <div class="detail-card">
      <div class="detail-card-header">Status &amp; Genehmigung</div>
      <div class="detail-card-body">
        <div class="ap-current-status">${statusWithApprover(item)}</div>
        <div class="approval-stages">
          ${mainContent}
          ${ungroupedHtml}
        </div>
      </div>
    </div>`;
}

// ── EINKAUF ORDER MODAL ───────────────────────────────────────────────────────
async function openOrderModal(itemId) {
  const item = allItems.find(i => String(i.id) === String(itemId));
  if (!item || !/^freigegeben$/i.test((getStatusVal(item)||'').trim())) {
    toast('Einkauf-Daten können nur bei Status „Freigegeben" eingetragen werden.', 'error');
    return;
  }
  const orderNr    = getField(item,'Bestellnummer')       || getField(item, resolvedFields['Bestellnummer'])       || '';
  const lieferd    = getField(item,'Lieferdatum')         || getField(item, resolvedFields['Lieferdatum'])         || '';
  const tatPreis   = getField(item,'TatsaechlicherPreis') || getField(item, resolvedFields['TatsaechlicherPreis']) || '';
  const termin     = getField(item,'Termin')              || getField(item, resolvedFields['Termin'])              || '';
  const geschPreis = getField(item,'GeschaetzterPreis')   || getField(item, resolvedFields['GeschaetzterPreis'])   || '';

  // Fetch existing attachments
  let existingFiles = [];
  try {
    const tok = await getSpToken();
    const r = await fetch(`${SP_BASE}/_api/web/lists/getByTitle('${SP_LIST}')/items(${itemId})/AttachmentFiles`, {
      headers: { Authorization: 'Bearer ' + tok, Accept: 'application/json;odata=nometadata' }
    });
    if (r.ok) existingFiles = (await r.json()).value || [];
  } catch(e) { /* non-critical */ }

  const existingHtml = existingFiles.length ? `
    <div class="attach-existing">
      ${existingFiles.map(attachmentLink).join('')}
    </div>` : '';

  const required   = angeboteAnzahl(parseFloat(geschPreis) || 0);
  const attachNote = required === 0
    ? 'Angebotsunterlagen (optional)'
    : `Angebote – Regelwerk: mind. ${required} erforderlich`;
  // Always show 1 slot; extra slots (up to required-1 or 4 more) go into a collapsed <details>
  const extraN     = Math.max(0, required - 1);
  const extraSlots = Array.from({length: extraN}, (_, i) =>
    `<input type="file" class="attach-file-input" accept=".pdf,.PDF" style="margin-bottom:6px" data-slot="${i+2}"/>`
  ).join('');

  $id('modal-title').textContent = 'Einkauf-Daten eintragen';
  $id('modal-body').innerHTML = `
    <div class="form-group" style="margin-bottom:12px">
      <label>Bestellnummer</label>
      <input type="text" id="m-ordernr" value="${esc(orderNr)}" placeholder="z. B. BE252093"/>
    </div>
    <div class="form-group" style="margin-bottom:12px">
      <label style="display:flex;align-items:center;justify-content:space-between">
        Lieferdatum
        ${termin ? `<button type="button" class="btn-prefill" onclick="document.getElementById('m-delivery').value='${toLocalInputDate(termin)}'">← Aus Anfrage (${fmtDate(termin)})</button>` : ''}
      </label>
      <input type="date" id="m-delivery" value="${toLocalInputDate(lieferd)}"/>
    </div>
    <div class="form-group" style="margin-bottom:12px">
      <label style="display:flex;align-items:center;justify-content:space-between">
        Tatsächlicher Preis netto (€)
        ${geschPreis ? `<button type="button" class="btn-prefill" onclick="document.getElementById('m-price').value='${geschPreis}'">← Aus Anfrage (${fmtEuro(geschPreis)})</button>` : ''}
      </label>
      <input type="number" id="m-price" value="${tatPreis}" min="0" step="0.01" placeholder="0,00"/>
    </div>
    <div class="form-group">
      <label>${esc(attachNote)}</label>
      ${existingHtml}
      <div id="attach-inputs">
        <input type="file" class="attach-file-input" accept=".pdf,.PDF" style="margin-bottom:6px" data-slot="1"/>
      </div>
      <details id="attach-extra-wrap" style="margin-top:4px">
        <summary class="attach-more-toggle">+ Weitere Angebote${required > 1 ? ` (${extraN} gem. Regelwerk vorausgefüllt)` : ''}</summary>
        <div id="attach-extra-inputs" style="margin-top:6px">${extraSlots}</div>
        <button type="button" class="btn btn-sm btn-ghost" id="btn-add-attach" onclick="addAttachInput()" style="margin-top:4px">+ Weiteres hinzufügen</button>
      </details>
    </div>`;
  $id('modal-footer').innerHTML = `
    <button class="btn btn-ghost" onclick="closeModal()">Abbrechen</button>
    <button class="btn btn-primary" onclick="saveOrderData(${itemId})">Speichern</button>`;
  $id('modal-overlay').classList.remove('hidden');
}

function addAttachInput() {
  const extra  = $id('attach-extra-inputs') || $id('attach-inputs');
  const addBtn = $id('btn-add-attach');
  if (!extra) return;
  const total = document.querySelectorAll('.attach-file-input').length;
  if (total >= 5) { if (addBtn) addBtn.disabled = true; return; }
  const inp = document.createElement('input');
  inp.type = 'file'; inp.accept = '.pdf,.PDF'; inp.className = 'attach-file-input';
  inp.style.marginBottom = '6px';
  extra.appendChild(inp);
  if (document.querySelectorAll('.attach-file-input').length >= 5 && addBtn) addBtn.disabled = true;
}

async function saveOrderData(itemId) {
  const orderNr  = $id('m-ordernr').value.trim();
  const delivery = $id('m-delivery').value;
  const price    = $id('m-price').value;

  // Collect selected files
  const fileInputs = [...document.querySelectorAll('.attach-file-input')];
  const files = fileInputs.map(i => i.files[0]).filter(Boolean);

  // Build patch using resolved names; fall back to raw key if column was never resolved
  const patch = {};
  const add = (key, val) => {
    if (!val) return;
    const col = resolvedFields[key] || key;
    if (col) patch[col] = val;
  };
  add('Bestellnummer', orderNr);
  if (delivery) add('Lieferdatum', toSpDate(delivery, resolvedFields['Lieferdatum'] || 'Lieferdatum'));
  if (price) add('TatsaechlicherPreis', parseFloat(price));

  const hasFields = Object.keys(patch).length > 0;
  if (!hasFields && !files.length) { closeModal(); return; }

  // PATCH fields if any
  if (hasFields) {
    const skipped = [];
    for (let i = 0; i < 10; i++) {
      try {
        await gPatch(`/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, patch);
        break;
      } catch(e) {
        const m = e.message.match(/Field '([^']+)' (?:is not recognized|does not exist)/i);
        if (!m) { toast('Fehler: ' + e.message, 'error'); return; }
        skipped.push(m[1]);
        delete patch[m[1]];
      }
    }
  }

  // Upload attachments
  if (files.length) {
    const tok = await getSpToken();
    const errors = [];
    for (const file of files) {
      try {
        const fname = encodeURIComponent(file.name);
        const r = await fetch(
          `${SP_BASE}/_api/web/lists/getByTitle('${SP_LIST}')/items(${itemId})/AttachmentFiles/add(FileName='${fname}')`,
          { method: 'POST', headers: { Authorization: 'Bearer ' + tok, Accept: 'application/json;odata=nometadata' }, body: await file.arrayBuffer() }
        );
        if (!r.ok) errors.push(file.name + ' (' + r.status + ')');
      } catch(e) { errors.push(file.name); }
    }
    if (errors.length) toast('Fehler beim Hochladen: ' + errors.join(', '), 'error');
    else toast('Einkauf-Daten und Anhänge gespeichert ✓', 'success');
  } else {
    toast('Einkauf-Daten gespeichert ✓', 'success');
  }

  closeModal();
  await loadItems(false);
  renderDetail(itemId);
}

// ── BESCHAFFUNGSDETAILS MODAL (Einkauf) ───────────────────────────────────────
function openBeschModal(itemId) {
  const item = allItems.find(i => String(i.id) === String(itemId));
  if (!item) return;

  const gv  = key => getField(item, resolvedFields[key] || key) ?? '';
  const BL_OPTS = (colByKey[resolvedFields['Beschaffungslogik'] || 'Beschaffungslogik']?.choice?.choices)
    || ['Bestandsmaterial (bestandsgeführt)','Nicht-bestandsgeführtes Material','Direktes Material','Indirektes Material / Dienstleistung'];

  const selOpts = (opts, cur) => opts.map(o =>
    `<option value="${esc(o)}"${cur === o ? ' selected' : ''}>${esc(o)}</option>`
  ).join('');

  $id('modal-title').textContent = 'Beschaffungsdetails bearbeiten';
  $id('modal-body').innerHTML = `
    <div class="form-group" style="margin-bottom:12px">
      <label>Beschaffungsart</label>
      <select id="b-beschaffungslogik" class="form-control" style="width:100%">
        <option value="">–</option>${selOpts(BL_OPTS, gv('Beschaffungslogik'))}
      </select>
    </div>
    <div class="form-group" style="margin-bottom:12px">
      <label>Artikelnummer</label>
      <input type="text" id="b-artikelnummer" value="${esc(gv('Artikelnummer'))}" placeholder="z. B. 4001-00010"/>
    </div>
    <div class="form-group" style="margin-bottom:12px">
      <label>Lieferant 1</label>
      <input type="text" id="b-lieferant" value="${esc(gv('Lieferant'))}" placeholder="Firmenname oder Lieferanten-Nr."/>
    </div>
    <div class="form-group" style="margin-bottom:12px">
      <label>Lieferant 2</label>
      <input type="text" id="b-lieferant2" value="${esc(gv('Lieferant2'))}" placeholder="optional"/>
    </div>
    <div class="form-group" style="margin-bottom:12px">
      <label>Geschätzter Preis</label>
      <input type="number" id="b-preis" value="${esc(String(gv('GeschaetzterPreis')))}" min="0" step="0.01" placeholder="0,00"/>
    </div>
    <div class="form-group">
      <label>Kostenstelle</label>
      <input type="text" id="b-kostenstelle" value="${esc(gv('Kostenstelle'))}" placeholder="z. B. 4200"/>
    </div>`;
  $id('modal-footer').innerHTML = `
    <button class="btn btn-ghost" onclick="closeModal()">Abbrechen</button>
    <button class="btn btn-primary" onclick="saveBeschData(${itemId})">Speichern</button>`;
  $id('modal-overlay').classList.remove('hidden');
}

async function saveBeschData(itemId) {
  const patch = {};
  const set = (key, val) => {
    if (val === null || val === undefined || val === '') return;
    const col = resolvedFields[key] || key;
    if (col) patch[col] = val;
  };
  set('Beschaffungslogik', $id('b-beschaffungslogik').value);
  set('Artikelnummer',     $id('b-artikelnummer').value.trim());
  set('Lieferant',         $id('b-lieferant').value.trim());
  set('Lieferant2',        $id('b-lieferant2').value.trim());
  const preis = parseFloat($id('b-preis').value);
  if (!isNaN(preis) && preis > 0) {
    const col = resolvedFields['GeschaetzterPreis'] || 'GeschaetzterPreis';
    const colDef = colByKey[col];
    patch[col] = colDef?.number ? preis : String(preis);
  }
  set('Kostenstelle', $id('b-kostenstelle').value.trim());
  if (!Object.keys(patch).length) { closeModal(); return; }

  const btn = $id('modal-footer')?.querySelector('.btn-primary');
  if (btn) { btn.disabled = true; btn.textContent = 'Speichert…'; }
  const skipped = [];
  for (let i = 0; i < 10; i++) {
    try {
      await gPatch(`/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, patch);
      toast('Beschaffungsdetails gespeichert ✓', 'success');
      closeModal();
      await loadItems(false);
      const pi = allItems.find(x => String(x.id) === String(itemId));
      if (pi && panelItemId) { $id(`panel-${currentView}-content`).innerHTML = renderPanel(pi); bindPanelEvents(itemId); }
      return;
    } catch(e) {
      const m = e.message.match(/Field '([^']+)' (?:is not recognized|does not exist)/i);
      if (!m) { toast('Fehler: ' + e.message, 'error'); if (btn) { btn.disabled=false; btn.textContent='Speichern'; } return; }
      skipped.push(m[1]); delete patch[m[1]];
    }
  }
  toast('Fehler: Zu viele unbekannte Felder.', 'error');
}

// ── WIZARD ────────────────────────────────────────────────────────────────────
function lookupTID(val) {
  const tid = (val || '').trim().toUpperCase();
  const matchEl = $id('tid-match');
  // Normalize: user might type "4001-00010" or "400100010" — try exact then padded
  const hit = TID_MAP[tid] || TID_MAP[tid.replace(/^(\d{4})(\d{5})$/, '$1-$2')] || null;
  if (hit) {
    // Auto-fill Bezeichnung (Title) and Warengruppe
    const titleEl = $id('f-Title');
    const wgEl    = $id('f-Warengruppe');
    if (titleEl && !titleEl.dataset.manual) titleEl.value = hit.b;
    if (wgEl) {
      // Try to select matching option
      let found = false;
      for (const opt of wgEl.options) {
        if (opt.value === hit.w) { wgEl.value = hit.w; found = true; break; }
      }
      if (!found) {
        const newOpt = new Option(hit.w, hit.w);
        wgEl.add(newOpt);
        wgEl.value = hit.w;
      }
    }
    if (matchEl) {
      matchEl.innerHTML = `<span class="tid-ok">✓ ${esc(hit.b)}</span> <span class="tid-wg">${esc(hit.w)}</span>`;
      matchEl.style.display = 'block';
    }
  } else if (tid.length >= 4) {
    if (matchEl) {
      matchEl.innerHTML = '<span class="tid-miss">Artikelnummer nicht im Stamm – bitte Bezeichnung und Warengruppe manuell eintragen.</span>';
      matchEl.style.display = 'block';
    }
  } else {
    if (matchEl) matchEl.style.display = 'none';
  }
}

function initWizard() {
  wizardData = {};
  showStep(1);
  // Show field availability hints
  for (const fd of FORM_FIELDS) {
    const hint = $id('hint-' + fd.key);
    if (!hint) continue;
    if (resolvedFields[fd.key]) {
      hint.textContent = '';
      hint.className = 'field-hint';
    } else if (fd.key !== 'Title') {
      hint.textContent = '⚠ Spalte „' + fd.label + '" existiert noch nicht in der SP-Liste – wird beim Speichern übersprungen.';
      hint.className = 'field-hint warn';
    }
  }
  // Reset fields
  ['Title','Beschreibung','Warengruppe','Prioritaet','Mengeneinheit',
   'Mindestlagermenge','Termin','Artikelnummer','Lieferant','Lieferant2','Lieferant3','Lieferant4',
   'GeschaetzterPreis','Kostenstelle']
    .forEach(k => { const el = $id('f-'+k); if(el) el.value = ''; });
  const mengeEl = $id('f-Menge'); if (mengeEl) mengeEl.value = '1'; // Standardwert
  const firstRadio = document.querySelector('input[name=Beschaffungslogik]');
  if (firstRadio) firstRadio.checked = true;
  document.querySelectorAll('#beschaffungslogik-extra-cards .check-card').forEach(c => c.classList.remove('selected'));
  // Reset wizard attachments
  wizardFilesArr = [];
  renderWizardFiles();
  const wSec = $id('wizard-attachments');
  if (wSec) wSec.style.display = 'none';
  const wNote = $id('attach-required-note');
  if (wNote) wNote.textContent = '';
  initWizardDrop();
  const lbaEl = $id('f-LeadBuyerAbschluss');
  if (lbaEl) lbaEl.checked = false;
  // Reset progressive Lieferant disclosure
  [2,3,4].forEach(n => {
    const grp = $id('lieferant-extra-' + n);
    if (grp) grp.style.display = 'none';
  });
  // Restore "+" button on Lieferant 1 if it was removed
  const lief1 = $id('f-Lieferant')?.closest('.form-group');
  if (lief1 && !$id('btn-add-lieferant-1')) {
    const btn = document.createElement('button');
    btn.type = 'button'; btn.className = 'btn-add-lief';
    btn.id = 'btn-add-lieferant-1'; btn.onclick = () => addLieferant(2);
    btn.textContent = '+ Weiteren Lieferanten hinzufügen';
    lief1.appendChild(btn);
  }
  $id('preis-route-hint').style.display = 'none';
}

function showStep(n) {
  [1,2,3,4].forEach(i => {
    $id('wstep-'+i).classList.toggle('hidden', i !== n);
    const s = document.querySelector(`.wstep[data-step="${i}"]`);
    s.classList.remove('active','done');
    if (i < n)  s.classList.add('done');
    if (i === n) s.classList.add('active');
  });
}

function wNext(step) {
  if (step === 1) {
    const title = $id('f-Title').value.trim();
    const wg    = $id('f-Warengruppe').value;
    if (!title) { toast('Bitte Bezeichnung angeben.', 'error'); return; }
    if (!wg)    { toast('Bitte Warengruppe wählen.', 'error'); return; }
    wizardData.step1 = {
      Title:        title,
      Beschreibung: $id('f-Beschreibung').value.trim(),
      Warengruppe:  wg,
      Prioritaet:   $id('f-Prioritaet').value,
      Artikelnummer: $id('f-Artikelnummer').value.trim(),
    };
  } else if (step === 2) {
    const menge  = $id('f-Menge').value;
    const me     = $id('f-Mengeneinheit').value;
    const termin = $id('f-Termin').value;
    if (!menge || parseFloat(menge) <= 0) { toast('Bitte gültige Menge eingeben.', 'error'); return; }
    if (!me)     { toast('Bitte Mengeneinheit wählen.', 'error'); return; }
    if (!termin) { toast('Bitte Benötigt-bis-Datum angeben.', 'error'); return; }
    wizardData.step2 = {
      Menge:             menge,
      Mengeneinheit:     me,
      Mindestlagermenge: $id('f-Mindestlagermenge').value || null,
      Termin:            termin,
    };
  } else if (step === 3) {
    const lieferant = $id('f-Lieferant').value.trim();
    if (!lieferant) { toast('Bitte mindestens Lieferant 1 angeben.', 'error'); return; }
    wizardData.step3 = {
      Beschaffungslogik: [
        document.querySelector('input[name=Beschaffungslogik]:checked')?.value || '',
        ...[...document.querySelectorAll('#beschaffungslogik-extra-cards .check-card.selected')].map(c => c.dataset.value)
      ].filter(Boolean).join(', '),
      Lieferant:         $id('f-Lieferant').value.trim(),
      Lieferant2:        $id('f-Lieferant2').value.trim(),
      Lieferant3:        $id('f-Lieferant3').value.trim(),
      Lieferant4:        $id('f-Lieferant4').value.trim(),
      GeschaetzterPreis:  $id('f-GeschaetzterPreis').value ? parseFloat($id('f-GeschaetzterPreis').value) : null,
      Kostenstelle:       $id('f-Kostenstelle').value.trim(),
      LeadBuyerAbschluss: $id('f-LeadBuyerAbschluss')?.checked ?? false,
    };
    buildReview();
  }
  showStep(step + 1);
}

function toggleBLExtra(el) {
  const wasSelected = el.classList.contains('selected');
  document.querySelectorAll('#beschaffungslogik-extra-cards .check-card').forEach(c => c.classList.remove('selected'));
  if (!wasSelected) el.classList.add('selected');
}

function angeboteAnzahl(gesamt) {
  if (!gesamt || gesamt <= 500)  return 0;
  if (gesamt <= 2500)            return 2;
  if (gesamt <= 5000)            return 3;
  if (gesamt <= 10000)           return 3;
  if (gesamt <= 50000)           return 4;
  return 5;
}

function wBack(step) { showStep(step - 1); }

function addLieferant(n) {
  if (n > 4) return;
  const grp = $id('lieferant-extra-' + n);
  if (grp) grp.style.display = 'block';
  $id('btn-add-lieferant-' + (n - 1))?.remove();
  $id('f-Lieferant' + n)?.focus();
}

function updatePreisHint() {
  const preis = parseFloat($id('f-GeschaetzterPreis').value) || 0;
  const lba   = $id('f-LeadBuyerAbschluss')?.checked ?? false;
  const hint  = $id('preis-route-hint');
  if (preis > 0) {
    let text = genehmigungsweg(preis);
    if (lba) text = text.replace(/ \(entfällt bei Lead-Buyer-Abschluss\)/g, ' ✓ entfällt');
    hint.textContent = text;
    hint.style.display = 'block';
  } else {
    hint.style.display = 'none';
  }
  // Show/hide attachment section and update required-count note
  const req     = lba ? 0 : angeboteAnzahl(preis);
  const section = $id('wizard-attachments');
  const note    = $id('attach-required-note');
  if (!section) return;
  section.style.display = 'block';
  if (note) note.textContent = req === 0 ? '(optional)' : `mind. ${req} erforderlich gemäß Regelwerk`;
}

// ── WIZARD DRAG-AND-DROP ATTACHMENTS ─────────────────────────────────────────
function initWizardDrop() {
  const zone  = $id('wizard-drop-zone');
  const input = $id('wizard-file-input');
  if (!zone || !input) return;
  // Guard: only attach listeners once (called on every wizard reset)
  if (zone._dropReady) return;
  zone._dropReady = true;
  zone.addEventListener('dragover',  e => { e.preventDefault(); zone.classList.add('drag-over'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
  zone.addEventListener('drop', e => {
    e.preventDefault();
    zone.classList.remove('drag-over');
    addWizardFiles([...e.dataTransfer.files]);
  });
  zone.addEventListener('click', () => input.click());
  input.addEventListener('change', () => { addWizardFiles([...input.files]); input.value = ''; });
}

function addWizardFiles(files) {
  for (const f of files) {
    if (wizardFilesArr.length >= 5) break;
    if (!/\.pdf$/i.test(f.name)) { toast('Nur PDF-Dateien sind erlaubt.', 'error'); continue; }
    if (wizardFilesArr.some(x => x.name === f.name && x.size === f.size)) continue; // skip duplicate
    wizardFilesArr.push(f);
  }
  renderWizardFiles();
}

function removeWizardFile(idx) {
  wizardFilesArr.splice(idx, 1);
  renderWizardFiles();
}

function renderWizardFiles() {
  const list = $id('wizard-file-list');
  if (!list) return;
  list.innerHTML = wizardFilesArr.map((f, i) =>
    `<div class="wizard-file-item">
      <span class="wf-icon">📎</span>
      <span class="wf-name">${esc(f.name)}</span>
      <span class="wf-size">${(f.size/1024).toFixed(0)} KB</span>
      <button class="wf-remove" onclick="removeWizardFile(${i})" title="Entfernen">✕</button>
    </div>`
  ).join('');
  const zone = $id('wizard-drop-zone');
  if (zone) zone.classList.toggle('has-files', wizardFilesArr.length > 0);
}

function genehmigungsweg(gesamt) {
  let stufe, freigabe, angebote;
  if (gesamt <= 500) {
    stufe    = 1;
    freigabe = 'Anforderer Fachabt. + Einkäufer Werk';
    angebote = 'kein Angebot nötig';
  } else if (gesamt <= 2500) {
    stufe    = 2;
    freigabe = 'Einkäufer Werk + Werkleiter';
    angebote = 'mind. 2 Angebote (entfällt bei Lead-Buyer-Abschluss)';
  } else if (gesamt <= 5000) {
    stufe    = 2;
    freigabe = 'Einkäufer Werk + Werkleiter';
    angebote = 'mind. 2–3 Angebote (entfällt bei Lead-Buyer-Abschluss)';
  } else if (gesamt <= 10000) {
    stufe    = 3;
    freigabe = 'Werkleiter + Einkaufsleitung Holding';
    angebote = 'mind. 3 Angebote (entfällt bei Lead-Buyer-Abschluss)';
  } else if (gesamt <= 50000) {
    stufe    = 3;
    freigabe = 'Werkleiter + Einkaufsleitung Holding';
    angebote = 'mind. 3–4 Angebote (entfällt bei Lead-Buyer-Abschluss)';
  } else {
    stufe    = 4;
    freigabe = 'Einkaufsleitung Holding + GF kfm. Leitung Holding';
    angebote = 'Europ. Ausschreibung (mind. 5 Angebote, entfällt bei Lead-Buyer-Abschluss)';
  }
  return `Stufe ${stufe} · ${fmtEuro(gesamt)} · ${angebote} · Freigabe: ${freigabe}`;
}

function buildReview() {
  const d = { ...wizardData.step1, ...wizardData.step2, ...wizardData.step3 };
  const missingCols = FORM_FIELDS
    .filter(fd => fd.key !== 'Title' && !resolvedFields[fd.key] && d[fd.key])
    .map(fd => fd.label);

  $id('review-content').innerHTML = `
    <div class="review-grid">
      ${reviewSection('Bedarf', [
        ['Bezeichnung', d.Title],
        ['Beschreibung', d.Beschreibung],
        ['Warengruppe', d.Warengruppe],
        ['Priorität', d.Prioritaet],
      ])}
      ${reviewSection('Mengengaben', [
        ['Menge', d.Menge && d.Mengeneinheit ? `${d.Menge} ${d.Mengeneinheit}` : d.Menge],
        ['Mindestlagermenge', d.Mindestlagermenge],
        ['Benötigt bis', d.Termin ? fmtDate(d.Termin) : null],
      ])}
      ${reviewSection('Beschaffungsdetails', [
        ['Beschaffungsart', d.Beschaffungslogik],
        ['Artikelnummer', d.Artikelnummer],
        ['Lieferant 1', d.Lieferant],
        ['Lieferant 2', d.Lieferant2],
        ['Lieferant 3', d.Lieferant3],
        ['Lieferant 4', d.Lieferant4],
        ['Geschätzter Preis', d.GeschaetzterPreis ? fmtEuro(d.GeschaetzterPreis) : null],
        ['Kostenstelle', d.Kostenstelle],
      ])}
    </div>
    ${d.GeschaetzterPreis ? `<div class="info-box info" style="margin-top:12px">${genehmigungsweg(d.GeschaetzterPreis)}</div>` : ''}
    <div class="info-box info" style="margin-top:10px">
      Nach dem Einreichen wird <strong>Power Automate</strong> automatisch den Genehmigungsprozess starten und den zuständigen Genehmiger per E-Mail / Teams benachrichtigen.
    </div>`;

  const mw = $id('missing-cols-warn');
  if (missingCols.length) {
    mw.innerHTML = `⚠ Folgende Felder werden <strong>nicht gespeichert</strong> (Spalten fehlen in SP-Liste): <em>${missingCols.join(', ')}</em>`;
    mw.style.display = 'block';
  } else {
    mw.style.display = 'none';
  }
}

function reviewSection(title, rows) {
  const content = rows.filter(([,v]) => v).map(([l,v]) =>
    `<div class="review-row"><span class="review-label">${esc(l)}</span><span class="review-value">${esc(String(v))}</span></div>`
  ).join('');
  if (!content) return '';
  return `<div class="review-section"><h3>${title}</h3>${content}</div>`;
}

// ── SUBMIT ───────────────────────────────────────────────────────────────────
const NUMBER_FIELDS = new Set(['GeschaetzterPreis','Menge','Mindestlagermenge','TatsaechlicherPreis']);
const DATE_FIELDS   = new Set(['Termin','Lieferdatum']);
const BOOL_FIELDS   = new Set(['LeadBuyerAbschluss']);

// SP Graph requires ISO 8601 for DateTime columns: "2026-05-10T00:00:00Z"
// "Date Only" columns want just the date: "2026-05-10"
// <input type="date"> yields "2026-05-10" → append time only if column is dateTime
function toSpDate(val, spColName) {
  if (!val) return null;
  if (!/^\d{4}-\d{2}-\d{2}$/.test(val)) return val; // already has time component
  const col = spColName ? colByKey[spColName] : null;
  const isDateOnly = col?.dateTime?.format === 'dateOnly';
  return isDateOnly ? val : val + 'T00:00:00Z';
}

function buildFields(data, fieldDefs) {
  const fields = {};
  for (const fd of fieldDefs) {
    const spCol = resolvedFields[fd.key];
    const val   = data[fd.key];
    if (BOOL_FIELDS.has(fd.key)) {
      // Only send when explicitly set — avoids sending false to SP when user never touched the field.
      if (spCol && val != null && val !== '') fields[spCol] = Boolean(val);
      continue;
    }
    if (!spCol || val === null || val === undefined || val === '') continue;
    if (DATE_FIELDS.has(fd.key)) {
      const d = toSpDate(val, spCol);
      if (d) fields[spCol] = d;
    } else if (NUMBER_FIELDS.has(fd.key)) {
      const n = parseFloat(val);
      // Always send as JS number — Graph accepts numbers for all numeric/currency columns.
      // Sending String(n) to a Currency column causes 404 "itemNotFound" in some Graph versions.
      if (!isNaN(n)) fields[spCol] = n;
    } else {
      fields[spCol] = val;
    }
  }
  return fields;
}

// Retry POST removing offending fields until SP accepts the payload.
// Handles 400 "not recognized", generic 400 "invalidRequest", and 500 "generalException".
async function postRetry(path, fields) {
  const skipped = [];
  // Drop order: optional fields first, then non-Title required fields
  const optionalKeys = Object.entries(resolvedFields)
    .filter(([k]) => !FORM_FIELDS.find(f => f.key === k && f.required))
    .map(([, v]) => v).filter(Boolean);
  const requiredKeys = Object.entries(resolvedFields)
    .filter(([k]) => FORM_FIELDS.find(f => f.key === k && f.required && f.key !== 'Title'))
    .map(([, v]) => v).filter(Boolean);
  const dropQueue = [...optionalKeys, ...requiredKeys];

  for (let i = 0; i < dropQueue.length + 5; i++) {
    try {
      const newItem = await gPost(path, { fields });
      return { skipped, newItem };
    } catch(e) {
      // 400 with named field: SP says exactly which field is wrong
      const m400 = e.message.match(/Field '([^']+)' (?:is not recognized|does not exist)/i);
      if (m400) {
        const bad = m400[1];
        skipped.push(bad);
        delete fields[bad];
        for (const [k, v] of Object.entries(resolvedFields)) {
          if (v === bad) resolvedFields[k] = null;
        }
        continue;
      }
      // 500 or generic 400 "invalidRequest": SP won't say which field is wrong
      // → drop the next field from the queue and retry
      const isRetryable = e.message.includes('500') || e.message.includes('generalException') ||
        e.message.includes('invalidRequest');
      if (isRetryable) {
        const dropKey = dropQueue.find(k => k && fields[k] !== undefined);
        if (dropKey) {
          const code = e.message.includes('400') ? '400' : '500';
          console.warn(`[postRetry] ${code} – dropping field:`, dropKey, 'value:', fields[dropKey]);
          skipped.push(dropKey + `(${code})`);
          delete fields[dropKey];
          continue;
        }
      }
      throw e;
    }
  }
  throw new Error('Einreichen fehlgeschlagen – zu viele problematische Felder.');
  return { skipped: [], newItem: null }; // unreachable, satisfies linters
}

// patchRetry: updates item fields with 404-wait (SP propagation) + 400/500 field-dropping.
// On generic 400 "invalidRequest" SP won't say which field is wrong → drop fields one by
// one (optional first, required last) until the PATCH succeeds. Returns array of dropped keys.
async function patchRetry(path, fields, retries404 = 6, delay404 = 2000) {
  const skipped = [];
  // Build drop queue: optional fields first, then required (excluding Title equivalents)
  const optKeys = Object.entries(resolvedFields)
    .filter(([k]) => !FORM_FIELDS.find(f => f.key === k && f.required))
    .map(([, v]) => v).filter(v => v && fields[v] !== undefined);
  const reqKeys = Object.entries(resolvedFields)
    .filter(([k]) => FORM_FIELDS.find(f => f.key === k && f.required && f.key !== 'Title'))
    .map(([, v]) => v).filter(v => v && fields[v] !== undefined);
  const dropQueue = [...optKeys, ...reqKeys];

  let notFound404s = 0;
  for (let attempt = 0; attempt < retries404 + dropQueue.length + 5; attempt++) {
    try {
      await gPatch(path, fields);
      if (skipped.length) console.warn('[patchRetry] succeeded after dropping:', skipped);
      return skipped;
    } catch(e) {
      const msg = e.message || '';
      // 404 / itemNotFound: item not yet propagated → wait and retry
      if ((msg.includes('404') || msg.includes('itemNotFound')) && !e._noRetry) {
        if (notFound404s < retries404) { notFound404s++; await new Promise(r => setTimeout(r, delay404)); continue; }
      }
      // Named field error: SP says exactly which field
      const mField = msg.match(/Field '([^']+)' (?:is not recognized|does not exist)/i);
      if (mField) {
        const bad = mField[1];
        console.warn('[patchRetry] SP rejected field:', bad);
        skipped.push(bad); delete fields[bad];
        for (const [k, v] of Object.entries(resolvedFields)) { if (v === bad) resolvedFields[k] = null; }
        continue;
      }
      // Generic 400 "invalidRequest" / 500: drop next candidate and retry
      const isRetryable = msg.includes('invalidRequest') || msg.includes('400') ||
        msg.includes('500') || msg.includes('generalException');
      if (isRetryable) {
        const dropKey = dropQueue.find(k => fields[k] !== undefined);
        if (dropKey) {
          console.warn('[patchRetry] 400/500 → dropping field:', dropKey, '=', JSON.stringify(fields[dropKey]));
          skipped.push(dropKey); delete fields[dropKey]; continue;
        }
      }
      throw e;
    }
  }
  throw new Error('PATCH fehlgeschlagen – alle Kandidatenfelder versucht.');
}

async function submitRequest() {
  const btn = $id('btn-submit');
  btn.disabled = true;
  btn.textContent = 'Wird eingereicht…';

  try {
    const d = { ...wizardData.step1, ...wizardData.step2, ...wizardData.step3 };
    const allFields = buildFields(d, FORM_FIELDS);

    if (!allFields['Title']) { toast('Titel fehlt.', 'error'); btn.disabled=false; btn.textContent='✓ Anfrage einreichen'; return; }

    const wizardFiles = [...wizardFilesArr];

    if (!siteId || !listId) await discoverSP();

    // Use the list GUID (listId) instead of the display name — more reliable, avoids
    // name-resolution overhead and potential encoding issues in Graph API paths.
    const listPath = `/sites/${siteId}/lists/${listId}`;

    // ── Schritt 1: Element mit nur Title anlegen (minimal, kann nicht an Feldern scheitern) ──
    btn.textContent = 'Anfrage wird angelegt…';
    const newItem = await gPost(`${listPath}/items`, { fields: { Title: allFields['Title'] } });
    const itemId  = newItem.id;
    console.log('[submitRequest] item created, id=', itemId);

    // ── Schritt 2: Restliche Felder per PATCH nachpflegen ──
    // SP Online has a propagation delay: the new item may not be addressable via PATCH
    // immediately after creation. Retry with generous delay (6 × 2 s = up to 12 s total).
    const patchFields = { ...allFields };
    delete patchFields['Title'];
    if (Object.keys(patchFields).length) {
      btn.textContent = 'Felder werden gespeichert…';
      console.log('[submitRequest] PATCH fields:', JSON.stringify(patchFields));
      try {
        const skipped = await patchRetry(`${listPath}/items/${itemId}/fields`, patchFields);
        if (skipped.length) console.warn('[submitRequest] übersprungene Felder (SP ablehnte):', skipped);
      } catch(patchErr) {
        console.warn('[submitRequest] PATCH endgültig fehlgeschlagen:', patchErr.message);
        toast('Anfrage erstellt, aber einige Felder konnten nicht gespeichert werden.', 'error');
      }
    }

    // ── Schritt 3: Anhänge hochladen ──
    if (wizardFiles.length) {
      btn.textContent = 'Anhänge werden hochgeladen…';
      let tok;
      try { tok = await getSpToken(); }
      catch(tokErr) {
        console.warn('[submitRequest] SP-Token:', tokErr);
        toast('Anfrage erstellt. Anhänge konnten nicht hochgeladen werden (Token-Fehler).', 'error');
        await loadItems(false);
        navigate('mine');
        return;
      }
      const errors = [];
      for (const file of wizardFiles) {
        try {
          // SP REST may not yet have propagated the new item (same lag as PATCH 404).
          // Wrap in retryOn404 so we wait up to ~4.8s before giving up.
          const fname = encodeURIComponent(file.name);
          const uploadUrl = `${SP_BASE}/_api/web/lists/getByTitle('${SP_LIST}')/items(${itemId})/AttachmentFiles/add(FileName='${fname}')`;
          await retryOn404(async () => {
            const r = await fetch(uploadUrl, {
              method: 'POST',
              headers: { Authorization: 'Bearer ' + tok, Accept: 'application/json;odata=nometadata' },
              body: await file.arrayBuffer()
            });
            if (!r.ok) {
              const txt = await r.text().catch(() => '');
              const err = new Error(`${r.status}: ${txt}`);
              // Only retry on 404; other errors are immediate failures
              if (r.status !== 404) err._noRetry = true;
              throw err;
            }
          }, 3, 1200);
        } catch(e) { console.warn('[submitRequest] Anhang:', e); errors.push(`${file.name} (${e.message})`); }
      }
      if (errors.length) toast(`Anfrage erstellt. Anhänge fehlgeschlagen: ${errors.join(', ')}`, 'error');
      else toast(`Anfrage eingereicht! ${wizardFiles.length} Angebot(e) angehängt.`, 'success');
    } else {
      toast('Anfrage eingereicht! Power Automate startet den Genehmigungsprozess.', 'success');
    }

    await loadItems(false);
    navigate('mine');
  } catch(e) {
    toast('Fehler: ' + e.message, 'error');
    btn.disabled = false;
    btn.textContent = '✓ Anfrage einreichen';
  }
}

// ── CARD / ROW TEMPLATES ─────────────────────────────────────────────────────
function beschlShort(v) {
  if (!v) return '';
  if (/bestand/i.test(v))  return '📦 Bestand';
  if (/neu|nicht.bestand/i.test(v)) return '🆕 Neu';
  if (/dienst/i.test(v))   return '🔧 Dienstl.';
  return v.substring(0, 18);
}

function getApprovalSummary(item) {
  const vals = Object.entries(colByKey)
    .filter(([k,c]) => APPROVAL_RE.test(c.displayName || k) && !SYSTEM_FIELDS.has(k))
    .map(([k]) => getField(item, k)).filter(Boolean);
  if (!vals.length) return '';
  if (vals.some(v => /abgelehnt|rejected/i.test(v)))
    return `<span class="ica-no">✗ Abgelehnt</span>`;
  const ok = vals.filter(v => /freigegeben|genehmigt|approved/i.test(v));
  if (ok.length) return `<span class="ica-ok">✓ ${ok.length}× Freigabe</span>`;
  return `<span class="ica-pending">⏳ In Prüfung</span>`;
}

function itemCard(item) {
  const title   = esc(getField(item,'Title') || '–');
  const status  = getStatusVal(item);
  const wg      = getField(item, resolvedFields['Warengruppe']    || 'Warengruppe')      || '';
  const preis   = parseFloat(getField(item, resolvedFields['GeschaetzterPreis'] || 'GeschaetzterPreis')) || null;
  const prio    = getField(item, resolvedFields['Prioritaet']     || 'Prioritaet')       || '';
  const menge   = getField(item, resolvedFields['Menge']          || 'Menge')            || '';
  const me      = getField(item, resolvedFields['Mengeneinheit']  || 'Mengeneinheit')    || '';
  const beschl  = getField(item, resolvedFields['Beschaffungslogik'] || 'Beschaffungslogik') || '';
  const liefant = getField(item, resolvedFields['Lieferant']      || 'Lieferant')        || '';
  const termin  = getField(item, resolvedFields['Termin']         || 'Termin')           || '';
  const ks      = getField(item, resolvedFields['Kostenstelle']   || 'Kostenstelle')     || '';
  const created = item.createdDateTime ? fmtDate(item.createdDateTime) : '';
  const creator = item.createdBy?.user?.displayName || item.createdBy?.user?.email || '';
  const sel     = panelItemId === String(item.id) ? ' selected' : '';

  return `
    <div class="item-card${sel}" data-id="${item.id}" onclick="openPanel('${item.id}')">
      <div class="ic-top">
        <div class="ic-title">${prioDot(prio)}${cleanTitle(getField(item,'Title'))}</div>
        <div class="ic-topright">
          <span class="ic-price">${preis ? fmtEuro(preis) : '–'}</span>
          ${statusBadge(status)}
        </div>
      </div>
      <div class="ic-tags">
        <span class="ic-tag ic-banf">#${item.id}</span>
        ${wg      ? `<span class="ic-tag ic-wg">${esc(wg)}</span>` : ''}
        ${beschl  ? `<span class="ic-tag">${beschlShort(beschl)}</span>` : ''}
        ${menge && me ? `<span class="ic-tag">⚖ ${esc(menge)} ${esc(me)}</span>` : ''}
        ${ks      ? `<span class="ic-tag">KST ${esc(ks)}</span>` : ''}
        ${liefant ? `<span class="ic-tag">🏭 ${esc(liefant)}</span>` : ''}
        ${termin  ? `<span class="ic-tag">📅 bis ${fmtDate(termin)}</span>` : ''}
      </div>
      <div class="ic-footer">
        <span class="ic-by">${creator ? `👤 ${esc(creator)}` : ''} ${created ? `· ${created}` : ''}</span>
      </div>
    </div>`;
}

function renderApprovalHighlight(item) {
  const found = Object.entries(colByKey)
    .filter(([k,c]) => APPROVAL_RE.test(c.displayName||k) && !SYSTEM_FIELDS.has(k))
    .map(([k,c]) => ({ label: c.displayName||k, val: getField(item,k) }))
    .filter(c => c.val);
  if (!found.length) return '';

  const rejected = found.find(c => /abgelehnt|rejected/i.test(c.val));
  const approved = found.filter(c => /freigegeben|genehmigt|approved/i.test(c.val));
  const pending  = found.filter(c => !/abgelehnt|rejected|freigegeben|genehmigt|approved/i.test(c.val));

  if (rejected) {
    const comment = found.find(c => /kommentar|ablehn|grund/i.test(c.label) && c.val);
    return `<div class="cr-appr cr-appr-no">
      <span class="cr-appr-icon">✗</span>
      <div><strong>Abgelehnt</strong>${comment ? ` — ${esc(String(comment.val))}` : ''}</div>
    </div>`;
  }
  if (approved.length) {
    const stages = approved.map(c => esc(c.label)).join(', ');
    return `<div class="cr-appr cr-appr-ok">
      <span class="cr-appr-icon">✓</span>
      <div><strong>${approved.length}× Freigabe</strong> — ${stages}</div>
    </div>`;
  }
  if (pending.length) {
    const stages = pending.map(c => `${esc(c.label)}: ${esc(String(c.val))}`).join(' · ');
    return `<div class="cr-appr cr-appr-pending">
      <span class="cr-appr-icon">⏳</span>
      <div>${stages}</div>
    </div>`;
  }
  return '';
}

function compactRow(item) {
  const title   = getField(item,'Title') || '–';
  const status  = getStatusVal(item);
  const prio    = getField(item, resolvedFields['Prioritaet'] || 'Prioritaet') || '';
  const wg      = getField(item, resolvedFields['Warengruppe'] || 'Warengruppe') || '';
  const preis   = parseFloat(getField(item, resolvedFields['GeschaetzterPreis'] || 'GeschaetzterPreis')) || null;
  const beschl  = getField(item, resolvedFields['Beschaffungslogik'] || 'Beschaffungslogik') || '';
  const menge   = getField(item, resolvedFields['Menge'] || 'Menge') || '';
  const me      = getField(item, resolvedFields['Mengeneinheit'] || 'Mengeneinheit') || '';
  const termin  = getField(item, resolvedFields['Termin'] || 'Termin') || '';
  const created = item.createdDateTime ? fmtDate(item.createdDateTime) : '';
  const creator = item.createdBy?.user?.displayName || item.createdBy?.user?.email || '';
  const apprSum = getApprovalSummary(item);

  const tags = [
    wg     ? `<span class="cr2-tag cr2-wg">${esc(wg)}</span>` : '',
    beschl ? `<span class="cr2-tag">${beschlShort(beschl)}</span>` : '',
    menge && me ? `<span class="cr2-tag">⚖ ${esc(menge)} ${esc(me)}</span>` : '',
    termin ? `<span class="cr2-tag">📅 ${fmtDate(termin)}</span>` : '',
  ].filter(Boolean).join('');

  return `
    <div class="cr2" onclick="navigate('detail','${item.id}')">
      <div class="cr2-head">
        <div class="cr2-title">${prioDot(prio)}${esc(title)}</div>
        <div class="cr2-right">
          ${preis ? `<span class="cr2-price">${fmtEuro(preis)}</span>` : ''}
          ${statusBadge(status)}
        </div>
      </div>
      ${tags ? `<div class="cr2-tags">${tags}</div>` : ''}
      <div class="cr2-foot">
        <span class="cr2-meta">${creator ? `👤 ${esc(creator)}` : ''}${created ? ` · ${created}` : ''} · #${item.id}</span>
        ${apprSum}
      </div>
    </div>`;
}

// ── SPLIT PANEL ───────────────────────────────────────────────────────────────
function openPanel(itemId) {
  if (!['mine','all','dashboard'].includes(currentView)) { navigate('detail', itemId); return; }
  panelItemId = String(itemId);
  const item  = allItems.find(i => String(i.id) === panelItemId);
  if (!item) return;

  $id(`panel-${currentView}-content`).innerHTML = renderPanel(item);
  $id(`panel-${currentView}`).classList.remove('hidden');
  $id(`split-${currentView}`).classList.add('has-panel');

  document.querySelectorAll('.item-card').forEach(c =>
    c.classList.toggle('selected', c.dataset.id === panelItemId));

  bindPanelEvents(itemId);
}

function closePanel() {
  ['mine','all','dashboard'].forEach(v => {
    $id('panel-' + v)?.classList.add('hidden');
    $id('split-' + v)?.classList.remove('has-panel');
  });
  panelItemId = null;
  _vhLoaded = {};
  document.querySelectorAll('.item-card').forEach(c => c.classList.remove('selected'));
}

function bindPanelEvents(itemId) {
  $id('panel-close')?.addEventListener('click', closePanel);
  $id('panel-order')?.addEventListener('click', () => openOrderModal(itemId));
  $id('panel-besch')?.addEventListener('click', () => openBeschModal(itemId));
  // Load attachments
  const attachEl = $id('panel-attach-body');
  if (attachEl) {
    getSpToken()
      .then(tok => fetch(`${SP_BASE}/_api/web/lists/getByTitle('${SP_LIST}')/items(${itemId})/AttachmentFiles?_=${Date.now()}`,
        { headers: { Authorization: 'Bearer ' + tok, Accept: 'application/json;odata=nometadata',
            'Cache-Control': 'no-cache', Pragma: 'no-cache' } }))
      .then(r => r.ok ? r.json() : { value: [] })
      .then(data => {
        const files = data.value || [];
        attachEl.innerHTML = files.length
          ? files.map(attachmentLink).join('')
          : '<span class="no-order">Keine Anhänge.</span>';
      })
      .catch(() => { attachEl.innerHTML = '<span class="no-order">Anhänge konnten nicht geladen werden.</span>'; });
  }
  $id('panel-history')?.addEventListener('click', () => {
    const sec = $id('panel-history-section');
    if (!sec) return;
    const isOpen = sec.style.display !== 'none';
    sec.style.display = isOpen ? 'none' : '';
    $id('panel-history').textContent = isOpen ? '📋 Verlauf' : '📋 Verlauf ▲';
    if (!isOpen) loadVersionHistory(itemId);
  });
}

let _vhLoaded = {};
async function loadVersionHistory(itemId) {
  const el = $id('panel-history-body');
  if (!el) return;
  if (_vhLoaded[itemId]) return;
  _vhLoaded[itemId] = true;

  try {
    const data = await gGet(`/sites/${siteId}/lists/${listId}/items/${itemId}/versions?$expand=fields($select=*)&$top=50`);
    const vers = (data.value || []).sort((a,b) => new Date(b.lastModifiedDateTime) - new Date(a.lastModifiedDateTime));
    if (!vers.length) { el.innerHTML = '<div class="vh-empty">Keine Versionen gefunden.</div>'; return; }

    // Helper: SP Choice → {Value:...}, Lookup → {LookupValue:...}
    const spVal = v => {
      if (v === null || v === undefined) return '';
      if (typeof v === 'object') return String(v.Value ?? v.LookupValue ?? '');
      return String(v);
    };
    const fmtVH = (k, v) => {
      const s = spVal(v);
      if (!s) return '–';
      if (/preis/i.test(k))              return fmtEuro(parseFloat(s)) || s;
      if (/datum|termin|ben.?tigt/i.test(k)) return fmtDate(s);
      return s;
    };

    // Reverse map: any known SP column name → display label
    const labelMap = { Status:'Status' };
    for (const f of [...FORM_FIELDS, ...EINKAUF_FIELDS]) {
      labelMap[f.key] = f.label;
      const col = resolvedFields[f.key];
      if (col) labelMap[col] = f.label;
    }

    // Fields to always skip
    const skipKey = k =>
      SYSTEM_FIELDS.has(k) || k.endsWith('LookupId') ||
      k.startsWith('@') || k.startsWith('_') ||
      /^(id|Edit|Attachments|ContentType|AppAuthor|AppEditor|LinkTitle|ComplianceAsset)$/i.test(k);

    const statusKey = resolvedFields['Status'] || 'Status';

    const rows = vers.map((v, idx) => {
      const prev    = vers[idx + 1];
      const status  = spVal(v.fields?.[statusKey]);
      const by      = v.lastModifiedBy?.user?.displayName || v.lastModifiedBy?.user?.email || 'System';
      const dt      = fmtDateTime(v.lastModifiedDateTime);
      const isFirst = idx === vers.length - 1;

      const changes = [];
      if (isFirst) {
        changes.push('Anfrage erstellt');
      } else {
        const allKeys = new Set([
          ...Object.keys(v.fields || {}),
          ...Object.keys(prev?.fields || {}),
        ]);
        for (const k of allKeys) {
          if (skipKey(k)) continue;
          const cur = v.fields?.[k];
          const prv = prev?.fields?.[k];
          if (spVal(cur) === spVal(prv)) continue;
          const label = labelMap[k] || k;
          changes.push(`${esc(label)}: <em>${esc(fmtVH(k, prv))}</em> → <em>${esc(fmtVH(k, cur))}</em>`);
        }
        if (!changes.length) changes.push('Systemfelder aktualisiert');
      }

      const st  = STATUS_STYLES[(status||'').toLowerCase().trim()] || { bg:'#f3f4f6', color:'#6b7280' };
      const dot = status?.toLowerCase().includes('freigegeben') ? '#15803d'
                : status?.toLowerCase().includes('abgelehnt')   ? '#b91c1c'
                : status?.toLowerCase().includes('prüfung')     ? '#b45309'
                : '#6b7280';

      return `<div class="vh-item${isFirst?' vh-first':''}">
        <div class="vh-line"></div>
        <div class="vh-dot-wrap"><div class="vh-dot" style="background:${dot}"></div></div>
        <div class="vh-content">
          <div class="vh-meta"><span class="vh-by">👤 ${esc(by)}</span><span class="vh-time">${dt}</span></div>
          <div class="vh-changes">${changes.join(' · ')}</div>
          ${status ? `<span class="status-badge" style="background:${st.bg};color:${st.color};margin-top:4px;display:inline-block">${esc(status)}</span>` : ''}
        </div>
      </div>`;
    });

    el.innerHTML = `<div class="vh-timeline">${rows.join('')}</div>`;
  } catch(e) {
    el.innerHTML = `<div class="vh-error">Verlauf konnte nicht geladen werden.</div>`;
  }
}

function bindPanelEditEvents(itemId) {
  $id('panel-close')?.addEventListener('click', closePanel);
  $id('panel-save')?.addEventListener('click', () => saveEdits(itemId));
  $id('panel-cancel')?.addEventListener('click', () => {
    const item = allItems.find(i => String(i.id) === String(itemId));
    if (item) { $id(`panel-${currentView}-content`).innerHTML = renderPanel(item); bindPanelEvents(itemId); }
  });
}

async function saveEdits(itemId) {
  const data = {};
  document.querySelectorAll('.pf-input[data-key]').forEach(inp => {
    data[inp.dataset.key] = inp.value;
  });
  const fields = buildFields(data, FORM_FIELDS);
  if (!Object.keys(fields).length) return;

  const btn = $id('panel-save');
  if (btn) { btn.disabled = true; btn.textContent = 'Speichert…'; }

  const patch = { ...fields };
  const skipped = [];
  for (let i = 0; i < 15; i++) {
    try {
      await gPatch(`/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, patch);
      if (skipped.length) toast(`Gespeichert (übersprungen: ${skipped.join(', ')})`, 'info');
      else toast('Gespeichert ✓', 'success');
      await loadItems(false);
      return;
    } catch(e) {
      const m = e.message.match(/Field '([^']+)' (?:is not recognized|does not exist)/i);
      if (!m) { toast('Fehler: ' + e.message, 'error'); if (btn) { btn.disabled=false; btn.textContent='Speichern'; } return; }
      skipped.push(m[1]); delete patch[m[1]];
    }
  }
  toast('Fehler: Zu viele unbekannte Felder.', 'error');
}

function renderPanel(item, editMode = false) {
  const statusVal = getStatusVal(item) || 'Eingereicht';
  const createdBy = item.createdBy?.user?.displayName || item.createdBy?.user?.email || '–';
  const createdAt = item.createdDateTime ? fmtDate(item.createdDateTime) : '–';

  const gv  = key => getField(item, resolvedFields[key] || key) ?? '';
  const dv  = v   => v ? String(v).slice(0,10) : '';  // YYYY-MM-DD for date inputs

  const choices = key => {
    const sp = resolvedFields[key] || key;
    return colByKey[sp]?.choice?.choices || null;
  };

  const WG_OPTS   = choices('Warengruppe')      || ['Schrott','Legierungen','Kaufteile','Kaufteile Formerei (Kühlm. ..)','Formstoffkomponenten','Eingußsysteme','exotherme Kappen (Speiser)','Gießfilter','Kerne (bez. Kerne f. Gießerei)','metallurgische Zusatzstoffe','Hilfsstoffe Gießerei','Feuerfestmaterial Ofen/Pfannen','Kohleprodukte (Grafitelektr., Aufk...)','Ersatzteile','Befestigungsteile','Schleif- und Fräsmittel','Strahlmittel','Werkzeug','Arbeitsschutzmittel','Schweißbedarf','Gemeinkostenmaterial','Technische Gase','Treibstoffe, Öle, Fette','sonst. Hilfs- u. Betriebsstoffe'];
  const PRIO_OPTS = choices('Prioritaet')       || ['Normal','Hoch','Dringend'];
  const ME_OPTS   = choices('Mengeneinheit')    || ['Lagereinheiten','kg','Stück','Anzahl','m','Paar','Liter'];
  const BL_OPTS   = choices('Beschaffungslogik')|| ['Bestandsmaterial (bestandsgeführt)','Neues Material (nicht-bestandsgeführt)','Dienstleistung'];

  const fRow = (fd, type, opts) => {
    const raw = gv(fd.key);
    const lbl = `<span class="pf-label">${esc(fd.label)}</span>`;
    if (editMode) {
      let inp;
      if (opts) {
        const optHtml = (fd.required ? [] : ['']).concat(opts)
          .map(o => `<option value="${esc(o)}"${String(raw)===o?' selected':''}>${esc(o||'–')}</option>`).join('');
        inp = `<select class="pf-input" data-key="${fd.key}">${optHtml}</select>`;
      } else if (type === 'textarea') {
        inp = `<textarea class="pf-input" data-key="${fd.key}" rows="2">${esc(String(raw))}</textarea>`;
      } else {
        const val = type === 'date' ? dv(raw) : esc(String(raw));
        inp = `<input type="${type}" class="pf-input" data-key="${fd.key}" value="${val}"/>`;
      }
      return `<div class="pf-row">${lbl}${inp}</div>`;
    }
    if (!raw && raw !== 0) return '';
    let display = String(raw);
    if (fd.key === 'GeschaetzterPreis' || fd.key === 'TatsaechlicherPreis') display = fmtEuro(raw);
    else if (fd.key === 'Termin' || fd.key === 'Lieferdatum') display = fmtDate(raw);
    return `<div class="pf-row">${lbl}<span class="pf-val">${esc(display)}</span></div>`;
  };

  const preis = parseFloat(gv('GeschaetzterPreis')) || 0;
  const menge = parseFloat(gv('Menge')) || 1;
  const gesamtHint = !editMode && preis > 0 ? genehmigungsweg(preis) : '';

  const orderNr  = gv('Bestellnummer');
  const lieferd  = gv('Lieferdatum');
  const tatPreis = gv('TatsaechlicherPreis');

  const isFreigegeben = /^freigegeben$/i.test((statusVal || '').trim());
  const buttons = `${isFreigegeben
      ? `<button class="btn btn-outline btn-sm" id="panel-order">📦 Einkauf-Daten</button>`
      : `<button class="btn btn-outline btn-sm" id="panel-order" disabled title="Nur bei Status 'Freigegeben' möglich">📦 Einkauf-Daten</button>`}
     <button class="btn btn-outline btn-sm" id="panel-besch">✏️ Beschaffung</button>
     <button class="btn btn-outline btn-sm" id="panel-history">📋 Verlauf</button>`;

  // Approval inner HTML (reuse logic from renderApprovalCard but without the wrapping card)
  const approvalInner = (() => {
    const found = Object.entries(colByKey)
      .filter(([k,c]) => APPROVAL_RE.test(c.displayName||k) && !SYSTEM_FIELDS.has(k))
      .map(([k,c]) => ({ key:k, label:c.displayName||k, val:getField(item,k) }))
      .filter(c => c.val !== null && c.val !== undefined && c.val !== '');
    if (!found.length) return `<div class="approval-stages">${statusTimeline(statusVal, item)}</div>`;
    const stages = STAGE_MAP.map(s => ({ label:s.label, cols:found.filter(c=>s.re.test(c.label)||s.re.test(c.key)) })).filter(s=>s.cols.length);
    const assigned = new Set(stages.flatMap(s=>s.cols.map(c=>c.key)));
    const extra = found.filter(c=>!assigned.has(c.key));
    const mkStage = s => {
      const dec = s.cols.find(c=>/genehmig|entscheid|freigab|ablehn/i.test(c.label));
      const others = s.cols.filter(c=>c!==dec);
      const st = dec ? approvalStyle(dec.val) : {bg:'#f3f4f6',color:'#6b7280',dot:'○',cls:'ap-neutral'};
      return `<div class="approval-stage"><div class="ap-dot ${st.cls}">${st.dot}</div><div class="ap-body">
        <div class="ap-stage-label">${esc(s.label)}</div>
        ${dec?`<span class="ap-badge" style="background:${st.bg};color:${st.color}">${esc(String(dec.val))}</span>`:''}
        ${others.map(c=>`<div class="ap-meta">${esc(c.label)}: ${esc(String(c.val))}</div>`).join('')}
      </div></div>`;
    };
    const mkExtra = c => {
      const st = approvalStyle(c.val);
      if (/kommentar|ablehn|grund/i.test(c.label)) return `<div class="ap-comment-box"><strong>${esc(c.label)}:</strong> ${esc(String(c.val))}</div>`;
      return `<div class="approval-stage"><div class="ap-dot ${st.cls}">${st.dot}</div><div class="ap-body"><div class="ap-stage-label">${esc(c.label)}</div><span class="ap-badge" style="background:${st.bg};color:${st.color}">${esc(String(c.val))}</span></div></div>`;
    };
    // Same rule as renderApprovalCard: always show timeline when no stage columns
    // have data; comments/ungrouped shown BELOW, not instead of the timeline.
    const mainHtml = stages.length ? stages.map(mkStage).join('') : statusTimeline(statusVal, item);
    return `<div class="approval-stages">${mainHtml}${extra.map(mkExtra).join('')}</div>`;
  })();

  return `
    <div class="panel-hdr">
      <div class="panel-hdr-top">
        <div class="panel-meta">
          <span class="item-id">ID ${item.id}</span>
          ${statusWithApprover(item)}
          ${prioTag(gv('Prioritaet'))}
        </div>
        <button class="panel-close" id="panel-close" title="Schließen">✕</button>
      </div>
      <div class="panel-title">${esc(gv('Title') || '–')}</div>
      <div class="panel-byline">von ${esc(createdBy)} · ${createdAt}</div>
      <div class="panel-actions">${buttons}</div>
    </div>

    <div class="panel-body">
      <div class="pf-section">
        <div class="pf-sec-title">Bedarf</div>
        ${fRow(FORM_FIELDS.find(f=>f.key==='Title'),       'text')}
        ${fRow(FORM_FIELDS.find(f=>f.key==='Beschreibung'),'textarea')}
        ${fRow(FORM_FIELDS.find(f=>f.key==='Warengruppe'), 'text', WG_OPTS)}
        ${fRow(FORM_FIELDS.find(f=>f.key==='Prioritaet'),  'text', PRIO_OPTS)}
      </div>
      <div class="pf-section">
        <div class="pf-sec-title">Mengengaben</div>
        ${fRow(FORM_FIELDS.find(f=>f.key==='Menge'),             'number')}
        ${fRow(FORM_FIELDS.find(f=>f.key==='Mengeneinheit'),     'text', ME_OPTS)}
        ${fRow(FORM_FIELDS.find(f=>f.key==='Mindestlagermenge'), 'number')}
        ${fRow(FORM_FIELDS.find(f=>f.key==='Termin'),            'date')}
      </div>
      <div class="pf-section">
        <div class="pf-sec-title">Beschaffungsdetails</div>
        ${fRow(FORM_FIELDS.find(f=>f.key==='Beschaffungslogik'),'text', BL_OPTS)}
        ${fRow(FORM_FIELDS.find(f=>f.key==='Artikelnummer'),    'text')}
        ${fRow(FORM_FIELDS.find(f=>f.key==='Lieferant'),        'text')}
        ${fRow(FORM_FIELDS.find(f=>f.key==='GeschaetzterPreis'),'number')}
        ${fRow(FORM_FIELDS.find(f=>f.key==='Kostenstelle'),     'text')}
        ${gesamtHint ? `<div class="info-box info" style="margin-top:8px;font-size:.78rem">${gesamtHint}</div>` : ''}
      </div>
      <div class="pf-section">
        <div class="pf-sec-title">Status &amp; Genehmigung</div>
        ${approvalInner}
      </div>
      <div class="pf-section">
        <div class="pf-sec-title">Bestellung (Einkauf)</div>
        ${orderNr  ? `<div class="pf-row"><span class="pf-label">Bestellnummer</span><span class="pf-val">${esc(orderNr)}</span></div>` : ''}
        ${lieferd  ? `<div class="pf-row"><span class="pf-label">Lieferdatum</span><span class="pf-val">${fmtDate(lieferd)}</span></div>` : ''}
        ${tatPreis ? `<div class="pf-row"><span class="pf-label">Tatsächl. Preis</span><span class="pf-val">${fmtEuro(tatPreis)}</span></div>` : ''}
        ${!orderNr && !lieferd && !tatPreis ? '<p class="no-order">Noch keine Bestelldaten.</p>' : ''}
      </div>
      <div class="pf-section" id="panel-attach-section">
        <div class="pf-sec-title">Anhänge</div>
        <div id="panel-attach-body"><div class="vh-loading">Lädt…</div></div>
      </div>
      <div class="pf-section vh-section" id="panel-history-section" style="display:none">
        <div class="pf-sec-title">Versionsverlauf</div>
        <div id="panel-history-body"><div class="vh-loading">Lädt…</div></div>
      </div>
    </div>`;
}

function emptyState(msg) {
  return `<div class="empty-state">
    <svg viewBox="0 0 20 20" fill="currentColor"><path fill-rule="evenodd" d="M5 4a3 3 0 00-3 3v6a3 3 0 003 3h10a3 3 0 003-3V7a3 3 0 00-3-3H5zm-1 9v-1h5v2H5a1 1 0 01-1-1zm7 1h4a1 1 0 001-1v-1h-5v2zm0-4h5V8h-5v2zM9 8H4v2h5V8z" clip-rule="evenodd"/></svg>
    <p>${msg}</p></div>`;
}

// ── MODAL ─────────────────────────────────────────────────────────────────────
function closeModal() { $id('modal-overlay').classList.add('hidden'); }
document.addEventListener('DOMContentLoaded', () => {
  $id('modal-overlay')?.addEventListener('click', e => {
    if (e.target === $id('modal-overlay')) closeModal();
  });
});

// ── ATTACHMENT LINKS ─────────────────────────────────────────────────────────
// Cross-origin fetch to SharePoint is blocked by CORS.
// The user already has an AAD browser session → direct navigation works fine.
function spFullUrl(relUrl) {
  // Encode special chars that break URL parsing (# → %23, ? → %3F, space → %20)
  const encoded = relUrl.replace(/%/g, '%25')   // must be first to avoid double-encoding
                        .replace(/#/g, '%23')
                        .replace(/\?/g, '%3F')
                        .replace(/ /g,  '%20');
  // But if SP already percent-encoded some chars, we'd double-encode — safer: only fix bare #
  // Use a simpler targeted approach:
  const safe = relUrl.split('/').map(seg =>
    seg.replace(/#/g, '%23').replace(/\?/g, '%3F').replace(/ /g, '%20')
  ).join('/');
  return 'https://' + SP_SITE.split(':/')[0] + safe;
}

function openAttachment(relUrl) {
  // Legacy: open directly in new tab
  window.open(spFullUrl(relUrl), '_blank', 'noopener');
}

async function openPdfViewer(relUrl, fileName) {
  const fullUrl = spFullUrl(relUrl);
  const overlay = $id('pdf-viewer-modal');
  const iframe  = $id('pdf-viewer-iframe');
  const titleEl = $id('pdf-viewer-title');
  const extBtn  = $id('pdf-viewer-ext');
  if (!overlay) { window.open(fullUrl, '_blank', 'noopener'); return; }

  // Show modal immediately with loading state
  if (titleEl) titleEl.textContent = (fileName || 'Dokument') + ' – wird geladen…';
  if (extBtn)  extBtn.onclick = () => window.open(fullUrl, '_blank', 'noopener');
  iframe.src = 'about:blank';
  overlay.style.display = 'flex';

  // Fetch file content via SP REST API $value endpoint.
  // Using /_api/ URL instead of direct file URL: avoids X-Frame-Options on direct SP URLs
  // and benefits from CORS headers that SP sets for /_api/ routes with OAuth tokens.
  try {
    const tok = await getSpToken();
    // Decode any %23 → # to get the raw server-relative URL
    const rawRelUrl = decodeURIComponent(relUrl);

    // GetFileByServerRelativeUrl cannot handle '#' in filenames (SP treats it as URL
    // fragment regardless of encoding). Instead use the AttachmentFiles endpoint which
    // addresses files by name via a query-string alias — '#' round-trips correctly there.
    //   /Lists/Bedarfsanfrage/Attachments/{itemId}/{filename}
    const attachMatch = rawRelUrl.match(/\/Attachments\/(\d+)\/(.+)$/);
    let apiUrl;
    if (attachMatch) {
      const itemId  = attachMatch[1];
      const rawName = attachMatch[2]; // may contain #, spaces, etc.
      apiUrl = `${SP_BASE}/_api/web/lists/getByTitle('${SP_LIST}')/items(${itemId})/AttachmentFiles(@f)/$value?@f='${encodeURIComponent(rawName)}'`;
    } else {
      // Fallback for non-attachment URLs
      const spSafe = rawRelUrl.replace(/#/g, '%23');
      apiUrl = `${SP_BASE}/_api/web/GetFileByServerRelativeUrl(@u)/$value?@u='${encodeURIComponent(spSafe)}'`;
    }
    const resp   = await fetch(apiUrl, { headers: { Authorization: 'Bearer ' + tok } });
    if (!resp.ok) throw new Error('HTTP ' + resp.status);
    const blob    = await resp.blob();
    const blobUrl = URL.createObjectURL(blob);
    // Revoke previous blob URL if any
    if (iframe._blobUrl) URL.revokeObjectURL(iframe._blobUrl);
    iframe._blobUrl = blobUrl;
    iframe.src      = blobUrl;
    if (titleEl) titleEl.textContent = fileName || 'Dokument';
  } catch(e) {
    console.warn('[openPdfViewer]', e);
    if (titleEl) titleEl.textContent = fileName || 'Dokument';
    // Fallback: plain message with external open button
    iframe.srcdoc = `<html><body style="font-family:sans-serif;display:flex;align-items:center;
      justify-content:center;height:100%;margin:0;background:#f8fafc;color:#374151">
      <div style="text-align:center;padding:32px">
        <p style="font-size:1rem;margin-bottom:20px">PDF kann nicht direkt angezeigt werden<br>(SP blockiert Einbettung).</p>
        <button onclick="parent.document.getElementById('pdf-viewer-ext').click()"
          style="padding:10px 24px;background:#1a56db;color:#fff;border:none;border-radius:8px;
          font-size:1rem;cursor:pointer">↗ In SharePoint öffnen</button>
      </div></body></html>`;
  }
}

// ── SETTINGS MODAL ───────────────────────────────────────────────────────────
function openSettings() {
  const email   = (account?.username || '').toLowerCase();
  const s       = getSettings(email);
  const isAdmin = email === ADMIN_EMAIL;

  // Admin section row: per-user feature toggles.
  // autoRefreshGranted distinguishes admin-granted from old self-set defaults.
  const userRow = (em, us) => `
    <div class="su-row">
      <span class="su-email" title="${esc(em)}">${esc(em)}</span>
      <label class="tgl-wrap" title="Auto-Refresh freischalten">
        <input type="checkbox" ${us.autoRefresh && us.autoRefreshGranted ? 'checked' : ''}
          onchange="saveUserSettings('${esc(em)}',{autoRefresh:this.checked,autoRefreshGranted:this.checked});updateARBtn()">
        <span class="tgl"></span>
      </label>
      <input type="number" value="${us.pageSize||100}" min="10" max="500" step="10"
        class="su-num" title="Elemente"
        onchange="saveUserSettings('${esc(em)}',{pageSize:parseInt(this.value)||100})">
      <span class="su-lbl">Elem.</span>
    </div>`;

  const adminSection = isAdmin ? `
    <hr class="modal-hr">
    <h4 class="settings-h4">Benutzereinstellungen verwalten</h4>
    <div class="su-add">
      <input type="email" id="su-new-email" placeholder="user@dihag.com" class="su-input">
      <button class="btn btn-sm btn-primary" onclick="addUserSetting()">+ Hinzufügen</button>
    </div>
    <div class="su-header-row">
      <span class="su-email" style="font-size:.7rem;color:#6b7280">E-Mail</span>
      <span class="su-col-lbl" title="Auto-Refresh">⏱</span>
      <span class="su-col-lbl" title="Max. Elemente">Elem.</span>
    </div>
    <div id="su-list">${
      Object.entries(getAllUserSettings())
        .filter(([k]) => k !== email)
        .map(([k, v]) => userRow(k, v)).join('') ||
      '<p class="su-empty">Noch keine weiteren Benutzer konfiguriert.</p>'
    }</div>` : '';

  // Auto-Refresh row: only admin can manage this for themselves and others.
  // Non-admin users see no toggle; admin also saves autoRefreshGranted to survive the
  // login migration that clears old self-set autoRefresh values.
  const arRow = isAdmin ? `
    <div class="settings-row">
      <span class="settings-label">Auto-Aktualisierung (alle 30s)</span>
      <label class="tgl-wrap">
        <input type="checkbox" id="s-ar" ${s.autoRefresh && s.autoRefreshGranted ? 'checked' : ''}
          onchange="saveUserSettings('${esc(email)}',{autoRefresh:this.checked,autoRefreshGranted:this.checked});this.checked?startAutoRefresh():stopAutoRefresh()">
        <span class="tgl"></span>
      </label>
    </div>` : '';

  $id('settings-body').innerHTML = `
    <h4 class="settings-h4">Meine Einstellungen <small>(${esc(email)})</small></h4>
    ${arRow}
    <div class="settings-row">
      <span class="settings-label">Elemente laden (max.)</span>
      <input type="number" id="s-ps" value="${s.pageSize}" min="10" max="500" step="10" class="su-num"
        onchange="saveUserSettings('${esc(email)}',{pageSize:parseInt(this.value)||100})">
    </div>
    ${adminSection}`;
  $id('settings-modal').classList.remove('hidden');
}
function closeSettings() { $id('settings-modal').classList.add('hidden'); }
function addUserSetting() {
  const em = ($id('su-new-email')?.value || '').trim().toLowerCase();
  if (!em || !em.includes('@') || !em.includes('.')) {
    toast('Bitte gültige E-Mail-Adresse eingeben.', 'error'); return;
  }
  const all = JSON.parse(localStorage.getItem(SETTINGS_KEY) || '{}');
  if (all[em]) { toast(`${em} ist bereits vorhanden.`, 'info'); openSettings(); return; }
  // Add with neutral defaults — admin can enable features via the toggles below.
  saveUserSettings(em, { pageSize: 100 });
  toast(`Benutzer ${em} hinzugefügt. Berechtigungen können jetzt gesetzt werden.`, 'success');
  openSettings();
}

function closePdfViewer() {
  const overlay = $id('pdf-viewer-modal');
  const iframe  = $id('pdf-viewer-iframe');
  if (overlay) overlay.style.display = 'none';
  if (iframe) {
    if (iframe._blobUrl) { URL.revokeObjectURL(iframe._blobUrl); iframe._blobUrl = null; }
    iframe.src    = 'about:blank';
    iframe.srcdoc = '';
  }
}

function attachmentLink(f) {
  const safeUrl = esc(f.ServerRelativeUrl.replace(/#/g,'%23').replace(/\?/g,'%3F').replace(/ /g,'%20'));
  const name    = esc(f.FileName);
  return `<div class="attach-item">
    📎 <span class="attach-name">${name}</span>
    <span class="attach-actions">
      <button class="btn-attach-dl" onclick="openAttachment('${safeUrl}')" title="In SharePoint öffnen">↗ Öffnen</button>
    </span>
  </div>`;
}

// ── TOAST ─────────────────────────────────────────────────────────────────────
function toast(msg, type='info') {
  const icons = { success:'✅', error:'❌', info:'ℹ️' };
  const t = document.createElement('div');
  t.className = 'toast ' + type;
  t.innerHTML = `<span>${icons[type]||'ℹ️'}</span><span>${esc(msg)}</span>`;
  $id('toast-c').appendChild(t);
  setTimeout(() => { t.classList.add('out'); setTimeout(()=>t.remove(),260); }, 4000);
}

// ── HELPERS ──────────────────────────────────────────────────────────────────
const $id = id => document.getElementById(id);
const esc = s  => String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');

function getField(item, key) {
  if (!key || !item?.fields) return null;
  return item.fields[key] ?? null;
}

function cleanTitle(t) {
  if (!t) return '–';
  return t
    .replace(/^Einkauf\s*[-–]\s*BANF\s*#\d+\s*[-–]\s*/i, '')
    .replace(/\s*[-–]\s*[\d.,]+\s*€\s*$/i, '')
    .trim() || t;
}

// Show status badge + current approver (if a column matching APPROVER_COL_RE exists)
function statusWithApprover(item) {
  const approver = getApproverVal(item);
  const badge    = statusBadge(getStatusVal(item));
  if (!approver) return badge;
  return `${badge}<span class="status-approver" title="Aktueller Genehmiger">👤 ${esc(approver)}</span>`;
}

function statusBadge(s) {
  const label = s || 'Eingereicht';
  const sl = label.toLowerCase().trim();
  let st = STATUS_STYLES[sl];
  if (!st) {
    for (const [key, style] of Object.entries(STATUS_STYLES)) {
      if (sl.includes(key)) { st = style; break; }
    }
  }
  st = st || { bg:'#f3f4f6', color:'#374151' };
  let icon = '';
  if      (sl.includes('abgelehnt') || sl.includes('rejected'))               icon = '✗ ';
  else if (sl.includes('freigegeben') || sl.includes('bestellt') || sl.includes('erledigt')) icon = '✓ ';
  else if (sl.includes('prüfung') || sl.includes('bearbeitung'))               icon = '⏳ ';
  else if (sl.includes('eingereicht') || sl.includes('angefragt'))             icon = '📋 ';
  return `<span class="status-badge" style="background:${st.bg};color:${st.color}">${icon}${esc(label)}</span>`;
}

function prioTag(p) {
  if (!p || p.toLowerCase() === 'normal') return '';
  const color = p.toLowerCase() === 'dringend' ? '#b91c1c' : '#b45309';
  const bg    = p.toLowerCase() === 'dringend' ? '#fef2f2' : '#fffbeb';
  return `<span class="status-badge" style="background:${bg};color:${color}">${esc(p)}</span>`;
}

function prioDot(p) {
  const pl = (p || '').toLowerCase();
  let c;
  if      (pl.includes('hoch') || pl.includes('dringend')) c = '#ef4444';
  else if (pl.includes('mittel'))                          c = '#f59e0b';
  else                                                     c = '#22c55e';
  return `<span style="display:inline-block;width:7px;height:7px;border-radius:50%;background:${c};margin-right:5px;vertical-align:middle;flex-shrink:0"></span>`;
}

function fmtEuro(v) {
  if (!v && v !== 0) return '';
  return new Intl.NumberFormat('de-DE',{style:'currency',currency:'EUR'}).format(v);
}
function fmtDate(s) {
  if (!s) return '';
  return new Date(s).toLocaleDateString('de-DE',{day:'2-digit',month:'2-digit',year:'numeric'});
}
function toLocalInputDate(s) {
  if (!s) return '';
  const d = new Date(s);
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}
function fmtDateTime(s) {
  if (!s) return '';
  return new Date(s).toLocaleString('de-DE',{day:'2-digit',month:'2-digit',year:'numeric',hour:'2-digit',minute:'2-digit'});
}
function fmtRelative(s) {
  const d = Date.now() - new Date(s).getTime();
  if (d < 60000)     return 'gerade eben';
  if (d < 3600000)   return Math.floor(d/60000) + ' Min.';
  if (d < 86400000)  return Math.floor(d/3600000) + ' Std.';
  if (d < 2592000000)return Math.floor(d/86400000) + ' Tage';
  return fmtDate(s);
}

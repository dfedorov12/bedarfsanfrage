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
  { key:'Title',                 label:'Bezeichnung',                    step:1, required:true  },
  { key:'Beschreibung',          label:'Beschreibung',                   step:1, alsoTry:['Description','Beschreibung_x002f_Begruendung','Grund'] },
  { key:'Warengruppe',           label:'Warengruppe',                    step:1, required:true, alsoTry:['ProductCategory'] },
  { key:'Prioritaet',            label:'Priorität',                      step:1, alsoTry:['Priority','Priorit_x00e4_t'] },
  // Step 2: Menge
  { key:'Menge',                 label:'Menge',                          step:2, required:true, alsoTry:['Quantity','Amount'] },
  { key:'Mengeneinheit',         label:'Mengeneinheit',                  step:2, required:true, alsoTry:['Unit','UnitOfMeasure'] },
  { key:'Mindestlagermenge',     label:'Mindestlagermenge',              step:2, alsoTry:['MinStock','MinLager'] },
  { key:'Termin',                label:'Benötigt bis',                   step:2, required:true, alsoTry:['Deadline','DueDate','Ben_x00f6_tigtBis','Ben_x00f6_tigtbis'] },
  // Step 3: Beschaffung
  { key:'Artikelnummer',         label:'Artikelnummer',                  step:1, alsoTry:['MaterialNumber','ItemNumber','Artikelnummer_x002f_Nummernangab'] },
  { key:'ExterneArtikelnummer',  label:'Externe Artikelnummer',          step:1, alsoTry:['ExternalItemNumber','ExtArtikelNr','ExterneArtikelnr'] },
  { key:'Beschaffungslogik',     label:'Beschaffungsart',                step:3, required:true, alsoTry:['Materialtyp','ProcurementType'] },
  { key:'Lieferant',             label:'Lieferant 1',                    step:3, alsoTry:['Vendor','Supplier'] },
  { key:'Lieferant2',            label:'Lieferant 2 (Alternative)',      step:3, alsoTry:['Vendor2','Supplier2','Lieferant_2'] },
  { key:'Lieferant3',            label:'Lieferant 3 (Alternative)',      step:3, alsoTry:['Vendor3','Supplier3','Lieferant_3'] },
  { key:'Lieferant4',            label:'Lieferant 4 (Alternative)',      step:3, alsoTry:['Vendor4','Supplier4','Lieferant_4'] },
  { key:'GeschaetzterPreis',     label:'Bestellvolumen in €',            step:3, alsoTry:['EstimatedPrice','Preis','Price','Gesch_x00e4_tzterPreisnetto_x002'] },
  { key:'Kostenstelle',          label:'Kostenstelle',                   step:3, alsoTry:['CostCenter'] },
  { key:'LeadBuyerAbschluss',    label:'Lead-Buyer-Abschluss',           step:3, alsoTry:['LeadBuyer','LeadBuyerAbschlus'] },
  { key:'Positionen',            label:'Positionen (JSON)',               step:1, alsoTry:['Positions'] },
];

// Felder die Einkauf nach der Einreichung befüllt
const EINKAUF_FIELDS = [
  { key:'Bestellnummer',  label:'Bestellnummer',        alsoTry:['OrderNumber','PO_Number'] },
  { key:'Lieferdatum',    label:'Lieferdatum',          alsoTry:['DeliveryDate'] },
  { key:'TatsaechlicherPreis', label:'Tatsächlicher Preis (€)', alsoTry:['ActualPrice','FinalPrice'] },
];

// Kommentarfeld — separates SP-Listenfeld "Kommentare" (Mehrere Textzeilen)
// Format jedes Eintrags: "[DD.MM.YYYY HH:MM – Autor]: Text\n---\n"
const KOMMENTAR_FIELD = { key: 'Kommentare', alsoTry: ['Comments', 'Kommentar'] };

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


// ── KOSTENSTELLEN-STAMMDATEN ────────────────────────────────────────────────
// Format: { nr: "10011", label: "Fertigungsleitung" }
// Overrideable via localStorage key DIHAG_KOSTENST
const KOSTENST_DEFAULT = [
  {nr:'10011',label:'Fertigungsleitung'},
  {nr:'10051',label:'Arbeitsvorbereitung/Fertigungssteuerung'},
  {nr:'11011',label:'Qualitätssicherung'},
  {nr:'11021',label:'Qualitätsprüfung'},
  {nr:'1111', label:'Produktionskomplex'},
  {nr:'1115', label:'Rückstellung Personalaufwand'},
  {nr:'1211', label:'Wachdienst'},
  {nr:'1221', label:'Wasch- und Umkleideräume'},
  {nr:'1223', label:'Telefon- und Computeranlage'},
  {nr:'12311',label:'Modellbau'},
  {nr:'20011',label:'Einkauf und Eingangsfrachten'},
  {nr:'20012',label:'Lagerwirtschaft'},
  {nr:'2011', label:'Stromversorgung'},
  {nr:'2021', label:'Erdgasversorgung'},
  {nr:'2031', label:'Wasserversorgung'},
  {nr:'2111', label:'Acetylen'},
  {nr:'2112', label:'Sauerstoff'},
  {nr:'2113', label:'Kohlendioxyd'},
  {nr:'2211', label:'Druckluftversorgung'},
  {nr:'2421', label:'Wärmeversorgung'},
  {nr:'3011', label:'Fuhrpark'},
  {nr:'3021', label:'Innerbetrieblicher Transport'},
  {nr:'31011',label:'Lichtbogenöfen'},
  {nr:'31025',label:'MF-Ofen'},
  {nr:'31031',label:'Gießen einschließlich Pfannenwirtschaft'},
  {nr:'32011',label:'Handkernmacherei'},
  {nr:'32012',label:'Maschinenkernmacherei'},
  {nr:'33011',label:'Formanlage HWS'},
  {nr:'33012',label:'Formstoffaufbereitung FA HWS'},
  {nr:'33021',label:'Handformerei Furanharz'},
  {nr:'33031',label:'Handformerei Alphaset'},
  {nr:'33041',label:'Mechanisierte Handformerei'},
  {nr:'33042',label:'Sandaufbereitung Alphaset'},
  {nr:'36011',label:'Strahlanlagen Stahlkies'},
  {nr:'36014',label:'Umlaufhängebahnstrahlanlage Typ 13 U/III'},
  {nr:'36015',label:'Wheelabrator Strahlanlagen'},
  {nr:'36021',label:'Brennen'},
  {nr:'36031',label:'Trennmaschine'},
  {nr:'36042',label:'Putzen / Pendeln'},
  {nr:'36043',label:'Putzzentrum Bodenplatten'},
  {nr:'36044',label:'Farbspritzanlage'},
  {nr:'36051',label:'Schweißen'},
  {nr:'36061',label:'HWO Schweißvorwärmen Großteileputzerei'},
  {nr:'36062',label:'Großteileputzerei'},
  {nr:'38011',label:'Qualitätsprüfung'},
  {nr:'51011',label:'Kammerofenanlage'},
  {nr:'53011',label:'Herdwagenofen'},
  {nr:'54011',label:'MF-Induktionshärteanlagen'},
  {nr:'54031',label:'MF-Härtemaschinen Unitherma'},
  {nr:'6011', label:'Instandhaltung'},
  {nr:'61011',label:'Karuselldrehmaschinen NC'},
  {nr:'61021',label:'Karuselldrehmaschinen konvent.'},
  {nr:'61031',label:'Fertigungszellen Body'},
  {nr:'61041',label:'Bearbeitungszentrum CWK 1000'},
  {nr:'61051',label:'Bohr-u.Fräswerke BFT 110'},
  {nr:'62011',label:'Schakenfertigung'},
  {nr:'62021',label:'Formzeugfertigung'},
  {nr:'64011',label:'Fräs-/Dreh-/Karusselldreh-/Säge-/Trennmaschinen'},
  {nr:'65011',label:'Konstruktionsschweißen und Baugruppenkomponenten'},
  {nr:'70011',label:'Verwaltung'},
  {nr:'70021',label:'Vertrieb'},
  {nr:'7011', label:'Betriebsrat'},
  {nr:'95011',label:'Einzelkosten'},
  {nr:'95012',label:'Sondereinzelkosten Vertrieb'},
  {nr:'98012',label:'Modellkosten'},
  {nr:'98211',label:'Nacharbeit u. Garantieleistungen'},
  {nr:'98311',label:'Versuchsproduktion'},
];

// Active lookup tables — reloaded from localStorage by reloadLookupData()
let TID_MAP_ACTIVE = TID_MAP;         // may be replaced by imported data
let TID_ENTRIES    = Object.entries(TID_MAP);
let KOSTENST_DATA  = KOSTENST_DEFAULT; // may be replaced by imported data

const LS_TID_KEY    = 'DIHAG_TID_DATA';
const LS_KOSTENST_KEY = 'DIHAG_KOSTENST';

function reloadLookupData() {
  try {
    const rawTid = localStorage.getItem(LS_TID_KEY);
    if (rawTid) {
      const obj = JSON.parse(rawTid);
      TID_MAP_ACTIVE = Object.assign({}, TID_MAP, obj); // merge: imported wins over default for same keys
      TID_ENTRIES = Object.entries(TID_MAP_ACTIVE);
    } else {
      TID_MAP_ACTIVE = TID_MAP;
      TID_ENTRIES = Object.entries(TID_MAP);
    }
  } catch(e) { console.warn('reloadLookupData TID:', e); }
  try {
    const rawKst = localStorage.getItem(LS_KOSTENST_KEY);
    if (rawKst) KOSTENST_DATA = JSON.parse(rawKst);
    else        KOSTENST_DATA = KOSTENST_DEFAULT;
  } catch(e) { console.warn('reloadLookupData KOSTENST:', e); }
}

// ── STATE ───────────────────────────────────────────────────────────────────
let msalApp, account;
let siteId = null, listId = null;
let allItems = [];
let colByKey   = {};  // internal name → column definition (from SP)
let spUserMap  = {};  // SP user id (string) → display name (for Person-column LookupId resolution)
let resolvedFields = {};  // FORM_FIELDS key → actual SP internal name (null if not found)
let statusChoices = []; // All SP Status column choices in order (populated in discoverSP)

// ── EDIT LOCK ─────────────────────────────────────────────────────────────────
// Stores a "user|ISO-expires" value in a SP list field to prevent concurrent edits.
const LOCK_FIELD_CANDIDATES = ['BearbeitungsSperre', 'EditLock', 'Bearbeitungssperre'];
let _lockField  = undefined; // undefined = not yet resolved; '' = field not in SP
let lockTimers  = {};        // itemId (string) → timeout handle
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
  return Object.assign({
    autoRefresh:        true,
    autoRefreshGranted: true,
    pageSize:           100,
    compactView:        false,   // dense single-line cards in list views
    hideCompleted:      false,   // hide bestellt/erledigt/abgelehnt in Meine Anfragen
    defaultSort:        'date-desc', // default sort for Meine Anfragen
    canSeeDashboard:    false,   // admin-granted: show Dashboard tab (all items visible)
  }, all[(email||'').toLowerCase()] || {});
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
// SharePoint list/drive creation requires Manage.All permissions the app does
// not have. Instead we use a token-based grant system:
//   Admin  → generates a compact base64 token from current user grants
//          → copies it and sends to the user (e.g. via Teams/Email)
//   User   → pastes token in Settings modal → grants applied locally
// localStorage is the runtime cache; token is the out-of-band sync channel.
const SP_REST = 'https://' + SP_SITE.replace(':/', '/') + '/_api/web';
let _cfgItemId = null;

// SP REST: common JSON headers (Bearer token makes X-RequestDigest unnecessary)
function _spHdr(tok, extra = {}) {
  return Object.assign({
    Authorization: 'Bearer ' + tok,
    Accept:        'application/json;odata=nometadata',
    'Content-Type':'application/json;odata=verbose',
  }, extra);
}

// No-op: SP sync removed (insufficient permissions to create lists or write drives).
// Grants are distributed via copy-paste tokens (see openSettings / applyGrantToken).
async function loadSpSettings()    { /* token-based, no remote fetch needed */ }
async function persistSpSettings() { /* token-based, no remote write needed */ }

// ── AUTO-REFRESH ─────────────────────────────────────────────────────────────
// autoRefreshTimer: 1-second tick; arCountdown: seconds until next refresh
let autoRefreshTimer = null;
let arCountdown = 30;
let arPaused = false;   // user can pause without admin losing the feature-enable

function startAutoRefresh() {
  stopAutoRefresh();
  if (!account) return;
  arCountdown = 20;
  arPaused    = false;
  autoRefreshTimer = setInterval(() => {
    if (arPaused) return;
    arCountdown--;
    if (arCountdown <= 0) {
      arCountdown = 20;
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
  if (!account) return;
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
    applyNavVisibility(); // re-apply after SP settings loaded (grants may have changed)
    // Log what the current user's effective settings are — visible in browser console
    if (account) {
      const _em = (account.username || '').toLowerCase();
      const _s  = getSettings(_em);
      console.log('[bootDone] User:', _em, '| canSeeDashboard:', _s.canSeeDashboard, '| autoRefresh:', _s.autoRefresh);
    }
    startAutoRefresh();
    applyDashboardVisibility();
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
  // Store the real internal name so all readers (loadVersionHistory etc.) use the correct key
  if (statusCol) resolvedFields['Status'] = statusCol.name;

  if (statusCol?.choice?.choices?.length) {
    const choices = statusCol.choice.choices;
    statusChoices = choices; // store all choices (incl. end states) for the timeline
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
  const map = { new:'Neue Bedarfsanfrage', multi:'Sammelanfrage', mine:'Meine Anfragen', all:'Alle Anfragen', detail:'Anfrage Details', importer:'Tabellen-Importer', reports:'Reports & Auswertungen', schulung:'Schulung (Beta)' };
  if (view === 'dashboard') return isAdmin() ? 'Dashboard (Alle Anfragen)' : 'Meine Anfragen';
  return map[view] || view;
}

// Dashboard is visible to all logged-in users.
// Non-admins see only their own items; admins see all items.
function canSeeDashboard() { return !!account; }
function isAdmin() { return account?.username?.toLowerCase() === ADMIN_EMAIL; }

// Show/hide nav items based on role and admin-granted permissions.
function applyNavVisibility() {
  const email    = (account?.username || '').toLowerCase();
  const s        = getSettings(email);
  const dashNav  = document.querySelector('.nav-item[data-view="dashboard"]');
  if (dashNav) dashNav.style.display = '';
}
// Alias kept for any remaining call-sites
const applyDashboardVisibility = applyNavVisibility;

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

  // Pause auto-refresh while filling in a form; resume when leaving
  if (view === 'new' || view === 'multi') {
    arPaused = true; updateARBtn();
  } else {
    if (arPaused) { arPaused = false; if (!autoRefreshTimer) startAutoRefresh(); updateARBtn(); }
  }

  if      (view === 'dashboard') renderDashboard();
  else if (view === 'mine')      renderList('mine');
  else if (view === 'all')       renderList('all');
  else if (view === 'new')       initWizard();
  else if (view === 'multi')     initMultiWizard();
  else if (view === 'importer')  initImporter();
  else if (view === 'reports')   initReports();
  else if (view === 'schulung')  initSchulung();
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

  // Boot — load persisted lookup table overrides (TID + Kostenstellen)
  reloadLookupData();

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
let dashWGFilter = '';
let dashMAFilter = '';

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

  const baseItems = allItems;

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
  renderExtraFilters();
}

function setDashFilter(status) {
  dashStatusFilter = (dashStatusFilter === status) ? '' : status;
  renderStatusChips();
  filterDashboard();
}

function renderExtraFilters() {
  const el = $id('dash-extra-filters');
  if (!el) return;

  // Unique WG values
  const wgField = resolvedFields['Warengruppe'] || 'Warengruppe';
  const wgSet = new Set();
  for (const i of allItems) {
    const v = getField(i, wgField) || '';
    if (v) wgSet.add(v);
  }
  const wgOpts = [...wgSet].sort().map(v =>
    `<option value="${esc(v)}"${dashWGFilter === v ? ' selected' : ''}>${esc(v)}</option>`
  ).join('');

  // Unique Mitarbeiter values
  const maSet = new Set();
  for (const i of allItems) {
    const v = i.createdBy?.user?.displayName || i.createdBy?.user?.email || '';
    if (v) maSet.add(v);
  }
  const maOpts = [...maSet].sort().map(v =>
    `<option value="${esc(v)}"${dashMAFilter === v ? ' selected' : ''}>${esc(v)}</option>`
  ).join('');

  el.innerHTML = `
    <select class="dash-filter-select" onchange="setDashWGFilter(this.value)">
      <option value="">Alle Warengruppen</option>
      ${wgOpts}
    </select>
    <select class="dash-filter-select" onchange="setDashMAFilter(this.value)">
      <option value="">Alle Mitarbeiter</option>
      ${maOpts}
    </select>`;
}

function setDashWGFilter(val) {
  dashWGFilter = val;
  filterDashboard();
}

function setDashMAFilter(val) {
  dashMAFilter = val;
  filterDashboard();
}

function filterDashboard() {
  const search = ($id('search-dashboard')?.value || '').toLowerCase();
  let items = [...allItems];
  if (search) items = items.filter(i =>
    (getField(i,'Title')||'').toLowerCase().includes(search) || String(i.id||'').includes(search)
  );
  if (dashStatusFilter !== '') items = items.filter(i => (getStatusVal(i)||'') === dashStatusFilter);
  if (dashWGFilter) items = items.filter(i => (getField(i, resolvedFields['Warengruppe']||'Warengruppe')||'') === dashWGFilter);
  if (dashMAFilter) items = items.filter(i => {
    const creator = i.createdBy?.user?.displayName || i.createdBy?.user?.email || '';
    return creator === dashMAFilter;
  });

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
let mineStatusFilter  = '';
let _mineInitialized  = false; // so defaultSort only applies on first load
let mineSortOrder    = 'date-desc';

function myItems() {
  const myEmail = (account?.username || '').toLowerCase();
  return allItems.filter(i =>
    (i.createdBy?.user?.email || '').toLowerCase() === myEmail ||
    (i.createdBy?.user?.displayName || '').toLowerCase() === (account?.name || '').toLowerCase()
  );
}

function renderList(type) {
  if (type === 'mine') {
    // Apply user's defaultSort on first load (not on every re-render so user can override)
    if (account) {
      const s = getSettings(account.username);
      if (!_mineInitialized) { mineSortOrder = s.defaultSort || 'date-desc'; _mineInitialized = true; }
    }
    renderStatusChipsMine();
    filterView('mine');
    return;
  }
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

  const userSettings = account ? getSettings(account.username) : {};
  let items = type === 'mine' ? myItems() : [...allItems];

  // hideCompleted: filter out terminal states in Meine Anfragen
  if (type === 'mine' && userSettings.hideCompleted && !mineStatusFilter) {
    items = items.filter(i => !/^(bestellt|erledigt|abgelehnt)$/i.test(getStatusVal(i)));
  }

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
  if (container) {
    // compactView: toggle dense layout class
    container.classList.toggle('compact-list', !!(userSettings.compactView));
    container.innerHTML = items.length
      ? items.map(i => itemCard(i)).join('')
      : emptyState(type === 'mine' ? 'Sie haben noch keine Anfragen erstellt.' : 'Keine Anfragen gefunden.');
  }
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

// Timeline stages with direct SP field mappings (from Power Automate workflow).
// approverField / commentField are read directly from the list item — no version
// history needed.  A stage is "visible" if its approverField is non-empty
// (meaning that stage was actually reached) or it is the current status.
const TIMELINE_STAGES = [
  { label: 'Eingereicht',
    test: v => /^eingereicht$/i.test(v),
    approverField: null, commentField: null },
  { label: 'In Prüfung (Einkauf)',
    test: v => /pr[üu]fung/i.test(v) && /einkauf/i.test(v) && !/strategisch/i.test(v),
    approverField: 'Entscheider_x002a_in', commentField: 'Genehmigungskommentar' },
  { label: 'Genehmigt (Einkauf)',
    test: v => /genehmigt/i.test(v) && /einkauf/i.test(v),
    approverField: 'Entscheider_x002a_in', commentField: 'Genehmigungskommentar' },
  { label: 'In Prüfung (Werkleitung)',
    test: v => /werkleitung/i.test(v),
    approverField: 'Genehmiger3', commentField: 'Genehmigungskommentar3' },
  { label: 'In Prüfung (Controlling)',
    test: v => /controlling/i.test(v),
    approverField: 'Genehmiger2', commentField: 'Genehmigungskommentar2' },
  { label: 'In Prüfung (strategischer Einkauf)',
    test: v => /strategisch/i.test(v),
    approverField: 'Genehmiger4', commentField: 'Genehmigungskommentar4' },
  { label: 'Freigegeben',
    test: v => /^freigegeben$/i.test(v),
    approverField: null, commentField: null },
  { label: 'In Bestellung',
    test: v => /in bestellung/i.test(v),
    approverField: null, commentField: null },
  { label: 'Bestellt',
    test: v => /^bestellt$/i.test(v),
    approverField: null, commentField: null },
  { label: 'Abgelehnt',
    test: v => /^abgelehnt$/i.test(v),
    approverField: null, commentField: null },
];

// Resolve a person/lookup field from item.fields to a display name string.
function resolvePersonField(fields, fieldName) {
  if (!fieldName || !fields) return null;
  function tryExtract(raw) {
    if (raw == null || raw === '') return null;
    if (typeof raw === 'object') {
      const n = raw.displayName || raw.LookupValue || raw.Title || raw.title;
      return (n && n !== '[object Object]') ? String(n).trim() || null : null;
    }
    const s = String(raw).trim();
    if (!s || s === '[object Object]') return null;
    if (/^\d+$/.test(s)) return spUserMap[s] || null; // numeric LookupId
    return s;
  }
  return tryExtract(fields[fieldName])
      || tryExtract(spUserMap[String(fields[fieldName + 'LookupId'] ?? '')])
      || null;
}

// Workflow-Reihenfolge für die Status-Zeitleiste (wird in discoverSP() aus SP-Spalte befüllt)
const WORKFLOW_STAGES = [
  'Eingereicht',
  'In Prüfung (Einkauf)',
  'In Prüfung (Werkleitung)',
  'In Prüfung (Controlling)',
  'In Prüfung (strategischer Einkauf)',
  'Freigegeben',
  'In Bestellung',
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
// Iterates ONLY the real SP Status column choices (statusChoices, populated from SP).
// TIMELINE_STAGES is used purely as a lookup for approverField/commentField names.
// Falls back to TIMELINE_STAGES labels if SP choices haven't loaded yet.
function statusTimeline(statusVal, item) {
  const sv     = (statusVal || '').trim();
  const fields = item?.fields || item || {};

  // Source: real SP choices only, deduplicated (SP list sometimes has duplicate entries).
  // Fallback to [sv] if not yet loaded — never use hardcoded TIMELINE_STAGES labels.
  const seen   = new Set();
  const source = (statusChoices.length ? statusChoices : [sv])
    .filter(c => { const k = c.trim().toLowerCase(); return k && !seen.has(k) && seen.add(k); });

  const svLow            = sv.toLowerCase();
  const isRej            = /^abgelehnt$/i.test(sv);
  const isBestelltNow    = /^bestellt$/i.test(sv);
  const isFreigegebenNow = /^freigegeben$/i.test(sv);
  const TERMINAL_OK      = /^(freigegeben|bestellt)$/i;
  const IN_BESTELLG      = /^in bestellung$/i;

  // Sort source by logical workflow order (TIMELINE_STAGES index), not SP choice order.
  // Stages not found in TIMELINE_STAGES (unknown/custom) go at the end.
  const stageIndex = cv => {
    const idx = TIMELINE_STAGES.findIndex(d => d.test(cv.trim()));
    return idx === -1 ? 999 : idx;
  };
  const sortedSource = [...source].sort((a, b) => stageIndex(a) - stageIndex(b));

  const rows = sortedSource.map(choiceVal => {
    const cv        = choiceVal.trim();
    const isCurrent = cv.toLowerCase() === svLow;

    // Look up approver/comment field names from TIMELINE_STAGES
    const tsd      = TIMELINE_STAGES.find(d => d.test(cv));
    const approver = resolvePersonField(fields, tsd?.approverField);
    const comment  = tsd?.commentField ? String(fields[tsd.commentField] || '').trim() : '';

    const isInBestellungFuture = IN_BESTELLG.test(cv) && isFreigegebenNow;
    const isInBestellungPast   = IN_BESTELLG.test(cv) && isBestelltNow;

    // Visibility rules — NO index-based "past" (index order is meaningless for
    // branching workflows; "Abgelehnt" sits after "Freigegeben" in SP list order
    // which would falsely mark it as reached for "Bestellt" items, etc.)
    //
    // A stage is shown when:
    //   • It is the current status                   → always
    //   • It is "Eingereicht"                        → always (first step, always happened)
    //   • Its approverField was filled by workflow   → stage was actually reached
    //     EXCEPT: skip "Genehmigt"-labelled stages for rejected items
    //     (item was rejected, not approved, even if the field was filled earlier)
    //   • "In Bestellung" special cases
    const isGenehmigt = /genehmigt/i.test(cv);
    const visible = isCurrent
      || /^eingereicht$/i.test(cv)
      || isInBestellungFuture
      || isInBestellungPast
      || (approver != null && !(isRej && isGenehmigt));

    if (!visible) return null;

    let dot, cls;
    if (isCurrent) {
      if      (TERMINAL_OK.test(cv))  { dot = '✓'; cls = 'ap-ok';     }
      else if (IN_BESTELLG.test(cv))  { dot = '○'; cls = 'ap-circle'; }
      else if (isRej)                 { dot = '✗'; cls = 'ap-no';     }
      else                            { dot = '●'; cls = 'ap-pending';}
    } else if (isInBestellungFuture) {
      dot = '○'; cls = 'ap-circle-future';
    } else {
      dot = '✓'; cls = 'ap-ok'; // past stage
    }

    let approverHtml = '';
    if (approver) {
      const commentHtml = comment
        ? `<div class="ap-inline-comment">💬 ${esc(comment)}</div>` : '';
      const pastCls = isCurrent ? '' : ' ap-approver-past';
      approverHtml = `<div class="ap-approver${pastCls}">👤 ${esc(approver)}${commentHtml}</div>`;
    }

    const isActive = isCurrent && !TERMINAL_OK.test(cv) && !IN_BESTELLG.test(cv) && !isRej;
    const bold = isActive ? ' style="font-weight:600"' : '';
    return `<div class="approval-stage"><div class="ap-dot ${cls}">${dot}</div>`
         + `<div class="ap-body"><div class="ap-stage-label"${bold}>${esc(cv)}</div>${approverHtml}</div></div>`;
  }).filter(Boolean);

  return rows.join('');
}

function approvalStyle(val) {
  const v = (val || '').toLowerCase();
  if (/freigegeben|genehmigt|approved|ja\b/.test(v)) return { bg:'#f0fdf4', color:'#15803d', dot:'✓', cls:'ap-ok' };
  if (/abgelehnt|rejected|nein\b/.test(v))           return { bg:'#fef2f2', color:'#b91c1c', dot:'✗', cls:'ap-no' };
  return { bg:'#fffbeb', color:'#b45309', dot:'…', cls:'ap-pending' };
}

function renderApprovalCard(item) {
  // Delegate to buildApprovalInner so detail view and panel share identical logic
  // (same timeline, same inline comments, same approver placement).
  const statusVal = getStatusVal(item) || 'Eingereicht';
  return `
    <div class="detail-card">
      <div class="detail-card-header">Status &amp; Genehmigung</div>
      <div class="detail-card-body">
        <div class="ap-current-status">${statusWithApprover(item)}</div>
        ${buildApprovalInner(item, statusVal)}
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
      <input type="text" id="m-ordernr" value="${esc(orderNr)}" placeholder="z. B. BE252093"
        oninput="checkBeDuplicate(this.value, ${itemId})"/>
      <div id="be-warn" class="be-warn" style="display:none"></div>
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
  // Initiale Prüfung, falls bereits eine BE eingetragen ist
  checkBeDuplicate(orderNr, itemId);
}

// Warnt, wenn die eingegebene Bestellnummer bereits bei einem anderen Item verwendet wird
function checkBeDuplicate(val, currentId) {
  const el = $id('be-warn');
  if (!el) return;
  const v = (val || '').trim().toLowerCase();
  if (!v) { el.style.display = 'none'; el.innerHTML = ''; return; }
  const beCol = resolvedFields['Bestellnummer'] || 'Bestellnummer';
  const dupes = allItems.filter(i =>
    String(i.id) !== String(currentId) &&
    (getField(i, beCol) || '').trim().toLowerCase() === v
  );
  if (dupes.length) {
    el.style.display = 'block';
    const refs = dupes.map(d => {
      const t = cleanTitle(getField(d, 'Title') || '');
      return `#${d.id}${t ? ' – ' + esc(t) : ''}`;
    }).join(', ');
    el.innerHTML = `⚠️ Achtung: Bestellnummer „${esc(val.trim())}" wird bereits verwendet (${refs}).`;
  } else {
    el.style.display = 'none';
    el.innerHTML = '';
  }
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
  // Wenn Bestellnummer vergeben → Status auf "Bestellt" setzen
  if (orderNr) {
    const statusCol = resolvedFields['Status'] || 'Status';
    if (statusCol) patch[statusCol] = 'Bestellt';
  }

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
      <label>Bestellvolumen in €</label>
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

// ── ARTIKELNUMMER AUTOCOMPLETE ────────────────────────────────────────────────
// Shared between wizard (lookupTID) and panel edit mode (initTidAutocomplete).
// Searches TID_MAP_ACTIVE by number prefix OR name substring.

function initTidAutocomplete(inputEl, getRelatedInput) {
  // Dropdown appended to <body> with position:fixed to escape panel overflow clipping
  let dropdown = null;
  let activeIdx = -1;

  function getOrCreateDropdown() {
    if (!dropdown) {
      dropdown = document.createElement('div');
      dropdown.className = 'tid-ac-dropdown';
      dropdown.style.cssText = 'display:none;position:fixed;z-index:9999';
      document.body.appendChild(dropdown);
    }
    return dropdown;
  }

  function reposition() {
    if (!dropdown || dropdown.style.display === 'none') return;
    const r = inputEl.getBoundingClientRect();
    dropdown.style.left  = r.left + 'px';
    dropdown.style.top   = (r.bottom + 2) + 'px';
    dropdown.style.width = r.width + 'px';
  }

  function hide() {
    if (dropdown) dropdown.style.display = 'none';
    activeIdx = -1;
  }

  function show(items) {
    if (!items.length) { hide(); return; }
    activeIdx = -1;
    const dd = getOrCreateDropdown();
    dd.innerHTML = items.slice(0, 12).map(([ tid, h ], i) =>
      `<div class="tid-ac-item" data-tid="${esc(tid)}" data-idx="${i}">
        <span class="tid-ac-nr">${esc(tid)}</span>
        <span class="tid-ac-name">${esc(h.b)}</span>
        <span class="tid-ac-wg">${esc(h.w)}</span>
      </div>`
    ).join('');
    dd.querySelectorAll('.tid-ac-item').forEach(el => {
      el.addEventListener('mousedown', e => { e.preventDefault(); selectItem(el.dataset.tid); });
    });
    dd.style.display = 'block';
    reposition();
  }

  function selectItem(tid) {
    const hit = TID_MAP_ACTIVE[tid];
    if (!hit) return;
    inputEl.value = tid;
    const titleEl = getRelatedInput?.('Title');
    const wgEl    = getRelatedInput?.('Warengruppe');
    if (titleEl && !titleEl.dataset?.manual) titleEl.value = hit.b;
    if (wgEl) {
      let found = false;
      for (const opt of (wgEl.options || [])) {
        if (opt.value === hit.w) { wgEl.value = hit.w; found = true; break; }
      }
      if (!found && wgEl.tagName === 'SELECT') {
        wgEl.add(new Option(hit.w, hit.w)); wgEl.value = hit.w;
      } else if (wgEl.tagName === 'INPUT') { wgEl.value = hit.w; }
    }
    const confirm = inputEl.closest('.tid-ac-wrap')?.querySelector('.tid-ac-confirm');
    if (confirm) confirm.textContent = '✓ ' + hit.b + ' · ' + hit.w;
    // Also update wizard tid-match if present
    if (typeof lookupTID === 'function' && $id('tid-match')) lookupTID(tid);
    hide();
  }

  inputEl.addEventListener('input', () => {
    const q = inputEl.value.trim().toLowerCase();
    if (q.length < 2) { hide(); return; }
    const matches = TID_ENTRIES.filter(([tid, h]) =>
      tid.toLowerCase().includes(q) || h.b.toLowerCase().includes(q)
    );
    show(matches);
  });

  inputEl.addEventListener('keydown', e => {
    if (!dropdown || dropdown.style.display === 'none') return;
    const items = dropdown.querySelectorAll('.tid-ac-item');
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      activeIdx = Math.min(activeIdx + 1, items.length - 1);
      items.forEach((el, i) => el.classList.toggle('tid-ac-active', i === activeIdx));
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      activeIdx = Math.max(activeIdx - 1, 0);
      items.forEach((el, i) => el.classList.toggle('tid-ac-active', i === activeIdx));
    } else if (e.key === 'Enter' && activeIdx >= 0) {
      e.preventDefault();
      selectItem(items[activeIdx]?.dataset.tid);
    } else if (e.key === 'Escape') { hide(); }
  });

  // Reposition on scroll/resize; hide on blur
  inputEl.addEventListener('blur', () => setTimeout(hide, 150));
  window.addEventListener('scroll', reposition, true);
  window.addEventListener('resize', reposition);

  // Cleanup when input is removed from DOM (panel close/re-render)
  new MutationObserver(() => {
    if (!document.contains(inputEl)) { hide(); if (dropdown) dropdown.remove(); dropdown = null; }
  }).observe(document.body, { childList: true, subtree: true });
}

// ── KOSTENSTELLEN AUTOCOMPLETE ────────────────────────────────────────────────
// Wraps an <input> with a dropdown showing matching Kostenstellen entries.
// On select: fills input with "NUMMER – Bezeichnung"
function initKostenstAuto(inputEl) {
  if (!inputEl) return;
  let dropdown = null;
  let activeIdx = -1;

  function getOrCreate() {
    if (!dropdown) {
      dropdown = document.createElement('div');
      dropdown.className = 'tid-ac-dropdown';
      dropdown.style.cssText = 'display:none;position:fixed;z-index:9999';
      document.body.appendChild(dropdown);
    }
    return dropdown;
  }
  function reposition() {
    if (!dropdown || dropdown.style.display === 'none') return;
    const r = inputEl.getBoundingClientRect();
    dropdown.style.left  = r.left + 'px';
    dropdown.style.top   = (r.bottom + 2) + 'px';
    dropdown.style.width = r.width + 'px';
  }
  function hide() { if (dropdown) dropdown.style.display = 'none'; activeIdx = -1; }
  function show(items) {
    if (!items.length) { hide(); return; }
    activeIdx = -1;
    const dd = getOrCreate();
    dd.innerHTML = items.slice(0, 12).map((e, i) =>
      `<div class="tid-ac-item" data-nr="${esc(e.nr)}" data-idx="${i}">
        <span class="tid-ac-nr">${esc(e.nr)}</span>
        <span class="tid-ac-name">${esc(e.label)}</span>
      </div>`
    ).join('');
    dd.querySelectorAll('.tid-ac-item').forEach(el => {
      el.addEventListener('mousedown', ev => { ev.preventDefault(); selectItem(el.dataset.nr); });
    });
    dd.style.display = 'block';
    reposition();
  }
  function selectItem(nr) {
    const entry = KOSTENST_DATA.find(e => e.nr === nr);
    if (!entry) return;
    inputEl.value = entry.nr + ' – ' + entry.label;
    hide();
  }
  inputEl.addEventListener('input', () => {
    const q = inputEl.value.trim().toLowerCase();
    if (q.length < 1) { hide(); return; }
    const matches = KOSTENST_DATA.filter(e =>
      e.nr.includes(q) || e.label.toLowerCase().includes(q)
    );
    show(matches);
  });
  inputEl.addEventListener('keydown', e => {
    if (!dropdown || dropdown.style.display === 'none') return;
    const items = dropdown.querySelectorAll('.tid-ac-item');
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      activeIdx = Math.min(activeIdx + 1, items.length - 1);
      items.forEach((el, i) => el.classList.toggle('tid-ac-active', i === activeIdx));
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      activeIdx = Math.max(activeIdx - 1, 0);
      items.forEach((el, i) => el.classList.toggle('tid-ac-active', i === activeIdx));
    } else if (e.key === 'Enter' && activeIdx >= 0) {
      e.preventDefault();
      selectItem(items[activeIdx]?.dataset.nr);
    } else if (e.key === 'Escape') { hide(); }
  });
  inputEl.addEventListener('blur', () => setTimeout(hide, 150));
  window.addEventListener('scroll', reposition, true);
  window.addEventListener('resize', reposition);
  new MutationObserver(() => {
    if (!document.contains(inputEl)) { hide(); if (dropdown) dropdown.remove(); dropdown = null; }
  }).observe(document.body, { childList: true, subtree: true });
}

// ── WIZARD ────────────────────────────────────────────────────────────────────
function lookupTID(val) {
  const tid = (val || '').trim().toUpperCase();
  const matchEl = $id('tid-match');
  // Normalize: user might type "4001-00010" or "400100010" — try exact then padded
  const hit = TID_MAP_ACTIVE[tid] || TID_MAP_ACTIVE[tid.replace(/^(\d{4})(\d{5})$/, '$1-$2')] || null;
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
  // Artikelnummer + Kostenstelle autocomplete in wizard
  const wArtNr = $id('f-Artikelnummer');
  if (wArtNr) initTidAutocomplete(wArtNr, key => $id('f-' + key));
  initKostenstAuto($id('f-Kostenstelle'));
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
  ['Title','Beschreibung','Warengruppe','Mengeneinheit',
   'Mindestlagermenge','Termin','Artikelnummer','ExterneArtikelnummer',
   'Lieferant','Lieferant2','Lieferant3','Lieferant4',
   'GeschaetzterPreis','Kostenstelle']
    .forEach(k => { const el = $id('f-'+k); if(el) el.value = ''; });
  const prioEl = $id('f-Prioritaet');
  if (prioEl) {
    // SP populate-code may have replaced HTML options with SP choice values (e.g. "HIGH").
    // Find a "Normal"-like option case-insensitively; if not found, pick first non-empty option.
    const prioOpts = [...prioEl.options];
    const normalOpt = prioOpts.find(o => /^(normal|standard)$/i.test(o.value.trim()));
    if (normalOpt) prioEl.value = normalOpt.value;
    else {
      const first = prioOpts.find(o => o.value !== '');
      prioEl.value = first ? first.value : '';
    }
  }
  const mengeEl = $id('f-Menge'); if (mengeEl) mengeEl.value = '1'; // Standardwert
  // Reset submit button in case previous submission left it in a loading state
  const submitBtn = $id('btn-submit');
  if (submitBtn) { submitBtn.disabled = false; submitBtn.textContent = '✓ Anfrage einreichen'; }
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
      Title:                title,
      Beschreibung:         $id('f-Beschreibung').value.trim(),
      Warengruppe:          wg,
      Prioritaet:           $id('f-Prioritaet').value,
      Artikelnummer:        $id('f-Artikelnummer').value.trim(),
      ExterneArtikelnummer: $id('f-ExterneArtikelnummer')?.value.trim() || '',
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
    const extraSelected = document.querySelector('#beschaffungslogik-extra-cards .check-card.selected');
    if (!extraSelected) { toast('Bitte unter „Zusätzlich kombinierbar" eine Option auswählen.', 'error'); return; }
    const preisVal = $id('f-GeschaetzterPreis').value;
    if (!preisVal || parseFloat(preisVal) <= 0) { toast('Bitte Bestellvolumen in € angeben.', 'error'); return; }
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
        ['Bestellvolumen in €', d.GeschaetzterPreis ? fmtEuro(d.GeschaetzterPreis) : null],
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
          // Read the file buffer ONCE before entering the retry loop so a
          // second attempt doesn't try to re-read a potentially consumed stream.
          const buf = await file.arrayBuffer();
          // SP REST may not yet have propagated the new item (same lag as PATCH 404).
          // Wrap in retryOn404 so we wait up to ~4.8s before giving up.
          // Filename: single-quote chars must be escaped as %27 inside the OData string literal.
          const fname = file.name.replace(/'/g, "''");
          const fnameUrl = encodeURIComponent(fname);
          const uploadUrl = `${SP_BASE}/_api/web/lists/getByTitle('${SP_LIST}')/items(${itemId})/AttachmentFiles/add(FileName='${fnameUrl}')`;
          await retryOn404(async () => {
            const r = await fetch(uploadUrl, {
              method: 'POST',
              headers: {
                Authorization: 'Bearer ' + tok,
                Accept: 'application/json;odata=nometadata',
                'Content-Type': 'application/octet-stream',
              },
              body: buf,
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

// ── TABELLEN-IMPORTER ────────────────────────────────────────────────────────

function initImporter() {
  renderImporter();
}

function downloadCSV(csvText, filename) {
  const bom  = '﻿'; // UTF-8 BOM for Excel
  const blob = new Blob([bom + csvText], { type: 'text/csv;charset=utf-8;' });
  const url  = URL.createObjectURL(blob);
  const a    = Object.assign(document.createElement('a'), { href: url, download: filename });
  document.body.appendChild(a); a.click();
  document.body.removeChild(a); URL.revokeObjectURL(url);
}

function exportTidCSV() {
  const rows = ['Artikelnummer;Bezeichnung;Warengruppe'];
  for (const [nr, h] of Object.entries(TID_MAP_ACTIVE)) {
    rows.push(`${nr};${(h.b||'').replace(/;/g,'')};${(h.w||'').replace(/;/g,'')}`);
  }
  downloadCSV(rows.join('\r\n'), 'Artikelnummern_Export.csv');
}

function exportKstCSV() {
  const rows = ['Kostenstelle;Bezeichnung'];
  for (const e of KOSTENST_DATA) {
    rows.push(`${e.nr};${(e.label||'').replace(/;/g,'')}`);
  }
  downloadCSV(rows.join('\r\n'), 'Kostenstellen_Export.csv');
}

function downloadTidTemplate() {
  downloadCSV('Artikelnummer;Bezeichnung;Warengruppe\r\n1234-00001;Beispielartikel;Ersatzteile', 'Artikelnummern_Vorlage.csv');
}

function downloadKstTemplate() {
  downloadCSV('Kostenstelle;Bezeichnung\r\n10011;Fertigungsleitung', 'Kostenstellen_Vorlage.csv');
}

function renderImporter() {
  const tidCount  = Object.keys(TID_MAP_ACTIVE).length;
  const kstCount  = KOSTENST_DATA.length;
  const tidIsCustom = !!localStorage.getItem(LS_TID_KEY);
  const kstIsCustom = !!localStorage.getItem(LS_KOSTENST_KEY);

  $id('view-importer').innerHTML = `
    <div style="max-width:700px;margin:0 auto">
      <p style="color:#6b7280;margin-bottom:20px;font-size:.9rem">
        Laden Sie CSV- oder XLSX-Dateien hoch, um die Artikelnummer- und Kostenstellentabellen zu aktualisieren.
        Die Daten werden lokal im Browser gespeichert.
      </p>

      <!-- Artikelnummern -->
      <div class="imp-card">
        <div class="imp-card-header">
          <div>
            <div class="imp-card-title">📦 Artikelnummern (TID-Tabelle)</div>
            <div class="imp-card-sub">${tidCount} Einträge geladen${tidIsCustom ? ' <span class="imp-badge-custom">Angepasst</span>' : ' <span class="imp-badge-default">Standard</span>'}</div>
          </div>
          <div style="display:flex;gap:6px;flex-wrap:wrap">
            <button class="btn btn-sm btn-outline" onclick="exportTidCSV()" title="Aktuelle Tabelle als CSV herunterladen">⬇ Export</button>
            <button class="btn btn-sm btn-ghost" onclick="downloadTidTemplate()" title="Leere Vorlage herunterladen">📄 Vorlage</button>
            ${tidIsCustom ? `<button class="btn btn-sm btn-ghost" onclick="resetTable('tid')">↺ Zurücksetzen</button>` : ''}
          </div>
        </div>
        <div class="imp-format-hint">
          Format: CSV mit Semikolon <code>Artikelnummer;Bezeichnung;Warengruppe</code><br>
          oder XLSX mit denselben Spalten in Zeile 1. Bestehende Daten werden durch Import <strong>ergänzt</strong> (gleiche Nr. wird überschrieben).
        </div>
        <div class="imp-drop-zone" id="imp-drop-tid"
          ondragover="event.preventDefault();this.classList.add('drag-over')"
          ondragleave="this.classList.remove('drag-over')"
          ondrop="handleImpDrop(event,'tid')">
          <span>📂 CSV oder XLSX hier ablegen oder</span>
          <label class="btn btn-sm btn-outline" style="cursor:pointer;margin-left:6px">
            Datei auswählen
            <input type="file" accept=".csv,.xlsx" style="display:none" onchange="handleImpFile(this,'tid')">
          </label>
        </div>
        <div id="imp-result-tid" class="imp-result"></div>
      </div>

      <!-- Kostenstellen -->
      <div class="imp-card" style="margin-top:16px">
        <div class="imp-card-header">
          <div>
            <div class="imp-card-title">🏢 Kostenstellen</div>
            <div class="imp-card-sub">${kstCount} Einträge geladen${kstIsCustom ? ' <span class="imp-badge-custom">Angepasst</span>' : ' <span class="imp-badge-default">Standard</span>'}</div>
          </div>
          <div style="display:flex;gap:6px;flex-wrap:wrap">
            <button class="btn btn-sm btn-outline" onclick="exportKstCSV()" title="Aktuelle Tabelle als CSV herunterladen">⬇ Export</button>
            <button class="btn btn-sm btn-ghost" onclick="downloadKstTemplate()" title="Leere Vorlage herunterladen">📄 Vorlage</button>
            ${kstIsCustom ? `<button class="btn btn-sm btn-ghost" onclick="resetTable('kst')">↺ Zurücksetzen</button>` : ''}
          </div>
        </div>
        <div class="imp-format-hint">
          Format: CSV mit Semikolon <code>Kostenstelle;Bezeichnung</code> oder eine Spalte <code>10011 Fertigungsleitung</code><br>
          oder XLSX mit denselben Spalten. Hochladen <strong>ersetzt</strong> die gesamte Tabelle.
        </div>
        <div class="imp-drop-zone" id="imp-drop-kst"
          ondragover="event.preventDefault();this.classList.add('drag-over')"
          ondragleave="this.classList.remove('drag-over')"
          ondrop="handleImpDrop(event,'kst')">
          <span>📂 CSV oder XLSX hier ablegen oder</span>
          <label class="btn btn-sm btn-outline" style="cursor:pointer;margin-left:6px">
            Datei auswählen
            <input type="file" accept=".csv,.xlsx" style="display:none" onchange="handleImpFile(this,'kst')">
          </label>
        </div>
        <div id="imp-result-kst" class="imp-result"></div>
      </div>
    </div>`;
}

function handleImpDrop(e, type) {
  e.preventDefault();
  e.currentTarget.classList.remove('drag-over');
  const file = e.dataTransfer?.files?.[0];
  if (file) parseImportFile(file, type);
}

function handleImpFile(input, type) {
  const file = input.files?.[0];
  if (file) parseImportFile(file, type);
  input.value = '';
}

async function parseImportFile(file, type) {
  const resultEl = $id('imp-result-' + type);
  resultEl.textContent = '⏳ Wird verarbeitet…';
  try {
    const ext = file.name.split('.').pop().toLowerCase();
    let rows = []; // [{col0, col1, col2, ...}]

    if (ext === 'csv') {
      const text = await file.text();
      const lines = text.split(/\r?\n/).filter(l => l.trim());
      rows = lines.map(l => l.split(/;|\t/).map(c => c.trim().replace(/^"|"$/g, '')));
    } else if (ext === 'xlsx') {
      rows = await parseXlsxFile(file);
    } else {
      resultEl.innerHTML = '<span style="color:#ef4444">❌ Nur CSV oder XLSX-Dateien unterstützt.</span>';
      return;
    }

    if (type === 'tid') applyTidImport(rows, resultEl);
    else                applyKstImport(rows, resultEl);

  } catch(e) {
    resultEl.innerHTML = `<span style="color:#ef4444">❌ Fehler: ${esc(e.message)}</span>`;
  }
}

// Parse XLSX in-browser using built-in zip (no external library needed for simple cases)
async function parseXlsxFile(file) {
  // Use SheetJS if available, else fallback to manual parse
  if (typeof XLSX !== 'undefined') {
    const buf = await file.arrayBuffer();
    const wb  = XLSX.read(buf, { type: 'array' });
    const ws  = wb.Sheets[wb.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    return data.map(r => r.map(c => String(c ?? '').trim()));
  }
  // Fallback: manual zip parse (shared strings + sheet1)
  const buf  = await file.arrayBuffer();
  const u8   = new Uint8Array(buf);
  // Minimal zip extractor
  function readZipEntry(u8, name) {
    const enc = new TextEncoder();
    const nameBytes = enc.encode(name);
    for (let i = 0; i < u8.length - 30; i++) {
      if (u8[i] !== 0x50 || u8[i+1] !== 0x4B || u8[i+2] !== 0x03 || u8[i+3] !== 0x04) continue;
      const fnLen = u8[i+26] | (u8[i+27] << 8);
      const exLen = u8[i+28] | (u8[i+29] << 8);
      const fn = new TextDecoder().decode(u8.slice(i+30, i+30+fnLen));
      const dataStart = i + 30 + fnLen + exLen;
      const compSize  = u8[i+18] | (u8[i+19]<<8) | (u8[i+20]<<16) | (u8[i+21]<<24);
      const method    = u8[i+8] | (u8[i+9]<<8);
      if (fn === name) {
        if (method === 0) return new TextDecoder('utf-8').decode(u8.slice(dataStart, dataStart + compSize));
        // deflate — skip; user should use SheetJS for compressed entries
        return null;
      }
    }
    return null;
  }
  const ssXml  = readZipEntry(u8, 'xl/sharedStrings.xml');
  const shXml  = readZipEntry(u8, 'xl/worksheets/sheet1.xml');
  if (!ssXml || !shXml) throw new Error('XLSX-Datei konnte nicht gelesen werden. Bitte als CSV exportieren.');
  const parser = new DOMParser();
  const ssDoc  = parser.parseFromString(ssXml, 'text/xml');
  const shDoc  = parser.parseFromString(shXml, 'text/xml');
  const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const strings = [...ssDoc.getElementsByTagNameNS(ns, 'si')].map(si =>
    [...si.getElementsByTagNameNS(ns, 't')].map(t => t.textContent || '').join('')
  );
  const result = [];
  for (const row of shDoc.getElementsByTagNameNS(ns, 'row')) {
    const cells = [];
    for (const c of row.getElementsByTagNameNS(ns, 'c')) {
      const t  = c.getAttribute('t');
      const v  = c.getElementsByTagNameNS(ns, 'v')[0]?.textContent || '';
      cells.push(t === 's' ? (strings[parseInt(v)] || '') : v);
    }
    result.push(cells);
  }
  return result;
}

function applyTidImport(rows, resultEl) {
  // Detect header row: if first row looks like a header (non-numeric first cell)
  let start = 0;
  if (rows.length && !/^\d/.test(rows[0][0] || '')) start = 1;
  const imported = {};
  let count = 0;
  for (const row of rows.slice(start)) {
    const nr = String(row[0] || '').trim().toUpperCase();
    const b  = String(row[1] || '').trim();
    const w  = String(row[2] || '').trim();
    if (!nr || !b) continue;
    imported[nr] = { b, w: w || 'Sonstiges', g: w || 'Sonstiges' };
    count++;
  }
  if (!count) { resultEl.innerHTML = '<span style="color:#ef4444">❌ Keine gültigen Zeilen gefunden. Prüfen Sie das Format.</span>'; return; }
  const merged = Object.assign({}, TID_MAP, imported);
  localStorage.setItem(LS_TID_KEY, JSON.stringify(imported)); // store only the delta
  reloadLookupData();
  resultEl.innerHTML = `<span style="color:#16a34a">✓ ${count} Artikelnummern importiert. Gesamt: ${Object.keys(TID_MAP_ACTIVE).length} Einträge.</span>`;
  setTimeout(renderImporter, 800);
}

function applyKstImport(rows, resultEl) {
  let start = 0;
  const firstCell = String(rows[0]?.[0] || '').trim().toLowerCase();
  if (!firstCell || /kostenstelle|nr|nummer/i.test(firstCell)) start = 1;
  const result = [];
  for (const row of rows.slice(start)) {
    const col0 = String(row[0] || '').trim();
    const col1 = String(row[1] || '').trim();
    if (!col0) continue;
    if (col1) {
      // Two-column format: col0=nr, col1=label
      result.push({ nr: col0, label: col1 });
    } else {
      // Single-column: "10011 Fertigungsleitung"
      const m = col0.match(/^(\d+)\s+(.+)$/);
      if (m) result.push({ nr: m[1], label: m[2].trim() });
    }
  }
  if (!result.length) { resultEl.innerHTML = '<span style="color:#ef4444">❌ Keine gültigen Zeilen gefunden. Prüfen Sie das Format.</span>'; return; }
  localStorage.setItem(LS_KOSTENST_KEY, JSON.stringify(result));
  reloadLookupData();
  resultEl.innerHTML = `<span style="color:#16a34a">✓ ${result.length} Kostenstellen importiert.</span>`;
  setTimeout(renderImporter, 800);
}

function resetTable(type) {
  if (type === 'tid')  localStorage.removeItem(LS_TID_KEY);
  if (type === 'kst')  localStorage.removeItem(LS_KOSTENST_KEY);
  reloadLookupData();
  renderImporter();
}

// ── REPORTS & AUSWERTUNGEN ───────────────────────────────────────────────────
let reportFilters = {};
let reportSort    = { col: 'id', dir: 'desc' };

// rf = resolved field reader
function _rf(i, key) { return getField(i, resolvedFields[key] || key) || ''; }

const REPORT_COLS = [
  { key:'id',               label:'ID',             type:'num',  get:i => parseInt(i.id, 10) },
  { key:'Title',            label:'Bezeichnung',    type:'text', get:i => cleanTitle(getField(i,'Title') || '') },
  { key:'Status',           label:'Status',         type:'text', get:i => getStatusVal(i) || 'Eingereicht' },
  { key:'Warengruppe',      label:'Warengruppe',    type:'text', get:i => _rf(i,'Warengruppe') },
  { key:'Prioritaet',       label:'Priorität',      type:'text', get:i => _rf(i,'Prioritaet') },
  { key:'Menge',            label:'Menge',          type:'num',  get:i => { const n = parseFloat(_rf(i,'Menge')); return isNaN(n) ? null : n; } },
  { key:'Mengeneinheit',    label:'ME',             type:'text', get:i => _rf(i,'Mengeneinheit') },
  { key:'GeschaetzterPreis',label:'Volumen (€)',    type:'num',  get:i => { const n = parseFloat(_rf(i,'GeschaetzterPreis')); return isNaN(n) ? null : n; } },
  { key:'Beschaffungslogik',label:'Beschaffungsart',type:'text', get:i => _rf(i,'Beschaffungslogik') },
  { key:'Lieferant',        label:'Lieferant',      type:'text', get:i => _rf(i,'Lieferant') },
  { key:'Kostenstelle',     label:'Kostenstelle',   type:'text', get:i => _rf(i,'Kostenstelle') },
  { key:'Bestellnummer',    label:'Bestell-Nr.',    type:'text', get:i => _rf(i,'Bestellnummer') },
  { key:'Artikelnummer',    label:'Artikel-Nr.',    type:'text', get:i => _rf(i,'Artikelnummer') },
  { key:'Termin',           label:'Benötigt bis',   type:'date', get:i => _rf(i,'Termin') },
  { key:'creator',          label:'Ersteller',      type:'text', get:i => i.createdBy?.user?.displayName || i.createdBy?.user?.email || '' },
  { key:'created',          label:'Erstellt am',    type:'date', get:i => i.createdDateTime || '' },
];

function _uniqueReportVals(colKey) {
  const col = REPORT_COLS.find(c => c.key === colKey);
  if (!col) return [];
  const set = new Set();
  for (const i of allItems) { const v = col.get(i); if (v != null && String(v).trim() !== '') set.add(String(v)); }
  return [...set].sort((a, b) => a.localeCompare(b, 'de'));
}

function initReports() {
  reportFilters = {};
  renderReports();
}

function renderReports() {
  const statusOpts = _uniqueReportVals('Status');
  const wgOpts     = _uniqueReportVals('Warengruppe');
  const maOpts     = _uniqueReportVals('creator');
  const prioOpts   = _uniqueReportVals('Prioritaet');
  const baOpts     = _uniqueReportVals('Beschaffungslogik');

  const sel = (id, label, opts) => `
    <div class="rep-filter">
      <label>${label}</label>
      <select class="rep-input" id="${id}" onchange="applyReportFilters()">
        <option value="">Alle</option>
        ${opts.map(o => `<option value="${esc(o)}">${esc(o)}</option>`).join('')}
      </select>
    </div>`;
  const txt = (id, label, ph) => `
    <div class="rep-filter">
      <label>${label}</label>
      <input type="text" class="rep-input" id="${id}" placeholder="${ph||''}" oninput="applyReportFilters()"/>
    </div>`;
  const num = (id, label, ph) => `
    <div class="rep-filter">
      <label>${label}</label>
      <input type="number" class="rep-input" id="${id}" placeholder="${ph||''}" min="0" step="0.01" oninput="applyReportFilters()"/>
    </div>`;
  const dat = (id, label) => `
    <div class="rep-filter">
      <label>${label}</label>
      <input type="date" class="rep-input" id="${id}" onchange="applyReportFilters()"/>
    </div>`;

  $id('view-reports').innerHTML = `
    <div class="rep-wrap">
      <div class="rep-toolbar">
        <div class="rep-toolbar-left">
          <span class="rep-count" id="rep-count"></span>
          <span class="rep-sum" id="rep-sum"></span>
        </div>
        <div class="rep-toolbar-right">
          <button class="btn btn-sm btn-ghost" onclick="resetReportFilters()">↺ Filter zurücksetzen</button>
          <button class="btn btn-sm btn-primary" onclick="exportReportsExcel()">⬇ Als Excel exportieren</button>
        </div>
      </div>
      <div class="rep-filters">
        ${txt('rep-search','Suche','Titel, ID, Artikel…')}
        ${sel('rep-status','Status', statusOpts)}
        ${sel('rep-wg','Warengruppe', wgOpts)}
        ${sel('rep-ma','Mitarbeiter', maOpts)}
        ${sel('rep-prio','Priorität', prioOpts)}
        ${sel('rep-ba','Beschaffungsart', baOpts)}
        ${txt('rep-lieferant','Lieferant','enthält…')}
        ${txt('rep-kst','Kostenstelle','enthält…')}
        ${txt('rep-be','Bestell-Nr.','enthält…')}
        ${num('rep-preis-min','Volumen ab €','min')}
        ${num('rep-preis-max','Volumen bis €','max')}
        ${dat('rep-created-from','Erstellt von')}
        ${dat('rep-created-to','Erstellt bis')}
        ${dat('rep-termin-from','Benötigt von')}
        ${dat('rep-termin-to','Benötigt bis')}
      </div>
      <div class="rep-table-wrap">
        <table class="rep-table" id="rep-table"></table>
      </div>
    </div>`;
  renderReportsTable();
}

function applyReportFilters() {
  reportFilters = {
    search:      ($id('rep-search')?.value || '').toLowerCase().trim(),
    status:      $id('rep-status')?.value || '',
    wg:          $id('rep-wg')?.value || '',
    ma:          $id('rep-ma')?.value || '',
    prio:        $id('rep-prio')?.value || '',
    ba:          $id('rep-ba')?.value || '',
    lieferant:   ($id('rep-lieferant')?.value || '').toLowerCase().trim(),
    kst:         ($id('rep-kst')?.value || '').toLowerCase().trim(),
    be:          ($id('rep-be')?.value || '').toLowerCase().trim(),
    preisMin:    parseFloat($id('rep-preis-min')?.value) || null,
    preisMax:    parseFloat($id('rep-preis-max')?.value) || null,
    createdFrom: $id('rep-created-from')?.value || '',
    createdTo:   $id('rep-created-to')?.value || '',
    terminFrom:  $id('rep-termin-from')?.value || '',
    terminTo:    $id('rep-termin-to')?.value || '',
  };
  renderReportsTable();
}

function resetReportFilters() {
  reportFilters = {};
  ['rep-search','rep-status','rep-wg','rep-ma','rep-prio','rep-ba','rep-lieferant',
   'rep-kst','rep-be','rep-preis-min','rep-preis-max','rep-created-from','rep-created-to',
   'rep-termin-from','rep-termin-to'].forEach(id => { const el = $id(id); if (el) el.value = ''; });
  renderReportsTable();
}

function getReportRows() {
  const f = reportFilters;
  const get = (i, key) => { const c = REPORT_COLS.find(x => x.key === key); return c ? c.get(i) : ''; };
  let rows = allItems.filter(i => {
    if (f.status    && (get(i,'Status') || '') !== f.status) return false;
    if (f.wg        && (get(i,'Warengruppe') || '') !== f.wg) return false;
    if (f.ma        && (get(i,'creator') || '') !== f.ma) return false;
    if (f.prio      && (get(i,'Prioritaet') || '') !== f.prio) return false;
    if (f.ba        && (get(i,'Beschaffungslogik') || '') !== f.ba) return false;
    if (f.lieferant && !String(get(i,'Lieferant') || '').toLowerCase().includes(f.lieferant)) return false;
    if (f.kst       && !String(get(i,'Kostenstelle') || '').toLowerCase().includes(f.kst)) return false;
    if (f.be        && !String(get(i,'Bestellnummer') || '').toLowerCase().includes(f.be)) return false;
    const preis = get(i,'GeschaetzterPreis');
    if (f.preisMin != null && (preis == null || preis < f.preisMin)) return false;
    if (f.preisMax != null && (preis == null || preis > f.preisMax)) return false;
    const created = (get(i,'created') || '').slice(0,10);
    if (f.createdFrom && (!created || created < f.createdFrom)) return false;
    if (f.createdTo   && (!created || created > f.createdTo)) return false;
    const termin = (get(i,'Termin') || '').slice(0,10);
    if (f.terminFrom && (!termin || termin < f.terminFrom)) return false;
    if (f.terminTo   && (!termin || termin > f.terminTo)) return false;
    if (f.search) {
      const hay = [get(i,'id'), get(i,'Title'), get(i,'Artikelnummer'), get(i,'Lieferant'),
                   get(i,'Bestellnummer'), get(i,'Kostenstelle'), get(i,'creator')]
                  .map(v => String(v ?? '').toLowerCase()).join(' ');
      if (!hay.includes(f.search)) return false;
    }
    return true;
  });
  // Sort
  const col = REPORT_COLS.find(c => c.key === reportSort.col) || REPORT_COLS[0];
  const dir = reportSort.dir === 'asc' ? 1 : -1;
  rows.sort((a, b) => {
    let va = col.get(a), vb = col.get(b);
    if (col.type === 'num')  { va = va == null ? -Infinity : va; vb = vb == null ? -Infinity : vb; return (va - vb) * dir; }
    if (col.type === 'date') { return (String(va).localeCompare(String(vb))) * dir; }
    return String(va).localeCompare(String(vb), 'de') * dir;
  });
  return rows;
}

function reportSortBy(colKey) {
  if (reportSort.col === colKey) reportSort.dir = reportSort.dir === 'asc' ? 'desc' : 'asc';
  else { reportSort.col = colKey; reportSort.dir = 'asc'; }
  renderReportsTable();
}

function renderReportsTable() {
  const table = $id('rep-table');
  if (!table) return;
  const rows = getReportRows();

  // Count + sum
  const cnt = $id('rep-count');
  if (cnt) cnt.textContent = `${rows.length} ${rows.length === 1 ? 'Eintrag' : 'Einträge'}`;
  const sumVal = rows.reduce((s, i) => {
    const c = REPORT_COLS.find(x => x.key === 'GeschaetzterPreis');
    return s + (c.get(i) || 0);
  }, 0);
  const sumEl = $id('rep-sum');
  if (sumEl) sumEl.textContent = sumVal > 0 ? `· Gesamtvolumen ${fmtEuro(sumVal)}` : '';

  const arrow = key => reportSort.col === key ? (reportSort.dir === 'asc' ? ' ▲' : ' ▼') : '';
  const head = `<thead><tr>${REPORT_COLS.map(c =>
    `<th class="rep-th${c.type==='num'?' rep-right':''}" onclick="reportSortBy('${c.key}')">${esc(c.label)}${arrow(c.key)}</th>`
  ).join('')}</tr></thead>`;

  const body = rows.length === 0
    ? `<tbody><tr><td class="rep-empty" colspan="${REPORT_COLS.length}">Keine Einträge für die gewählten Filter.</td></tr></tbody>`
    : `<tbody>${rows.map(i => {
        const id = i.id;
        return `<tr class="rep-row" onclick="navigate('detail','${id}')">${REPORT_COLS.map(c => {
          let v = c.get(i);
          if (c.key === 'Status') return `<td class="rep-td">${statusBadge(v)}</td>`;
          if (c.type === 'date')  v = v ? fmtDate(v) : '–';
          else if (c.key === 'GeschaetzterPreis') v = v != null ? fmtEuro(v) : '–';
          else if (c.type === 'num') v = v != null ? v.toLocaleString('de-DE') : '–';
          else v = (v == null || v === '') ? '–' : v;
          return `<td class="rep-td${c.type==='num'?' rep-right':''}">${esc(String(v))}</td>`;
        }).join('')}</tr>`;
      }).join('')}</tbody>`;

  table.innerHTML = head + body;
}

function exportReportsExcel() {
  const rows = getReportRows();
  if (!rows.length) { toast('Keine Daten zum Exportieren.', 'error'); return; }
  const header = REPORT_COLS.map(c => c.label);
  const aoa = [header];
  for (const it of rows) {
    aoa.push(REPORT_COLS.map(c => {
      let v = c.get(it);
      if (c.type === 'date' && v) v = fmtDate(v);
      return v == null ? '' : v;
    }));
  }
  const stamp = new Date().toISOString().slice(0,10);
  if (typeof XLSX !== 'undefined') {
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = REPORT_COLS.map(c => ({ wch: Math.max(10, c.label.length + 2) }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Bedarfsanfragen');
    XLSX.writeFile(wb, `Bedarfsanfragen_Report_${stamp}.xlsx`);
    toast(`${rows.length} Einträge als Excel exportiert.`, 'success');
  } else {
    // Fallback: CSV
    const csv = aoa.map(r => r.map(c => `"${String(c).replace(/"/g,'""')}"`).join(';')).join('\r\n');
    downloadCSV(csv, `Bedarfsanfragen_Report_${stamp}.csv`);
    toast(`${rows.length} Einträge als CSV exportiert (Excel-Bibliothek nicht geladen).`, 'success');
  }
}

// ── SCHULUNG (BETA) ──────────────────────────────────────────────────────────
function initSchulung() {
  $id('view-schulung').innerHTML = `
    <div class="schulung-wrap">
      <div class="schulung-hero">
        <div class="schulung-badge">Beta</div>
        <h1>Willkommen im DIHAG Bedarfsanfragen-Tool</h1>
        <p>Diese Schulung erklärt Schritt für Schritt, wie Sie Bedarfsanfragen erstellen,
           verfolgen und genehmigen. Die Anwendung ist mit SharePoint verbunden – alle
           Eingaben werden zentral gespeichert und über Power Automate weitergeleitet.</p>
      </div>

      <div class="schulung-toc">
        <a href="#sch-1">1 · Überblick</a>
        <a href="#sch-2">2 · Neue Anfrage</a>
        <a href="#sch-3">3 · Sammelanfrage</a>
        <a href="#sch-4">4 · Status & Genehmigung</a>
        <a href="#sch-5">5 · Einkauf-Daten</a>
        <a href="#sch-6">6 · Reports</a>
        <a href="#sch-7">7 · Wiedervorlage/Favoriten</a>
        <a href="#sch-8">8 · Tabellen-Importer</a>
        <a href="#sch-9">9 · FAQ</a>
      </div>

      <section class="schulung-sec" id="sch-1">
        <h2>1 · Überblick</h2>
        <p>Die Navigation links führt zu den Hauptbereichen:</p>
        <ul>
          <li><b>Dashboard</b> – Liste aller Anfragen mit Status-Chips, Warengruppen- und Mitarbeiterfilter.</li>
          <li><b>Neue Anfrage</b> – Assistent für eine einzelne Position.</li>
          <li><b>Sammelanfrage</b> – Assistent für mehrere Positionen in einer Anfrage.</li>
          <li><b>Meine Anfragen</b> – nur Ihre eigenen Anfragen.</li>
          <li><b>Reports</b> – Auswertungen mit vielen Filtern und Excel-Export.</li>
          <li><b>Tabellen-Importer</b> – Artikelnummern- und Kostenstellentabellen aktualisieren.</li>
        </ul>
        <div class="schulung-tip">💡 Die Liste aktualisiert sich automatisch alle 20 Sekunden.
          Während Sie eine Anfrage ausfüllen, pausiert die Aktualisierung, damit nichts verloren geht.</div>
      </section>

      <section class="schulung-sec" id="sch-2">
        <h2>2 · Neue Anfrage erstellen</h2>
        <ol>
          <li><b>Bedarf:</b> Bezeichnung und Warengruppe (Pflicht). Geben Sie eine Artikelnummer
              ein – Bezeichnung und Warengruppe werden automatisch vorgeschlagen.</li>
          <li><b>Menge:</b> Menge, Mengeneinheit und „Benötigt bis"-Termin.</li>
          <li><b>Beschaffung:</b> Beschaffungsart, Lieferant(en), Kostenstelle und
              <b>Bestellvolumen in € (Pflicht)</b>. Anhand des Volumens zeigt das Tool den
              erforderlichen Genehmigungsweg an (z. B. Anzahl Angebote).</li>
          <li><b>Prüfen & Einreichen:</b> Zusammenfassung kontrollieren und absenden.</li>
        </ol>
        <div class="schulung-tip">💡 Über das Feld „Externe Artikelnummer" erfassen Sie bei
          Katalogware die Lieferanten-Nummer.</div>
      </section>

      <section class="schulung-sec" id="sch-3">
        <h2>3 · Sammelanfrage (mehrere Positionen)</h2>
        <p>Nutzen Sie die Sammelanfrage, wenn mehrere unterschiedliche Artikel zusammen
           angefragt werden.</p>
        <ol>
          <li><b>Positionen erfassen:</b> Pro Zeile Artikelnummer, Bezeichnung, Menge, ME,
              <b>Preis</b>, <b>Benötigt bis</b> und <b>Kostenstelle</b>. Mit „+ Position
              hinzufügen" beliebig erweitern.</li>
          <li>Das <b>Bestellvolumen</b> wird automatisch aus den Einzelpreisen summiert.</li>
          <li>Warengruppe, Priorität und Beschaffungsart gelten für die gesamte Anfrage.</li>
        </ol>
      </section>

      <section class="schulung-sec" id="sch-4">
        <h2>4 · Status & Genehmigung</h2>
        <p>Jede Anfrage durchläuft mehrere Stufen, die im „Verlauf" sichtbar sind:</p>
        <ul>
          <li><b>Eingereicht</b> → beim Öffnen im Dashboard automatisch <b>In Prüfung (Einkauf)</b>.</li>
          <li>Der Einkauf entscheidet direkt im Detailbereich über <b>Genehmigen</b> oder
              <b>Ablehnen</b>.</li>
          <li>Nach <b>Freigegeben</b> können die Einkauf-Daten (Bestellnummer etc.) erfasst werden.</li>
        </ul>
        <div class="schulung-tip">💡 In „Meine Anfragen" sehen Sie nur den Fortschritt – die
          Genehmigungsschaltflächen erscheinen dort nicht.</div>
      </section>

      <section class="schulung-sec" id="sch-5">
        <h2>5 · Einkauf-Daten eintragen</h2>
        <p>Bei Status <b>Freigegeben</b> öffnen Sie über „📦 Einkauf-Daten" das Formular für
           Bestellnummer, Lieferdatum, tatsächlichen Preis und Angebots-PDFs.</p>
        <div class="schulung-tip">⚠️ Ist eine <b>Bestellnummer (BE)</b> bereits bei einer anderen
          Anfrage hinterlegt, warnt das Tool sofort – so vermeiden Sie Doppelbestellungen.</div>
      </section>

      <section class="schulung-sec" id="sch-6">
        <h2>6 · Reports & Auswertungen</h2>
        <p>Im Reiter <b>Reports</b> filtern Sie alle Anfragen nach Status, Warengruppe,
           Mitarbeiter, Priorität, Beschaffungsart, Lieferant, Kostenstelle, Bestell-Nr.,
           Volumen und Zeiträumen – ähnlich wie in Excel.</p>
        <ul>
          <li>Spaltenüberschrift anklicken = sortieren.</li>
          <li>Zeile anklicken = Anfrage öffnen.</li>
          <li><b>„Als Excel exportieren"</b> lädt die aktuell gefilterte Liste als .xlsx herunter.</li>
        </ul>
      </section>

      <section class="schulung-sec" id="sch-7">
        <h2>7 · Wiedervorlage / Favoriten</h2>
        <p>Wiederkehrende Anfragen müssen Sie nicht neu tippen:</p>
        <ul>
          <li>In einer geöffneten Anfrage auf <b>„⭐ Als Wiedervorlage"</b> klicken.</li>
          <li>Die Anfrage wird als Favorit gespeichert (lokal in Ihrem Browser).</li>
          <li>Oben in <b>Neue Anfrage</b> bzw. <b>Sammelanfrage</b> erscheinen Ihre Favoriten –
              ein Klick füllt den Assistenten automatisch aus. Sammelanfragen landen dabei
              automatisch im richtigen Assistenten.</li>
        </ul>
      </section>

      <section class="schulung-sec" id="sch-8">
        <h2>8 · Tabellen-Importer</h2>
        <p>Aktualisieren Sie die Hinterlegungen für die Autovervollständigung:</p>
        <ul>
          <li><b>Artikelnummern</b> und <b>Kostenstellen</b> per CSV oder XLSX hochladen.</li>
          <li>Mit <b>„⬇ Export"</b> laden Sie die aktuelle Tabelle herunter, mit
              <b>„📄 Vorlage"</b> eine leere Beispieldatei.</li>
          <li>„↺ Zurücksetzen" stellt die Standardtabelle wieder her.</li>
        </ul>
      </section>

      <section class="schulung-sec" id="sch-9">
        <h2>9 · Häufige Fragen</h2>
        <p><b>Warum sehe ich keine Genehmigungs-Buttons?</b><br>
           Diese erscheinen nur im Dashboard / „Alle Anfragen", nicht in „Meine Anfragen".</p>
        <p><b>Meine Änderung ist weg?</b><br>
           Prüfen Sie, ob die Anfrage noch im Status „In Prüfung (Einkauf)" ist – nur dann ist
           sie bearbeitbar.</p>
        <p><b>Wo werden Favoriten gespeichert?</b><br>
           Lokal in Ihrem Browser. Bei Browserwechsel oder Cache-Löschung sind sie nicht mehr da.</p>
        <div class="schulung-tip">📨 Weitere Fragen? Wenden Sie sich an den Einkauf / IT.</div>
      </section>
    </div>`;
}

// ── SAMMELANFRAGE (MULTI-POSITION WIZARD) ────────────────────────────────────

let multiPositions = [];   // [{artNr, extArtNr, bezeichnung, menge, me, termin, kostenstelle, preis}, ...]
let multiWizardData = {};  // step2, step3

function initMultiWizard() {
  multiPositions = [{ artNr: '', extArtNr: '', bezeichnung: '', menge: '', me: '', termin: '', kostenstelle: '', preis: '' }];
  multiWizardData = {};
  showMultiStep(1);
  // Reset Allgemein fields
  const wg = $id('mf-Warengruppe'); if (wg) wg.value = '';
  const prioEl = $id('mf-Prioritaet');
  if (prioEl) {
    const prioOpts = [...prioEl.options];
    const normalOpt = prioOpts.find(o => /^(normal|standard)$/i.test(o.value.trim()));
    prioEl.value = normalOpt ? normalOpt.value : (prioOpts.find(o => o.value !== '')?.value || '');
  }
  const beschEl = $id('mf-Beschreibung'); if (beschEl) beschEl.value = '';
  const terminEl = $id('mf-Termin'); if (terminEl) terminEl.value = '';
  // Reset Beschaffung fields
  ['mf-Lieferant','mf-Lieferant2','mf-Lieferant3','mf-Lieferant4','mf-GeschaetzterPreis']
    .forEach(k => { const el = $id(k); if (el) el.value = ''; });
  const firstRadio = document.querySelector('input[name=mBeschaffungslogik]');
  if (firstRadio) firstRadio.checked = true;
  document.querySelectorAll('#m-beschaffungslogik-extra-cards .check-card').forEach(c => c.classList.remove('selected'));
  [2,3,4].forEach(n => {
    const grp = $id('m-lieferant-extra-' + n);
    if (grp) grp.style.display = 'none';
  });
  // Kostenstelle autocomplete
  // Reset submit button
  const sb = $id('btn-multi-submit');
  if (sb) { sb.disabled = false; sb.textContent = '✓ Sammelanfrage einreichen'; }
  renderMultiPositions();
}

function showMultiStep(n) {
  [1,2,3,4].forEach(i => {
    const body = $id('mwstep-' + i);
    if (body) body.classList.toggle('hidden', i !== n);
    const s = document.querySelector(`.wstep[data-mstep="${i}"]`);
    if (!s) return;
    s.classList.remove('active','done');
    if (i < n)  s.classList.add('done');
    if (i === n) s.classList.add('active');
  });
}

function renderMultiPositions() {
  const container = $id('multi-positions-table');
  if (!container) return;
  if (multiPositions.length === 0) multiPositions.push({ artNr: '', extArtNr: '', bezeichnung: '', menge: '', me: '', termin: '', kostenstelle: '', preis: '' });
  const meOptions = ['','Lagereinheiten','kg','Stück','Anzahl','m','Paar','Liter'];
  const rows = multiPositions.map((pos, i) => `
    <div class="multi-pos-row" data-idx="${i}">
      <div class="multi-pos-num">${i + 1}</div>
      <div class="multi-pos-fields">
        <div class="multi-pos-field">
          <label>Artikelnummer</label>
          <div class="tid-ac-wrap">
            <input type="text" class="mpos-artnr" data-idx="${i}"
              value="${esc(pos.artNr)}"
              placeholder="z. B. 4001-00010"
              oninput="multiPosChange(${i},'artNr',this.value)"
              autocomplete="off"/>
          </div>
        </div>
        <div class="multi-pos-field">
          <label>Externe ArtNr. <span class="field-sub">Katalog</span></label>
          <input type="text" class="mpos-extartnr" data-idx="${i}"
            value="${esc(pos.extArtNr || '')}"
            placeholder="Lieferanten-Nr."
            oninput="multiPosChange(${i},'extArtNr',this.value)"
            autocomplete="off"/>
        </div>
        <div class="multi-pos-field" style="flex:2">
          <label>Bezeichnung <span class="req">*</span></label>
          <input type="text" class="mpos-bez" data-idx="${i}"
            value="${esc(pos.bezeichnung)}"
            placeholder="Artikelbezeichnung"
            oninput="multiPosChange(${i},'bezeichnung',this.value)"/>
        </div>
        <div class="multi-pos-field" style="flex:0 0 90px">
          <label>Menge <span class="req">*</span></label>
          <input type="number" class="mpos-menge" data-idx="${i}"
            value="${esc(pos.menge)}"
            placeholder="1" min="0.001" step="any"
            oninput="multiPosChange(${i},'menge',this.value)"/>
        </div>
        <div class="multi-pos-field" style="flex:0 0 130px">
          <label>ME <span class="req">*</span></label>
          <select class="mpos-me" data-idx="${i}" onchange="multiPosChange(${i},'me',this.value)">
            ${meOptions.map(o => `<option value="${o}"${pos.me === o ? ' selected' : ''}>${o || '– wählen –'}</option>`).join('')}
          </select>
        </div>
        <div class="multi-pos-field" style="flex:0 0 110px">
          <label>Preis (€)</label>
          <input type="number" class="mpos-preis" data-idx="${i}"
            value="${esc(String(pos.preis || ''))}"
            placeholder="0.00" min="0" step="0.01"
            oninput="multiPosChange(${i},'preis',this.value); updateMultiTotal()"/>
        </div>
        <div class="multi-pos-field" style="flex:0 0 140px">
          <label>Benötigt bis</label>
          <input type="date" class="mpos-termin" data-idx="${i}"
            value="${esc(String(pos.termin || ''))}"
            oninput="multiPosChange(${i},'termin',this.value)"/>
        </div>
        <div class="multi-pos-field" style="flex:0 0 160px">
          <label>Kostenstelle</label>
          <div class="tid-ac-wrap">
            <input type="text" class="mpos-kostenstelle" data-idx="${i}"
              value="${esc(String(pos.kostenstelle || ''))}"
              placeholder="Nr. oder Bezeichnung…"
              oninput="multiPosChange(${i},'kostenstelle',this.value)"
              autocomplete="off"/>
          </div>
        </div>
      </div>
      ${multiPositions.length > 1
        ? `<button type="button" class="multi-pos-del" title="Position entfernen" onclick="removeMultiPosition(${i})">✕</button>`
        : '<div class="multi-pos-del-placeholder"></div>'}
    </div>`).join('');
  container.innerHTML = rows;
  // Init TID autocomplete for each artnr input
  container.querySelectorAll('.mpos-artnr').forEach(inp => {
    const idx = parseInt(inp.dataset.idx, 10);
    initTidAutocomplete(inp, key => {
      if (key === 'Artikelnummer') return inp;
      if (key === 'Title') {
        // Proxy: sets this row's Bezeichnung
        return {
          get value() { return multiPositions[idx]?.bezeichnung || ''; },
          set value(v) {
            if (multiPositions[idx]) {
              multiPositions[idx].bezeichnung = v;
              const bezEl = document.querySelector(`.mpos-bez[data-idx="${idx}"]`);
              if (bezEl) bezEl.value = v;
            }
          }
        };
      }
      if (key === 'Warengruppe') {
        // Fill global Warengruppe in step 2
        return $id('mf-Warengruppe');
      }
      return null;
    });
  });
  // Init Kostenstelle autocomplete for each row
  container.querySelectorAll('.mpos-kostenstelle').forEach(inp => {
    const idx = parseInt(inp.dataset.idx, 10);
    initKostenstAuto(inp);
    inp.addEventListener('change', e => { multiPosChange(idx, 'kostenstelle', e.target.value); });
  });
  updateMultiTotal();
}

function updateMultiTotal() {
  const total = multiPositions.reduce((s, p) => s + (parseFloat(p.preis) || 0), 0);
  const el = $id('mf-GeschaetzterPreis');
  if (el) el.value = total > 0 ? total.toFixed(2) : '';
  const displayEl = $id('mf-total-display');
  if (displayEl) displayEl.textContent = total > 0 ? fmtEuro(total) : '–';
}

function multiPosChange(idx, field, val) {
  if (multiPositions[idx]) multiPositions[idx][field] = val;
}

function addMultiPosition() {
  multiPositions.push({ artNr: '', extArtNr: '', bezeichnung: '', menge: '', me: '', termin: '', kostenstelle: '', preis: '' });
  renderMultiPositions();
}

function removeMultiPosition(idx) {
  multiPositions.splice(idx, 1);
  if (multiPositions.length === 0) multiPositions.push({ artNr: '', extArtNr: '', bezeichnung: '', menge: '', me: '', termin: '', kostenstelle: '', preis: '' });
  renderMultiPositions();
}

function toggleMBLExtra(el) {
  el.classList.toggle('selected');
}

function addMLieferant(n) {
  const grp = $id('m-lieferant-extra-' + n);
  if (grp) grp.style.display = '';
  const prevBtn = $id('mbtn-add-lieferant-' + (n - 1));
  if (prevBtn) prevBtn.style.display = 'none';
}

function wMultiNext(step) {
  if (step === 1) {
    // Validate all positions
    for (let i = 0; i < multiPositions.length; i++) {
      const p = multiPositions[i];
      // sync values from DOM (in case oninput didn't fire on all)
      const bezEl   = document.querySelector(`.mpos-bez[data-idx="${i}"]`);
      const mengeEl = document.querySelector(`.mpos-menge[data-idx="${i}"]`);
      const meEl    = document.querySelector(`.mpos-me[data-idx="${i}"]`);
      const artEl    = document.querySelector(`.mpos-artnr[data-idx="${i}"]`);
      const extArtEl = document.querySelector(`.mpos-extartnr[data-idx="${i}"]`);
      if (bezEl)    p.bezeichnung = bezEl.value.trim();
      if (mengeEl)  p.menge = mengeEl.value;
      if (meEl)     p.me = meEl.value;
      if (artEl)    p.artNr = artEl.value.trim();
      if (extArtEl) p.extArtNr = extArtEl.value.trim();
      const terminEl = document.querySelector(`.mpos-termin[data-idx="${i}"]`);
      const kstEl    = document.querySelector(`.mpos-kostenstelle[data-idx="${i}"]`);
      const preisEl  = document.querySelector(`.mpos-preis[data-idx="${i}"]`);
      if (terminEl) p.termin      = terminEl.value;
      if (kstEl)    p.kostenstelle = kstEl.value.trim();
      if (preisEl)  p.preis       = preisEl.value;
      if (!p.bezeichnung) { toast(`Position ${i + 1}: Bitte Bezeichnung angeben.`, 'error'); return; }
      if (!p.menge || parseFloat(p.menge) <= 0) { toast(`Position ${i + 1}: Bitte gültige Menge eingeben.`, 'error'); return; }
      if (!p.me)   { toast(`Position ${i + 1}: Bitte Mengeneinheit wählen.`, 'error'); return; }
    }
  } else if (step === 2) {
    const wg = $id('mf-Warengruppe').value;
    if (!wg) { toast('Bitte Warengruppe wählen.', 'error'); return; }
    multiWizardData.step2 = {
      Warengruppe:  wg,
      Prioritaet:   $id('mf-Prioritaet').value,
      Beschreibung: $id('mf-Beschreibung').value.trim(),
    };
  } else if (step === 3) {
    const lieferant = $id('mf-Lieferant').value.trim();
    if (!lieferant) { toast('Bitte mindestens Lieferant 1 angeben.', 'error'); return; }
    const extraSelected = document.querySelector('#m-beschaffungslogik-extra-cards .check-card.selected');
    if (!extraSelected) { toast('Bitte unter „Zusätzlich kombinierbar" eine Option auswählen.', 'error'); return; }
    multiWizardData.step3 = {
      Beschaffungslogik: [
        document.querySelector('input[name=mBeschaffungslogik]:checked')?.value || '',
        ...[...document.querySelectorAll('#m-beschaffungslogik-extra-cards .check-card.selected')].map(c => c.dataset.value)
      ].filter(Boolean).join(', '),
      Lieferant:         $id('mf-Lieferant').value.trim(),
      Lieferant2:        $id('mf-Lieferant2').value.trim(),
      Lieferant3:        $id('mf-Lieferant3').value.trim(),
      Lieferant4:        $id('mf-Lieferant4').value.trim(),
      GeschaetzterPreis: $id('mf-GeschaetzterPreis').value ? parseFloat($id('mf-GeschaetzterPreis').value) : null,
    };
    buildMultiReview();
  }
  showMultiStep(step + 1);
}

function wMultiBack(step) {
  showMultiStep(step - 1);
}

function buildMultiReview() {
  const s2 = multiWizardData.step2 || {};
  const s3 = multiWizardData.step3 || {};
  const posRows = multiPositions.map((p, i) => `
    <tr>
      <td style="padding:4px 8px;color:#6b7280">${i + 1}</td>
      <td style="padding:4px 8px">${esc(p.artNr) || '–'}</td>
      <td style="padding:4px 8px;color:#9ca3af">${esc(p.extArtNr) || '–'}</td>
      <td style="padding:4px 8px">${esc(p.bezeichnung)}</td>
      <td style="padding:4px 8px;text-align:right">${esc(p.menge)}</td>
      <td style="padding:4px 8px">${esc(p.me)}</td>
      <td style="padding:4px 8px;text-align:right">${p.preis ? fmtEuro(p.preis) : '–'}</td>
      <td style="padding:4px 8px">${p.termin || '–'}</td>
      <td style="padding:4px 8px">${esc(p.kostenstelle) || '–'}</td>
    </tr>`).join('');
  const liefs = [s3.Lieferant, s3.Lieferant2, s3.Lieferant3, s3.Lieferant4].filter(Boolean).join(', ');
  const preis = s3.GeschaetzterPreis != null ? new Intl.NumberFormat('de-DE', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(s3.GeschaetzterPreis) + ' €' : '–';
  $id('multi-review-content').innerHTML = `
    <div class="review-section">
      <h3 class="review-section-title">Positionen (${multiPositions.length})</h3>
      <div style="overflow-x:auto">
        <table style="width:100%;border-collapse:collapse;font-size:.875rem">
          <thead><tr style="background:#f3f4f6;font-weight:600">
            <th style="padding:6px 8px;text-align:left">#</th>
            <th style="padding:6px 8px;text-align:left">ArtNr.</th>
            <th style="padding:6px 8px;text-align:left">Ext. ArtNr.</th>
            <th style="padding:6px 8px;text-align:left">Bezeichnung</th>
            <th style="padding:6px 8px;text-align:right">Menge</th>
            <th style="padding:6px 8px;text-align:left">ME</th>
            <th style="padding:6px 8px;text-align:right">Preis (€)</th>
            <th style="padding:6px 8px;text-align:left">Termin</th>
            <th style="padding:6px 8px;text-align:left">KST</th>
          </tr></thead>
          <tbody>${posRows}</tbody>
        </table>
      </div>
    </div>
    <div class="review-section" style="margin-top:16px">
      <h3 class="review-section-title">Allgemein</h3>
      <dl class="review-dl">
        <dt>Warengruppe</dt><dd>${esc(s2.Warengruppe)}</dd>
        <dt>Priorität</dt><dd>${esc(s2.Prioritaet)}</dd>
        <dt>Beschreibung</dt><dd>${esc(s2.Beschreibung) || '–'}</dd>
      </dl>
    </div>
    <div class="review-section" style="margin-top:16px">
      <h3 class="review-section-title">Beschaffung</h3>
      <dl class="review-dl">
        <dt>Beschaffungsart</dt><dd>${esc(s3.Beschaffungslogik)}</dd>
        <dt>Lieferant(en)</dt><dd>${esc(liefs) || '–'}</dd>
        <dt>Bestellvolumen</dt><dd>${preis}</dd>
      </dl>
    </div>`;
}

async function submitMultiRequest() {
  const btn = $id('btn-multi-submit');
  btn.disabled = true;
  btn.textContent = 'Wird eingereicht…';
  try {
    const s2 = multiWizardData.step2 || {};
    const s3 = multiWizardData.step3 || {};
    const first = multiPositions[0] || {};

    // Auto-generate title from first position
    const autoTitle = first.bezeichnung
      ? (multiPositions.length > 1
          ? `Sammelanfrage: ${first.bezeichnung} +${multiPositions.length - 1} weitere`
          : `Sammelanfrage: ${first.bezeichnung}`)
      : 'Sammelanfrage';

    // Serialize all positions as JSON for the Positionen SP field
    const posJson = JSON.stringify(multiPositions.map((p, i) => ({
      Nr: i + 1,
      Artikelnummer: p.artNr,
      ExterneArtikelnummer: p.extArtNr || '',
      Bezeichnung: p.bezeichnung,
      Menge: p.menge,
      ME: p.me,
      Preis: p.preis ? parseFloat(p.preis) : null,
      Termin: p.termin || '',
      Kostenstelle: p.kostenstelle || '',
    })));

    const totalPreis = multiPositions.reduce((s,p) => s + (parseFloat(p.preis)||0), 0);

    const rawData = {
      Title:             autoTitle,
      Artikelnummer:     first.artNr || '',
      Menge:             first.menge || '',
      Mengeneinheit:     first.me || '',
      Warengruppe:       s2.Warengruppe || '',
      Prioritaet:        s2.Prioritaet || '',
      Beschreibung:      s2.Beschreibung || '',
      Termin:            '',
      Beschaffungslogik: s3.Beschaffungslogik || '',
      Lieferant:         s3.Lieferant || '',
      Lieferant2:        s3.Lieferant2 || '',
      Lieferant3:        s3.Lieferant3 || '',
      Lieferant4:        s3.Lieferant4 || '',
      GeschaetzterPreis: totalPreis > 0 ? totalPreis : (s3.GeschaetzterPreis || null),
    };

    const allFields = buildFields(rawData, FORM_FIELDS);

    if (!siteId || !listId) await discoverSP();

    const listPath = `/sites/${siteId}/lists/${listId}`;

    btn.textContent = 'Anfrage wird angelegt…';
    const newItem = await gPost(`${listPath}/items`, { fields: { Title: allFields['Title'] } });
    const itemId  = newItem.id;

    // Patch remaining fields + Positionen (JSON multi-line field)
    const patchFields = { ...allFields };
    delete patchFields['Title'];

    // Add Positionen field if SP field exists
    const posField = resolvedFields['Positionen'] || 'Positionen';
    patchFields[posField] = posJson;

    if (Object.keys(patchFields).length) {
      btn.textContent = 'Felder werden gespeichert…';
      try {
        await patchRetry(`${listPath}/items/${itemId}/fields`, patchFields);
      } catch(patchErr) {
        console.warn('[submitMultiRequest] PATCH fehlgeschlagen:', patchErr.message);
        toast('Anfrage erstellt, aber einige Felder konnten nicht gespeichert werden.', 'error');
      }
    }

    toast(`Sammelanfrage mit ${multiPositions.length} Positionen eingereicht! Power Automate startet den Genehmigungsprozess.`, 'success');
    await loadItems(false);
    navigate('mine');
  } catch(e) {
    toast('Fehler: ' + e.message, 'error');
    btn.disabled = false;
    btn.textContent = '✓ Sammelanfrage einreichen';
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
  // Overall item status is the primary source of truth
  const sv = (getStatusVal(item) || '').trim();
  if (/^abgelehnt$/i.test(sv)) return `<span class="ica-no">✗ Abgelehnt</span>`;
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
  // If the overall status is "Abgelehnt", show that regardless of field values
  const sv = (getStatusVal(item) || '').trim();
  if (/^abgelehnt$/i.test(sv)) {
    return `<div class="cr-appr cr-appr-no"><span class="cr-appr-icon">✗</span><div><strong>Abgelehnt</strong></div></div>`;
  }
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
async function openPanel(itemId) {
  if (!['mine','all','dashboard'].includes(currentView)) { navigate('detail', itemId); return; }
  panelItemId = String(itemId);
  const item  = allItems.find(i => String(i.id) === panelItemId);
  if (!item) return;

  // Auto-advance: läuft unabhängig von View und Rolle für jeden Nutzer
  {
    const sv        = (getStatusVal(item) || 'Eingereicht').trim(); // leeres Feld = Eingereicht
    const statusCol = resolvedFields['Status'] || 'Status';
    let   advanced  = false;

    // Eingereicht → In Prüfung (Einkauf)
    if (/^eingereicht$/i.test(sv)) {
      const target = statusChoices.find(c => /pr[üu]fung/i.test(c) && /einkauf/i.test(c) && !/strategisch/i.test(c))
                  || 'In Prüfung (Einkauf)';
      console.log('[auto-advance] sv=', JSON.stringify(sv), 'statusCol=', statusCol, 'target=', target, 'choices=', statusChoices);
      try {
        await gPatch(`/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, { [statusCol]: target });
        // Optimistisch allItems sofort aktualisieren (Graph-Cache umgehen)
        const cached = allItems.find(i => String(i.id) === String(itemId));
        if (cached?.fields) cached.fields[statusCol] = target;
        advanced = true;
        console.log('[auto-advance] PATCH OK → ' + target);
      } catch(e) {
        console.error('[auto-advance] PATCH fehlgeschlagen:', e.message);
        toast('Status-Update fehlgeschlagen: ' + e.message, 'error');
      }
    }

    // Freigegeben → In Bestellung
    if (/^freigegeben$/i.test(sv)) {
      const inBestKey = statusChoices.find(c => /^in bestellung$/i.test(c.trim()));
      if (inBestKey) {
        try {
          await gPatch(`/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, { [statusCol]: inBestKey });
          const cached = allItems.find(i => String(i.id) === String(itemId));
          if (cached?.fields) cached.fields[statusCol] = inBestKey;
          advanced = true;
        } catch(e) { console.warn('[openPanel] auto-advance Freigegeben failed:', e.message); }
      }
    }

    if (advanced) await loadItems(false);
  }

  // Always use the freshest item from allItems after potential reload
  const freshItem = allItems.find(i => String(i.id) === panelItemId) || item;

  $id(`panel-${currentView}-content`).innerHTML = renderPanel(freshItem);
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
  // IMPORTANT: scope all DOM queries to the CURRENT panel container.
  // Multiple views (mine, dashboard) each render renderPanel() and produce identical
  // IDs (panel-attach-body, panel-history-body, …). document.getElementById() would
  // return the first match — usually the wrong panel. Use scoped queries instead.
  const panelRoot = $id(`panel-${currentView}-content`);
  const pq = sel => panelRoot?.querySelector(sel);

  pq('#panel-close')?.addEventListener('click', async () => {
    await releaseLock(itemId);
    closePanel();
  });
  pq('#panel-edit')?.addEventListener('click', () => startEditMode(itemId));
  pq('#panel-order')?.addEventListener('click', () => openOrderModal(itemId));

  // Genehmigungsaktion-Buttons
  pq('#btn-approve')?.addEventListener('click', () => doApprove(itemId));
  pq('#btn-reject')?.addEventListener('click', () => {
    pq('#aab-reject-form').style.display = '';
    pq('#btn-approve').style.display = 'none';
    pq('#btn-reject').style.display  = 'none';
  });
  pq('#btn-reject-cancel')?.addEventListener('click', () => {
    pq('#aab-reject-form').style.display = 'none';
    pq('#btn-approve').style.display = '';
    pq('#btn-reject').style.display  = '';
  });
  pq('#btn-reject-confirm')?.addEventListener('click', () =>
    doReject(itemId, pq('#aab-reject-comment')?.value?.trim() || ''));
  pq('#panel-save')?.addEventListener('click', () => savePanelEdits(itemId, panelRoot));
  pq('#panel-cancel')?.addEventListener('click', async () => {
    await releaseLock(itemId);
    const item = allItems.find(i => String(i.id) === String(itemId));
    if (item) { panelRoot.innerHTML = renderPanel(item); bindPanelEvents(itemId); }
  });

  // Artikelnummer autocomplete in edit mode
  const tidInput = pq('[data-ac="tid"]');
  if (tidInput) {
    initTidAutocomplete(tidInput, key => pq(`.pf-input[data-key="${key}"]`));
  }

  // Approvers are read directly from item.fields — no async history load needed.

  // Load attachments — scoped to this panel
  const attachEl = pq('#panel-attach-body');
  if (attachEl) {
    getSpToken()
      .then(tok => fetch(
        `${SP_BASE}/_api/web/lists/getByTitle('${SP_LIST}')/items(${itemId})/AttachmentFiles?_=${Date.now()}`,
        { headers: { Authorization: 'Bearer ' + tok, Accept: 'application/json;odata=nometadata',
            'Cache-Control': 'no-cache', Pragma: 'no-cache' } }
      ))
      .then(r => r.ok ? r.json() : { value: [] })
      .then(data => {
        const files = data.value || [];
        attachEl.innerHTML = files.length
          ? files.map(attachmentLink).join('')
          : '<span class="no-order">Keine Anhänge.</span>';
      })
      .catch(() => { attachEl.innerHTML = '<span class="no-order">Anhänge konnten nicht geladen werden.</span>'; });
  }

  // Kommentare aus SP-Listenfeld lesen (Mehrere Textzeilen, Format: "[Datum – Autor]: Text\n---\n")
  const commentsSection = pq('#panel-comments-section');
  const commentsEl      = pq('#panel-comments-body');

  function loadComments() {
    if (!commentsEl) return;
    if (commentsSection) commentsSection.style.display = '';
    const cur  = allItems.find(i => String(i.id) === String(itemId));
    const raw  = getField(cur, getKommentarCol());
    const entries = parseKommentare(raw);
    if (!entries.length) {
      commentsEl.innerHTML = '<span class="no-order">Noch keine Kommentare.</span>';
      return;
    }
    commentsEl.innerHTML = [...entries].reverse().map(c =>
      `<div class="pf-comment-row">
        <div class="pf-comment-header"><span class="pf-comment-author">💬 ${esc(c.header || '–')}</span></div>
        <div class="pf-comment-text">${esc(c.text)}</div>
      </div>`
    ).join('');
  }

  loadComments();

  // Send button — post new comment then reload
  pq('#panel-comment-send')?.addEventListener('click', async () => {
    const inp = pq('#panel-comment-input');
    const text = inp?.value?.trim();
    if (!text) return;
    const btn = pq('#panel-comment-send');
    btn.disabled = true; btn.textContent = '…';
    const ok = await postSpComment(itemId, text);
    btn.disabled = false; btn.textContent = '💬 Senden';
    if (ok) { if (inp) inp.value = ''; await loadComments(); }
    else toast('Kommentar konnte nicht gespeichert werden.', 'error');
  });

  // Ctrl+Enter shortcut in textarea
  pq('#panel-comment-input')?.addEventListener('keydown', e => {
    if (e.key === 'Enter' && e.ctrlKey) pq('#panel-comment-send')?.click();
  });

  pq('#panel-history')?.addEventListener('click', () => {
    const sec = pq('#panel-history-section');
    if (!sec) return;
    const isOpen = sec.style.display !== 'none';
    sec.style.display = isOpen ? 'none' : '';
    pq('#panel-history').textContent = isOpen ? '📋 Verlauf' : '📋 Verlauf ▲';
    if (!isOpen) loadVersionHistory(itemId, pq);
  });
}

let _vhLoaded = {};
async function loadVersionHistory(itemId, pq) {
  // pq: scoped query function from bindPanelEvents (avoids duplicate-ID collision)
  const el = pq ? pq('#panel-history-body') : $id('panel-history-body');
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

async function savePanelEdits(itemId, panelRoot) {
  const data = {};
  (panelRoot || document).querySelectorAll('.pf-input[data-key]').forEach(inp => {
    data[inp.dataset.key] = inp.value;
  });
  const fields = buildFields(data, FORM_FIELDS);
  if (!Object.keys(fields).length) return;

  const btn = panelRoot?.querySelector('#panel-save') || $id('panel-save');
  if (btn) { btn.disabled = true; btn.textContent = 'Speichert…'; }

  const patch = { ...fields };
  const skipped = [];
  for (let i = 0; i < 15; i++) {
    try {
      await gPatch(`/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, patch);
      await releaseLock(itemId); // clear edit lock on successful save
      if (skipped.length) toast(`Gespeichert (übersprungen: ${skipped.join(', ')})`, 'info');
      else toast('Gespeichert ✓', 'success');
      await loadItems(false);
      return;
    } catch(e) {
      const m = e.message.match(/Field '([^']+)' (?:is not recognized|does not exist)/i);
      if (!m) { toast('Fehler: ' + e.message, 'error'); if (btn) { btn.disabled=false; btn.textContent='💾 Speichern'; } return; }
      skipped.push(m[1]); delete patch[m[1]];
    }
  }
  toast('Fehler: Zu viele unbekannte Felder.', 'error');
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

// Build the inner HTML for the "Status & Genehmigung" approval section.
// ── GENEHMIGUNG IM DASHBOARD ─────────────────────────────────────────────────

async function doApprove(itemId) {
  const banner = document.querySelector('#approval-action-banner');
  if (banner) banner.innerHTML = '<div class="aab-hint">⏳ Wird gespeichert…</div>';

  const statusCol   = resolvedFields['Status']               || 'Status';
  const approverCol = resolvedFields['Entscheider_x002a_in'] || 'Entscheider_x002a_in';
  const commentCol  = resolvedFields['Genehmigungskommentar']|| 'Genehmigungskommentar';

  const approverName = account?.name || account?.username || '';
  const genehmigt    = 'Genehmigt (Einkauf)';

  // Find the exact SP choice string (case-sensitive)
  const choiceMatch = statusChoices.find(c => /^genehmigt.*einkauf/i.test(c.trim())) || genehmigt;

  try {
    await gPatch(`/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, {
      [statusCol]:   choiceMatch,
      [approverCol]: approverName,
    });

    // Update local cache
    const item = allItems.find(i => String(i.id) === String(itemId));
    if (item?.fields) {
      item.fields[statusCol]   = choiceMatch;
      item.fields[approverCol] = approverName;
    }

    toast('✅ Genehmigung erteilt.', 'success');
    await loadItems(false);
    // Re-render panel with updated item
    const updated = allItems.find(i => String(i.id) === String(itemId));
    if (updated) {
      const panelRoot = $id(`panel-${currentView}-content`);
      if (panelRoot) { panelRoot.innerHTML = renderPanel(updated); bindPanelEvents(itemId); }
    }
  } catch(e) {
    toast('Fehler beim Genehmigen: ' + e.message, 'error');
    // Restore banner
    const item = allItems.find(i => String(i.id) === String(itemId));
    if (item) {
      const panelRoot = $id(`panel-${currentView}-content`);
      if (panelRoot) { panelRoot.innerHTML = renderPanel(item); bindPanelEvents(itemId); }
    }
  }
}

async function doReject(itemId, comment) {
  const banner = document.querySelector('#approval-action-banner');
  if (banner) banner.innerHTML = '<div class="aab-hint">⏳ Wird gespeichert…</div>';

  const statusCol   = resolvedFields['Status']               || 'Status';
  const approverCol = resolvedFields['Entscheider_x002a_in'] || 'Entscheider_x002a_in';
  const commentCol  = resolvedFields['Genehmigungskommentar']|| 'Genehmigungskommentar';

  const approverName = account?.name || account?.username || '';
  const abgelehnt    = statusChoices.find(c => /^abgelehnt$/i.test(c.trim())) || 'Abgelehnt';

  const patch = {
    [statusCol]:   abgelehnt,
    [approverCol]: approverName,
  };
  if (comment && commentCol) patch[commentCol] = comment;

  try {
    await gPatch(`/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, patch);

    const item = allItems.find(i => String(i.id) === String(itemId));
    if (item?.fields) {
      item.fields[statusCol]   = abgelehnt;
      item.fields[approverCol] = approverName;
      if (comment) item.fields[commentCol] = comment;
    }

    toast('❌ Anfrage abgelehnt.', 'success');
    await loadItems(false);
    const updated = allItems.find(i => String(i.id) === String(itemId));
    if (updated) {
      const panelRoot = $id(`panel-${currentView}-content`);
      if (panelRoot) { panelRoot.innerHTML = renderPanel(updated); bindPanelEvents(itemId); }
    }
  } catch(e) {
    toast('Fehler beim Ablehnen: ' + e.message, 'error');
    const item = allItems.find(i => String(i.id) === String(itemId));
    if (item) {
      const panelRoot = $id(`panel-${currentView}-content`);
      if (panelRoot) { panelRoot.innerHTML = renderPanel(item); bindPanelEvents(itemId); }
    }
  }
}

// Extracted so it can be called both from renderPanel and from the async
// loadApproverHistory refresh without re-rendering the whole panel.
function buildApprovalInner(item, statusVal) {
  const sv = statusVal || getStatusVal(item) || 'Eingereicht';
  // statusTimeline reads all approver/comment fields directly from item.fields —
  // it handles all stages including "In Bestellung" upcoming/past logic.
  return `<div class="approval-stages">${statusTimeline(sv, item)}</div>`;
}

// Panel-positions in-place editing
let _panelPosData = []; // working copy while editing
function _syncPanelPosFromDOM() {
  const tbody = $id('pos-edit-tbody');
  if (!tbody) return;
  tbody.querySelectorAll('input.pos-edit-input, select.pos-edit-input').forEach(el => {
    const idx = parseInt(el.dataset.idx, 10);
    const field = el.dataset.field;
    if (_panelPosData[idx] !== undefined) _panelPosData[idx][field] = el.value;
  });
}
function panelPosAdd() {
  _syncPanelPosFromDOM();
  _panelPosData.push({ Nr: _panelPosData.length + 1, Artikelnummer:'', ExterneArtikelnummer:'', Bezeichnung:'', Menge:'', ME:'', Preis:'', Termin:'', Kostenstelle:'' });
  _rebuildPosEditTable();
}
function panelPosDelete(idx) {
  _syncPanelPosFromDOM();
  _panelPosData.splice(idx, 1);
  _panelPosData.forEach((p,i) => p.Nr = i + 1);
  _rebuildPosEditTable();
}
function _rebuildPosEditTable() {
  const tbody = $id('pos-edit-tbody');
  if (!tbody) return;
  const meOpts = ['Lagereinheiten','kg','Stück','Anzahl','m','Paar','Liter'];
  tbody.innerHTML = _panelPosData.map((p, idx) => {
    const meSelOpts = meOpts.map(o => `<option value="${o}"${(p.ME||'')=== o?' selected':''}>${o}</option>`).join('');
    return `<tr class="pos-edit-row" data-pos-idx="${idx}">
      <td class="pos-td pos-nr">${p.Nr}</td>
      <td class="pos-td"><input class="pos-edit-input" data-idx="${idx}" data-field="Artikelnummer" value="${esc(String(p.Artikelnummer||''))}" placeholder="ArtNr."/></td>
      <td class="pos-td"><input class="pos-edit-input" data-idx="${idx}" data-field="ExterneArtikelnummer" value="${esc(String(p.ExterneArtikelnummer||''))}" placeholder="Ext.ArtNr."/></td>
      <td class="pos-td pos-bez"><input class="pos-edit-input" data-idx="${idx}" data-field="Bezeichnung" value="${esc(String(p.Bezeichnung||''))}" placeholder="Bezeichnung"/></td>
      <td class="pos-td pos-right"><input class="pos-edit-input pos-edit-num" type="number" data-idx="${idx}" data-field="Menge" value="${esc(String(p.Menge||''))}" placeholder="1" min="0.001" step="any"/></td>
      <td class="pos-td"><select class="pos-edit-input pos-edit-me" data-idx="${idx}" data-field="ME"><option value="">–</option>${meSelOpts}</select></td>
      <td class="pos-td pos-right"><input class="pos-edit-input pos-edit-num" type="number" data-idx="${idx}" data-field="Preis" value="${esc(String(p.Preis||''))}" placeholder="0,00" min="0" step="0.01"/></td>
      <td class="pos-td"><input class="pos-edit-input" type="date" data-idx="${idx}" data-field="Termin" value="${esc(String(p.Termin||''))}"/></td>
      <td class="pos-td"><input class="pos-edit-input" data-idx="${idx}" data-field="Kostenstelle" value="${esc(String(p.Kostenstelle||''))}" placeholder="KST"/></td>
      <td class="pos-td"><button type="button" class="pos-del-btn" onclick="panelPosDelete(${idx})" title="Entfernen">✕</button></td>
    </tr>`;
  }).join('');
  // Update count header
  const title = document.querySelector('#pf-pos-section .pf-sec-title');
  if (title) title.textContent = `Positionen (${_panelPosData.length})`;
}
async function panelPosSave(itemId) {
  _syncPanelPosFromDOM();
  const posField = resolvedFields['Positionen'] || 'Positionen';
  const priceField = resolvedFields['GeschaetzterPreis'] || 'GeschaetzterPreis';
  const totalPrice = _panelPosData.reduce((s,p) => s + (parseFloat(p.Preis)||0), 0);
  try {
    await gPatch(`/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, {
      [posField]: JSON.stringify(_panelPosData),
      [priceField]: totalPrice || null,
    });
    // Update allItems optimistically
    const cached = allItems.find(i => String(i.id) === String(itemId));
    if (cached?.fields) {
      cached.fields[posField] = JSON.stringify(_panelPosData);
      if (totalPrice) cached.fields[priceField] = totalPrice;
    }
    toast('Positionen gespeichert ✓', 'success');
    await loadItems(false);
  } catch(e) {
    toast('Speichern fehlgeschlagen: ' + e.message, 'error');
  }
}

function renderPanel(item, editMode = false) {
  const isMineView = currentView === 'mine';
  const statusVal  = getStatusVal(item) || 'Eingereicht';
  const createdBy  = item.createdBy?.user?.displayName || item.createdBy?.user?.email || '–';
  const createdAt = item.createdDateTime ? fmtDate(item.createdDateTime) : '–';

  const gv  = key => getField(item, resolvedFields[key] || key) ?? '';
  const dv  = v   => v ? String(v).slice(0,10) : '';  // YYYY-MM-DD for date inputs

  const choices = key => {
    const sp = resolvedFields[key] || key;
    return colByKey[sp]?.choice?.choices || null;
  };

  const WG_OPTS   = choices('Warengruppe')      || ['4001 – Schrott','4002 – Legierungen','4103 – Bleche','4116 – Kaufteile','4211 – Kaufteile Formerei (Kühlm. ..)','4304 – Strahlmittel','4305 – Modellbaumaterial','4306 – Formstoffkomponenten','4307 – Eingußsysteme','4308 – exotherme Kappen (Speiser)','4309 – Gießfilter','4310 – Kerne (bez. Kerne f. Gießerei)','4311 – metallurgische Zusatzstoffe','4312 – Hilfsstoffe Gießerei','4313 – Feuerfestmaterial Ofen/Pfannen','4314 – Kohleprodukte (Grafitelektr., Aufk...)','4317 – externer Modellbau','4318 – Befestigungsteile','4319 – WENDESCHNEIDPLATTE','4320 – Ersatzteile','4321 – Schleif- und Fräsmittel','4322 – Arbeitsschutzmittel','4323 – Schweißbedarf','4324 – Gemeinkostenmaterial','4325 – Technische Gase','4326 – Treibstoffe, Öle, Fette','4328 – Werkzeug','4329 – sonst. Hilfs- u. Betriebsstoffe','4333 – Kooperation','4334 – Reparatur','4335 – Serviceleistung','4336 – Entsorgung','4337 – Büromaterial','4338 – Mieten','4340 – Investitionen','4333 – Transporte','4335 – Leiharbeiter','2000 – Frachtkosten, sonstiges'];
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
      } else if (fd.key === 'Artikelnummer') {
        const val = esc(String(raw));
        const hit = TID_MAP_ACTIVE[String(raw).trim().toUpperCase()] || null;
        const hint = hit ? `<div class="tid-ac-confirm">✓ ${esc(hit.b)} · ${esc(hit.w)}</div>` : '';
        inp = `<div class="tid-ac-wrap">
          <input type="text" class="pf-input" data-key="Artikelnummer" data-ac="tid"
            value="${val}" autocomplete="off" placeholder="Nr. oder Bezeichnung…"/>
          <div class="tid-ac-dropdown" style="display:none"></div>
          ${hint}
        </div>`;
      } else {
        const val = type === 'date' ? dv(raw) : esc(String(raw));
        const isTsd = fd.key === 'Menge' || fd.key === 'Mindestlagermenge';
        const stepAttr = type === 'number' ? (isTsd ? ' step="0.001"' : ' step="0.01"') : '';
        inp = `<input type="${type}" class="pf-input" data-key="${fd.key}" value="${val}"${stepAttr}/>`;
      }
      return `<div class="pf-row">${lbl}${inp}</div>`;
    }
    if (!raw && raw !== 0) return '';
    let display = String(raw);
    if (fd.key === 'GeschaetzterPreis' || fd.key === 'TatsaechlicherPreis') display = fmtEuro(raw);
    else if (fd.key === 'Termin' || fd.key === 'Lieferdatum') display = fmtDate(raw);
    else if (fd.key === 'Menge' || fd.key === 'Mindestlagermenge') {
      const n = parseFloat(raw);
      display = isNaN(n) ? String(raw) : n.toLocaleString('de-DE', { minimumFractionDigits: 0, maximumFractionDigits: 3 });
    }
    return `<div class="pf-row">${lbl}<span class="pf-val">${esc(display)}</span></div>`;
  };

  const preis = parseFloat(gv('GeschaetzterPreis')) || 0;
  const menge = parseFloat(gv('Menge')) || 1;
  const gesamtHint = !editMode && preis > 0 ? genehmigungsweg(preis) : '';

  const orderNr  = gv('Bestellnummer');
  const lieferd  = gv('Lieferdatum');
  const tatPreis = gv('TatsaechlicherPreis');

  const isFreigegeben = /^freigegeben$/i.test((statusVal || '').trim());
  const orderBtn = isMineView ? '' : isFreigegeben
    ? `<button class="btn btn-outline btn-sm" id="panel-order">📦 Einkauf-Daten</button>`
    : `<button class="btn btn-outline btn-sm" id="panel-order" disabled title="Nur bei Status 'Freigegeben' möglich">📦 Einkauf-Daten</button>`;

  // Lock state for the edit button
  const lockField   = getLockField();
  const lockData    = parseLock(lockField ? (getField(item, lockField) || '') : '');
  const isLocked    = lockData && lockData.expiresAt > new Date();
  const me          = account?.name || account?.username || '';
  const lockedByMe  = isLocked && lockData.by === me;
  const lockBadge   = isLocked && !lockedByMe
    ? `<span class="panel-lock-badge">🔒 ${esc(lockData.by)} bearbeitet gerade</span>` : '';

  const canEdit = /pr[üu]fung/i.test(statusVal) && /einkauf/i.test(statusVal) && !/strategisch/i.test(statusVal);
  const editBtn = isMineView ? '' : (isLocked && !lockedByMe)
    ? `<button class="btn btn-outline btn-sm" disabled title="Gesperrt von ${esc(lockData.by)}">🔒 Gesperrt</button>`
    : canEdit
      ? `<button class="btn btn-outline btn-sm" id="panel-edit">✏️ Bearbeiten</button>`
      : `<button class="btn btn-outline btn-sm" disabled title="Nur bei Status 'In Prüfung (Einkauf)' möglich">✏️ Bearbeiten</button>`;

  const buttons = editMode
    ? `${orderBtn}
       <button class="btn btn-primary btn-sm" id="panel-save">💾 Speichern</button>
       <button class="btn btn-ghost btn-sm" id="panel-cancel">✕ Abbrechen</button>
       <button class="btn btn-outline btn-sm" id="panel-history">📋 Verlauf</button>`
    : `${editBtn} ${orderBtn}
       <button class="btn btn-outline btn-sm" id="panel-history">📋 Verlauf</button>`;

  const approvalInner = buildApprovalInner(item, statusVal);

  // Genehmigungsaktion-Banner: sichtbar für alle im Dashboard/All-View (nicht in eigener Mine-Ansicht)
  const needsApproval = !editMode && !isMineView
    && /pr[üu]fung/i.test(statusVal) && /einkauf/i.test(statusVal) && !/strategisch/i.test(statusVal);
  const approvalActionBanner = needsApproval ? `
    <div class="approval-action-banner" id="approval-action-banner">
      <div class="aab-title">⏳ Genehmigung ausstehend – Einkauf</div>
      <div class="aab-hint">Bitte prüfen Sie die Anfrage und erteilen Sie Ihre Entscheidung:</div>
      <div class="aab-buttons">
        <button class="btn btn-success btn-sm" id="btn-approve">✅ Genehmigen</button>
        <button class="btn btn-danger  btn-sm" id="btn-reject">❌ Ablehnen</button>
      </div>
      <div class="aab-reject-form" id="aab-reject-form" style="display:none">
        <textarea id="aab-reject-comment" class="pf-comment-textarea" rows="2"
          placeholder="Ablehnungsgrund (optional)…"></textarea>
        <div class="aab-buttons" style="margin-top:6px">
          <button class="btn btn-danger btn-sm" id="btn-reject-confirm">❌ Ablehnung bestätigen</button>
          <button class="btn btn-ghost  btn-sm" id="btn-reject-cancel">Abbrechen</button>
        </div>
      </div>
    </div>` : '';

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
      ${lockBadge}
      <div class="panel-actions">${buttons}</div>
    </div>

    ${approvalActionBanner}
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
        ${fRow(FORM_FIELDS.find(f=>f.key==='Artikelnummer'),         'text')}
        ${fRow(FORM_FIELDS.find(f=>f.key==='ExterneArtikelnummer'), 'text')}
        ${fRow(FORM_FIELDS.find(f=>f.key==='Lieferant'),        'text')}
        ${fRow(FORM_FIELDS.find(f=>f.key==='GeschaetzterPreis'),'number')}
        ${fRow(FORM_FIELDS.find(f=>f.key==='Kostenstelle'),     'text')}
        ${gesamtHint ? `<div class="info-box info" style="margin-top:8px;font-size:.78rem">${gesamtHint}</div>` : ''}
      </div>
      ${(() => {
        const posRaw = gv('Positionen');
        if (!posRaw) return '';
        let positions;
        try { positions = JSON.parse(posRaw); } catch { return ''; }
        if (!Array.isArray(positions) || positions.length === 0) return '';

        _panelPosData = positions.map(p => ({...p}));

        const canEditPos = !editMode && !isMineView && /pr[üu]fung/i.test(statusVal) && /einkauf/i.test(statusVal) && !/strategisch/i.test(statusVal);
        const meOpts = ['Lagereinheiten','kg','Stück','Anzahl','m','Paar','Liter'];

        const rows = positions.map((p, idx) => {
          if (canEditPos) {
            const meSelOpts = meOpts.map(o => `<option value="${o}"${(p.ME||p.me||'')=== o?' selected':''}>${o}</option>`).join('');
            return `<tr class="pos-edit-row" data-pos-idx="${idx}">
              <td class="pos-td pos-nr">${p.Nr ?? idx+1}</td>
              <td class="pos-td"><input class="pos-edit-input" data-idx="${idx}" data-field="Artikelnummer" value="${esc(String(p.Artikelnummer||''))}" placeholder="ArtNr."/></td>
              <td class="pos-td"><input class="pos-edit-input" data-idx="${idx}" data-field="ExterneArtikelnummer" value="${esc(String(p.ExterneArtikelnummer||''))}" placeholder="Ext.ArtNr."/></td>
              <td class="pos-td pos-bez"><input class="pos-edit-input" data-idx="${idx}" data-field="Bezeichnung" value="${esc(String(p.Bezeichnung||''))}" placeholder="Bezeichnung"/></td>
              <td class="pos-td pos-right"><input class="pos-edit-input pos-edit-num" type="number" data-idx="${idx}" data-field="Menge" value="${esc(String(p.Menge||''))}" placeholder="1" min="0.001" step="any"/></td>
              <td class="pos-td"><select class="pos-edit-input pos-edit-me" data-idx="${idx}" data-field="ME"><option value="">–</option>${meSelOpts}</select></td>
              <td class="pos-td pos-right"><input class="pos-edit-input pos-edit-num" type="number" data-idx="${idx}" data-field="Preis" value="${esc(String(p.Preis||''))}" placeholder="0,00" min="0" step="0.01"/></td>
              <td class="pos-td"><input class="pos-edit-input" type="date" data-idx="${idx}" data-field="Termin" value="${esc(String(p.Termin||''))}"/></td>
              <td class="pos-td"><input class="pos-edit-input" data-idx="${idx}" data-field="Kostenstelle" value="${esc(String(p.Kostenstelle||''))}" placeholder="KST"/></td>
              <td class="pos-td"><button type="button" class="pos-del-btn" onclick="panelPosDelete(${idx})" title="Entfernen">✕</button></td>
            </tr>`;
          } else {
            return `<tr>
              <td class="pos-td pos-nr">${p.Nr ?? ''}</td>
              <td class="pos-td">${esc(String(p.Artikelnummer||'–'))}</td>
              <td class="pos-td">${esc(String(p.ExterneArtikelnummer||'–'))}</td>
              <td class="pos-td pos-bez">${esc(String(p.Bezeichnung||'–'))}</td>
              <td class="pos-td pos-right">${esc(String(p.Menge||'–'))}</td>
              <td class="pos-td">${esc(String(p.ME||p.me||'–'))}</td>
              <td class="pos-td pos-right">${p.Preis ? fmtEuro(p.Preis) : '–'}</td>
              <td class="pos-td">${p.Termin ? fmtDate(p.Termin) : '–'}</td>
              <td class="pos-td">${esc(String(p.Kostenstelle||'–'))}</td>
            </tr>`;
          }
        }).join('');

        const addBtn = canEditPos ? `<button type="button" class="btn btn-ghost btn-sm" style="margin-top:6px" onclick="panelPosAdd()">+ Position hinzufügen</button>` : '';
        const saveBtn = canEditPos ? `<div style="margin-top:8px;display:flex;gap:8px"><button class="btn btn-primary btn-sm" onclick="panelPosSave('${item.id}')">💾 Positionen speichern</button></div>` : '';

        return `
        <div class="pf-section" id="pf-pos-section">
          <div class="pf-sec-title">Positionen (${positions.length})</div>
          <div class="pos-table-wrap">
            <table class="pos-table" id="pos-edit-table">
              <thead><tr>
                <th class="pos-th pos-nr">#</th>
                <th class="pos-th">Artikel-Nr.</th>
                <th class="pos-th">Ext. Artikel-Nr.</th>
                <th class="pos-th pos-bez">Bezeichnung</th>
                <th class="pos-th pos-right">Menge</th>
                <th class="pos-th">ME</th>
                <th class="pos-th pos-right">Preis (€)</th>
                <th class="pos-th">Benötigt bis</th>
                <th class="pos-th">Kostenstelle</th>
                ${canEditPos ? '<th class="pos-th"></th>' : ''}
              </tr></thead>
              <tbody id="pos-edit-tbody">${rows}</tbody>
            </table>
          </div>
          ${addBtn}
          ${saveBtn}
        </div>`;
      })()}
      <div class="pf-section">
        <div class="pf-sec-title">Status &amp; Genehmigung</div>
        <div id="panel-approval-body">${approvalInner}</div>
      </div>
      <div class="pf-section">
        <div class="pf-sec-title">Bestellung (Einkauf)</div>
        ${orderNr  ? `<div class="pf-row"><span class="pf-label">Bestellnummer</span><span class="pf-val">${esc(orderNr)}</span></div>` : ''}
        ${lieferd  ? `<div class="pf-row"><span class="pf-label">Lieferdatum</span><span class="pf-val">${fmtDate(lieferd)}</span></div>` : ''}
        ${tatPreis ? `<div class="pf-row"><span class="pf-label">Tatsächl. Preis</span><span class="pf-val">${fmtEuro(tatPreis)}</span></div>` : ''}
        ${!orderNr && !lieferd && !tatPreis ? '<p class="no-order">Noch keine Bestelldaten.</p>' : ''}
      </div>
      <div class="pf-section" id="panel-comments-section">
        <div class="pf-sec-title">Kommentare</div>
        <div id="panel-comments-body"><div class="vh-loading">Lädt…</div></div>
        <div class="pf-comment-compose">
          <textarea id="panel-comment-input" class="pf-comment-textarea" rows="2" placeholder="Kommentar schreiben…"></textarea>
          <button class="btn btn-sm btn-primary" id="panel-comment-send">💬 Senden</button>
        </div>
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

function getKommentarCol() {
  // Try resolved name first, then fallbacks
  const col = resolvedFields[KOMMENTAR_FIELD.key];
  if (col) return col;
  for (const t of [KOMMENTAR_FIELD.key, ...KOMMENTAR_FIELD.alsoTry]) {
    if (colByKey[t]) return t;
  }
  return KOMMENTAR_FIELD.key; // fallback — will fail gracefully
}

function parseKommentare(raw) {
  if (!raw) return [];
  return String(raw).split('\n---\n').map(s => s.trim()).filter(Boolean).map(entry => {
    const m = entry.match(/^\[(.+?)\]:\s*([\s\S]*)$/);
    if (m) return { header: m[1], text: m[2].trim() };
    return { header: '', text: entry };
  });
}

async function postSpComment(itemId, text) {
  if (!text?.trim()) return false;
  const col = getKommentarCol();
  const item = allItems.find(i => String(i.id) === String(itemId));
  const existing = String(getField(item, col) || '').trim();
  const author = account?.name || account?.username || 'Unbekannt';
  const now = new Date();
  const dt = now.toLocaleString('de-DE', { day:'2-digit', month:'2-digit', year:'numeric', hour:'2-digit', minute:'2-digit' });
  const newEntry = `[${dt} – ${author}]: ${text.trim()}`;
  const combined = existing ? existing + '\n---\n' + newEntry : newEntry;
  try {
    await gPatch(`/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, { [col]: combined });
    // Update local cache so re-render is instant
    if (item?.fields) item.fields[col] = combined;
    return true;
  } catch(e) {
    console.warn('[comments] Speichern fehlgeschlagen:', e.message);
    return false;
  }
}

// ── LOCK HELPERS ──────────────────────────────────────────────────────────────
function getLockField() {
  if (_lockField !== undefined) return _lockField;
  for (const n of LOCK_FIELD_CANDIDATES) { if (colByKey[n]) { _lockField = n; return n; } }
  _lockField = ''; return '';
}
function parseLock(raw) {
  if (!raw) return null;
  const i = String(raw).indexOf('|');
  if (i < 0) return null;
  const expiresAt = new Date(String(raw).slice(i + 1));
  return isNaN(expiresAt.getTime()) ? null : { by: String(raw).slice(0, i), expiresAt };
}
async function tryAcquireLock(itemId) {
  const field = getLockField();
  const item  = allItems.find(i => String(i.id) === String(itemId));
  const me    = account?.name || account?.username || 'Unbekannt';
  if (field && item) {
    const lock = parseLock(getField(item, field));
    if (lock && lock.expiresAt > new Date() && lock.by !== me) {
      const mins = Math.ceil((lock.expiresAt - new Date()) / 60000);
      toast(`Wird von „${lock.by}" bearbeitet – noch ${mins} Min. gesperrt.`, 'error');
      return false;
    }
    const until = new Date(Date.now() + 15 * 60 * 1000).toISOString();
    try {
      await gPatch(`/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, { [field]: `${me}|${until}` });
      if (item.fields) item.fields[field] = `${me}|${until}`;
    } catch(e) { console.warn('[lock] acquire failed (field may not exist in SP):', e.message); }
  }
  if (lockTimers[itemId]) clearTimeout(lockTimers[itemId]);
  lockTimers[String(itemId)] = setTimeout(() => autoReleaseLock(itemId), 15 * 60 * 1000);
  return true;
}
async function releaseLock(itemId) {
  const key = String(itemId);
  if (lockTimers[key]) { clearTimeout(lockTimers[key]); delete lockTimers[key]; }
  const field = getLockField();
  if (!field) return;
  const item = allItems.find(i => String(i.id) === key);
  const me   = account?.name || account?.username || 'Unbekannt';
  const lock = parseLock(getField(item, field));
  if (!lock || lock.by !== me) return; // never release someone else's lock
  try {
    await gPatch(`/sites/${siteId}/lists/${listId}/items/${key}/fields`, { [field]: '' });
    if (item?.fields) item.fields[field] = '';
  } catch(e) { console.warn('[lock] release failed:', e.message); }
}
async function autoReleaseLock(itemId) {
  const key = String(itemId);
  delete lockTimers[key];
  await releaseLock(key);
  if (String(panelItemId) === key) {
    const item      = allItems.find(i => String(i.id) === key);
    const panelRoot = $id(`panel-${currentView}-content`);
    if (item && panelRoot) {
      toast('Bearbeitungssperre abgelaufen (15 Min.) – Bearbeitungsmodus beendet.', 'info');
      panelRoot.innerHTML = renderPanel(item);
      bindPanelEvents(key);
    }
  }
}
async function startEditMode(itemId) {
  const ok = await tryAcquireLock(itemId);
  if (!ok) return;
  const item      = allItems.find(i => String(i.id) === String(itemId));
  const panelRoot = $id(`panel-${currentView}-content`);
  if (!item || !panelRoot) return;
  panelRoot.innerHTML = renderPanel(item, true);
  bindPanelEvents(itemId);
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
  const email    = (account?.username || '').toLowerCase();
  const s        = getSettings(email);
  const adminMode = email === ADMIN_EMAIL;

  // ── Admin: per-user row ───────────────────────────────────────────────────
  const userRow = (em, us) => `
    <div class="su-row" id="su-row-${btoa(em).replace(/=/g,'')}">
      <span class="su-email" title="${esc(em)}">${esc(em)}</span>
      <label class="tgl-wrap" title="Auto-Refresh">
        <input type="checkbox" ${us.autoRefresh && us.autoRefreshGranted ? 'checked' : ''}
          onchange="saveUserSettings('${esc(em)}',{autoRefresh:this.checked,autoRefreshGranted:this.checked})">
        <span class="tgl"></span>
      </label>
      <label class="tgl-wrap" title="Dashboard-Zugriff (alle Anfragen)">
        <input type="checkbox" ${us.canSeeDashboard ? 'checked' : ''}
          onchange="saveUserSettings('${esc(em)}',{canSeeDashboard:this.checked})">
        <span class="tgl"></span>
      </label>
      <input type="number" value="${us.pageSize||100}" min="10" max="500" step="10"
        class="su-num" title="Max. Elemente"
        onchange="saveUserSettings('${esc(em)}',{pageSize:parseInt(this.value)||100})">
      <button class="btn btn-sm btn-outline su-token-btn" title="Grant-Token kopieren und an Benutzer senden"
        onclick="copyGrantToken('${esc(em)}')">📋 Token</button>
      <button class="su-del" title="Benutzer entfernen"
        onclick="deleteUserSetting('${esc(em)}')">✕</button>
    </div>`;

  const adminSection = adminMode ? `
    <hr class="modal-hr">
    <div class="settings-section-title">👥 Benutzer verwalten</div>
    <div class="su-add">
      <input type="email" id="su-new-email" placeholder="user@dihag.com" class="su-input">
      <button class="btn btn-sm btn-primary" onclick="addUserSetting()">+ Hinzufügen</button>
    </div>
    <div class="su-header-row">
      <span class="su-email" style="font-size:.7rem;color:#6b7280">E-Mail</span>
      <span class="su-col-lbl" title="Auto-Refresh (30s)">⏱</span>
      <span class="su-col-lbl" title="Dashboard-Zugriff">📊</span>
      <span class="su-col-lbl" title="Max. Elemente">Elem.</span>
      <span style="width:22px"></span>
    </div>
    <div id="su-list">${
      Object.entries(getAllUserSettings())
        .filter(([k]) => k !== email)
        .map(([k, v]) => userRow(k, v)).join('') ||
      '<p class="su-empty">Noch keine weiteren Benutzer konfiguriert.</p>'
    }</div>` : '';

  // ── "Meine Einstellungen" sections ────────────────────────────────────────
  const roVal = (on, label) => `<span class="settings-val-ro">${on ? `✅ ${label}` : '—'}</span>`;

  // Auto-Refresh: admin can toggle for self; others see read-only
  const arRow = adminMode
    ? `<div class="settings-row">
        <span class="settings-label">⏱ Auto-Aktualisierung (alle 30s)</span>
        <label class="tgl-wrap">
          <input type="checkbox" ${s.autoRefresh && s.autoRefreshGranted ? 'checked' : ''}
            onchange="saveUserSettings('${esc(email)}',{autoRefresh:this.checked,autoRefreshGranted:this.checked});this.checked?startAutoRefresh():stopAutoRefresh()">
          <span class="tgl"></span>
        </label>
       </div>`
    : `<div class="settings-row"><span class="settings-label">⏱ Auto-Aktualisierung</span>
        ${roVal(s.autoRefresh && s.autoRefreshGranted, 'Aktiviert')}</div>`;

  // canSeeDashboard: admin can toggle for self; others see read-only
  const allRow = adminMode
    ? `<div class="settings-row">
        <span class="settings-label">📊 Dashboard-Zugriff</span>
        <label class="tgl-wrap">
          <input type="checkbox" ${s.canSeeDashboard ? 'checked' : ''}
            onchange="saveUserSettings('${esc(email)}',{canSeeDashboard:this.checked});applyNavVisibility()">
          <span class="tgl"></span>
        </label>
       </div>`
    : `<div class="settings-row"><span class="settings-label">📊 Dashboard-Zugriff</span>
        ${roVal(s.canSeeDashboard, 'Freigegeben')}</div>`;

  // Token-import row for non-admin (admin generates token, sends it, user pastes here)
  const syncBtn = !adminMode ? `
    <div class="settings-row" style="margin-top:6px;gap:6px;flex-wrap:wrap;align-items:center">
      <span class="settings-label" style="font-size:.78rem;color:#6b7280">Grant-Token vom Admin:</span>
      <input id="grant-token-input" type="text" placeholder="Token hier einfügen…"
        class="su-input" style="flex:1;min-width:160px;font-size:.78rem">
      <button class="btn btn-sm btn-primary" onclick="applyGrantToken()">✅ Anwenden</button>
    </div>` : '';

  $id('settings-body').innerHTML = `
    <div class="settings-section-title">⚙️ Meine Einstellungen</div>
    ${arRow}
    ${allRow}
    ${syncBtn}
    <div class="settings-row">
      <span class="settings-label">📄 Elemente laden (max.)</span>
      <input type="number" value="${s.pageSize}" min="10" max="500" step="10" class="su-num"
        onchange="saveUserSettings('${esc(email)}',{pageSize:parseInt(this.value)||100})">
    </div>
    <hr class="modal-hr">
    <div class="settings-section-title">🗂 Ansicht (Meine Anfragen)</div>
    <div class="settings-row">
      <span class="settings-label">Kompaktansicht</span>
      <label class="tgl-wrap">
        <input type="checkbox" ${s.compactView ? 'checked' : ''}
          onchange="saveUserSettings('${esc(email)}',{compactView:this.checked});filterView('mine')">
        <span class="tgl"></span>
      </label>
    </div>
    <div class="settings-row">
      <span class="settings-label">Abgeschlossene ausblenden</span>
      <label class="tgl-wrap">
        <input type="checkbox" ${s.hideCompleted ? 'checked' : ''}
          onchange="saveUserSettings('${esc(email)}',{hideCompleted:this.checked});filterView('mine')">
        <span class="tgl"></span>
      </label>
    </div>
    <div class="settings-row">
      <span class="settings-label">Standard-Sortierung</span>
      <select class="su-select"
        onchange="saveUserSettings('${esc(email)}',{defaultSort:this.value});mineSortOrder=this.value;filterView('mine')">
        <option value="date-desc"  ${s.defaultSort==='date-desc'  ? 'selected':''}>Neueste zuerst</option>
        <option value="date-asc"   ${s.defaultSort==='date-asc'   ? 'selected':''}>Älteste zuerst</option>
        <option value="status"     ${s.defaultSort==='status'     ? 'selected':''}>Nach Status</option>
        <option value="price-desc" ${s.defaultSort==='price-desc' ? 'selected':''}>Preis absteigend</option>
        <option value="price-asc"  ${s.defaultSort==='price-asc'  ? 'selected':''}>Preis aufsteigend</option>
      </select>
    </div>
    ${adminSection}`;

  $id('settings-modal').classList.remove('hidden');
}

function closeSettings() { $id('settings-modal').classList.add('hidden'); }

// Admin: generate a base64 token for a specific user and copy to clipboard.
function copyGrantToken(targetEmail) {
  const settings = getAllUserSettings();
  const userCfg  = settings[(targetEmail||'').toLowerCase()];
  if (!userCfg) { toast('Benutzer nicht gefunden.', 'error'); return; }
  // Token encodes: {email, grants} — only grant fields, not personal preferences
  const grantFields = ['canSeeDashboard', 'autoRefresh', 'autoRefreshGranted'];
  const grants = {};
  grantFields.forEach(k => { if (userCfg[k] !== undefined) grants[k] = userCfg[k]; });
  const token = btoa(unescape(encodeURIComponent(JSON.stringify({ email: targetEmail.toLowerCase(), grants }))));
  navigator.clipboard.writeText(token).then(
    () => toast(`Token für ${targetEmail} kopiert. Bitte an den Benutzer senden.`, 'success'),
    () => {
      // Fallback: show in prompt
      prompt('Token kopieren:', token);
    }
  );
}

// User: paste a token from admin and apply the contained grants.
function applyGrantToken() {
  const raw = ($id('grant-token-input')?.value || '').trim();
  if (!raw) { toast('Bitte Token eingeben.', 'error'); return; }
  try {
    const decoded = JSON.parse(decodeURIComponent(escape(atob(raw))));
    if (!decoded?.email || !decoded?.grants) throw new Error('Ungültiges Format');
    const myEmail = (account?.username || '').toLowerCase();
    if (decoded.email !== myEmail) {
      toast(`Dieser Token ist für ${decoded.email}, nicht für dein Konto.`, 'error'); return;
    }
    saveUserSettings(myEmail, decoded.grants);
    applyNavVisibility();
    toast('Berechtigungen erfolgreich übernommen!', 'success');
    if (decoded.grants.canSeeDashboard) {
      setTimeout(() => location.reload(), 900);
    } else {
      openSettings();
    }
  } catch(e) {
    toast('Ungültiger Token — bitte erneut beim Admin anfragen.', 'error');
    console.warn('[applyGrantToken]', e.message);
  }
}

function addUserSetting() {
  const em = ($id('su-new-email')?.value || '').trim().toLowerCase();
  if (!em || !em.includes('@') || !em.includes('.')) {
    toast('Bitte gültige E-Mail-Adresse eingeben.', 'error'); return;
  }
  const all = JSON.parse(localStorage.getItem(SETTINGS_KEY) || '{}');
  if (all[em]) { toast(`${em} ist bereits vorhanden.`, 'info'); openSettings(); return; }
  saveUserSettings(em, { pageSize: 100 });
  toast(`${em} hinzugefügt.`, 'success');
  openSettings();
}

function deleteUserSetting(em) {
  const all = JSON.parse(localStorage.getItem(SETTINGS_KEY) || '{}');
  delete all[(em||'').toLowerCase()];
  localStorage.setItem(SETTINGS_KEY, JSON.stringify(all));
  persistSpSettings().catch(() => {});
  toast(`${em} entfernt.`, 'info');
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

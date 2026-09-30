/*  NABAVA — Power Query (M) code
    ================================================================
    Paste each block into its own query via
    Podaci -> Dohvati podatke -> Iz drugih izvora -> Prazan upit
    -> Prikazi -> Napredni uredjivac  (Data -> Get Data -> Blank Query
       -> View -> Advanced Editor in an English Excel)

    Create them IN THIS ORDER. Names must match exactly.
    Queries 1-6 = "Samo stvori vezu" (Only Create Connection).
    Query 7 (qNabava) loads onto the NABAVA sheet.

    LANGUAGE INDEPENDENCE
    ---------------------
    M itself is identical in every language version of Excel, so none of
    this changes on a Croatian vs English install. Three extra guards:
      * column names are matched after stripping diacritics and case,
        so "Dobavljač" / "DOBAVLJAC" / "dobavljac" all resolve;
      * every date and number conversion passes an explicit culture, so
        a machine set to en-US parses "03.09.2026" the same as hr-HR;
      * sheets are taken by position, never by name, because the export
        sheet name carries the export date.
*/


/* ════════════════════════════════════════════════════════════════
   1.  fnNorm     — helper: strip diacritics + lowercase
   ════════════════════════════════════════════════════════════════ */
/*@ query: fnNorm | load: connection */
let
    fnNorm = (tekst as nullable text) as text =>
        let
            t = Text.Lower(Text.Trim(tekst ?? "")),
            parovi = {
                {"č","c"},{"ć","c"},{"ž","z"},{"š","s"},{"đ","d"},
                {"Č","c"},{"Ć","c"},{"Ž","z"},{"Š","s"},{"Đ","d"}
            },
            zamijeni = List.Accumulate(parovi, t,
                (stanje, par) => Text.Replace(stanje, par{0}, par{1}))
        in
            zamijeni
in
    fnNorm


/* ════════════════════════════════════════════════════════════════
   2.  fnStupac   — helper: fetch a column by normalised name
                    Returns null if the column is absent, so a renamed
                    or missing column degrades instead of erroring.
   ════════════════════════════════════════════════════════════════ */
/*@ query: fnStupac | load: connection */
let
    fnStupac = (tablica as table, trazeni as text) as nullable text =>
        let
            cilj = fnNorm(trazeni),
            svi = Table.ColumnNames(tablica),
            pogodak = List.Select(svi, each fnNorm(_) = cilj),
            rezultat = if List.Count(pogodak) > 0 then pogodak{0} else null
        in
            rezultat
in
    fnStupac


/* ════════════════════════════════════════════════════════════════
   3.  Putanja    — the folder this workbook lives in
                    Reads POSTAVKE!B4, which computes itself with
                    =LEFT(CELL("filename"),FIND("[",CELL("filename"))-1)
                    Move the folder, rename it, put it on OneDrive:
                    nothing needs editing.
   ════════════════════════════════════════════════════════════════ */
/*@ query: Putanja | load: connection */
let
    celija = Excel.CurrentWorkbook(){[Name="MapaPutanja"]}[Content],
    vrijednost = Text.From(Table.Column(celija, Table.ColumnNames(celija){0}){0}),
    // guard: workbook never saved -> CELL("filename") is empty
    provjera = if vrijednost = null or Text.Length(vrijednost) < 3
               then error "Spremite datoteku (Save) prije osvjezavanja."
               else vrijednost
in
    provjera


/* ════════════════════════════════════════════════════════════════
   4.  Postavke   — Prag / Tranzit / RollingN as numbers
   ════════════════════════════════════════════════════════════════ */
/*@ query: Postavke | load: connection */
let
    uzmi = (ime as text) as number =>
        let
            t = Excel.CurrentWorkbook(){[Name=ime]}[Content],
            v = Table.Column(t, Table.ColumnNames(t){0}){0}
        in
            Number.From(v, "hr-HR"),
    zapis = [
        Prag     = uzmi("Prag"),
        Tranzit  = uzmi("Tranzit"),
        RollingN = uzmi("RollingN")
    ]
in
    zapis


/* ════════════════════════════════════════════════════════════════
   5.  fnDatoteke — every .xlsx or .xls in the folder whose normalised
                    name starts with a given prefix, newest first.
                    The date comes from the filename (DD_MM_YYYY),
                    with the file timestamp as fallback. One file per
                    date: if an .xls was also re-saved as .xlsx, the
                    .xlsx wins.
   ════════════════════════════════════════════════════════════════ */
/*@ query: fnDatoteke | load: connection */
let
    fnDatoteke = (prefiks as text) as table =>
        let
            mapa = Folder.Files(Putanja),
            // .xls too: the ERP export can go in as it comes, no Save As
            samoExcel = Table.SelectRows(mapa, each
                (Text.EndsWith(fnNorm([Name]), ".xlsx")
                    or Text.EndsWith(fnNorm([Name]), ".xls"))
                and not Text.StartsWith(fnNorm([Name]), "~$")
                and not Text.StartsWith(fnNorm([Name]), "nabava")),
            odabir = Table.SelectRows(samoExcel, each
                Text.StartsWith(fnNorm([Name]), fnNorm(prefiks))),
            sDatumom = Table.AddColumn(odabir, "DatumIzvoza", each
                let
                    m = Text.Select([Name], {"0".."9","_","."}),
                    dijelovi = List.Select(
                        Text.Split(Text.Replace(m, ".", "_"), "_"),
                        each Text.Length(_) > 0),
                    n = List.Count(dijelovi),
                    kandidat =
                        if n >= 3 then
                            try #date(
                                Number.FromText(dijelovi{n-1}),
                                Number.FromText(dijelovi{n-2}),
                                Number.FromText(dijelovi{n-3})
                            ) otherwise null
                        else null
                in
                    if kandidat = null then Date.From([Date modified]) else kandidat,
                type date),
            // the same export as .xls and re-saved .xlsx would parse to one
            // date, and qStanjePrethodno would compare the stock with itself
            sXlsx = Table.AddColumn(sDatumom, "JeXlsx", each
                Text.EndsWith(fnNorm([Name]), ".xlsx"), type logical),
            sortirano = Table.Buffer(Table.Sort(sXlsx,
                {{"DatumIzvoza", Order.Descending}, {"JeXlsx", Order.Descending}})),
            poDatumu = Table.Distinct(sortirano, {"DatumIzvoza"})
        in
            poDatumu
in
    fnDatoteke


/* ════════════════════════════════════════════════════════════════
   6a. qStanje         — newest Stanje skladista export
   6b. qStanjePrethodno— the one before it (for stock movement)
   ════════════════════════════════════════════════════════════════

   Paste 6a as qStanje. Then paste it AGAIN as qStanjePrethodno and
   change the single marked line from  idx = 0  to  idx = 1.
   (import_queries.ps1 does this for you.)
*/
/*@ query: qStanje | load: connection */
let
    // qStanjePrethodno is this same query with idx = 1. The clamp means that
    // when only ONE Stanje export exists, "previous" falls back to the current
    // one, so movement reads zero instead of erroring on a missing element.
    idx = 0,                                     // <<< qStanjePrethodno: idx = 1
    // A real .xlsx gives numbers as numbers. The legacy reader (a renamed .xls)
    // gives them as text with a decimal POINT, "82.5", which hr-HR would read
    // as 825. So: a comma means Croatian format, otherwise a point is decimal.
    uBroj = (v) as number =>
        if v = null then 0
        else if v is number then v
        else let t = Text.Trim(Text.From(v)) in
            if t = "" then 0
            else if Text.Contains(t, ",") then Number.From(t, "hr-HR")
            else Number.From(t, "en-US"),
    // the export ends with a totals row: a count in Artikal and no name
    imaNaziv = (r as record) as logical =>
        Text.Trim(Text.From(Record.FieldOrDefault(r, "NazivArtikla", "?") ?? "")) <> "",
    datoteke = fnDatoteke("stanje"),
    broj     = Table.RowCount(datoteke),
    odabrana = if broj = 0
               then error "U mapi nema datoteke 'Stanje skladista...'"
               else datoteke{List.Min({idx, broj - 1})},
    sadrzaj  = odabrana[Content],
    knjiga   = Excel.Workbook(sadrzaj, null, true),
    // A file that is .xlsx by name only (an HTML/.xls export renamed) goes
    // through the legacy reader, which returns no Kind column. Take its
    // tables as they are, and name the file if even that fails.
    listovi  = if Table.HasColumns(knjiga, "Kind")
               then Table.SelectRows(knjiga, each [Kind] = "Sheet")
               else if Table.HasColumns(knjiga, "Data") then knjiga
               else error ("Datoteka '" & odabrana[Name] & "' nije ispravna "
                    & "Excel datoteka. Otvorite je u Excelu i spremite kao .xlsx."),
    prvi     = listovi{0}[Data],                 // by position, not by name
    // the legacy reader may already have promoted the header row
    zaglavlje= if fnStupac(prvi, "Artikal") <> null then prvi
               else Table.PromoteHeaders(prvi, [PromoteAllScalars=true]),

    cArt = fnStupac(zaglavlje, "Artikal"),
    cNaz = fnStupac(zaglavlje, "Naziv artikla"),
    cZal = fnStupac(zaglavlje, "Zaliha"),
    cDob = fnStupac(zaglavlje, "Dobavljac"),
    cDobN= fnStupac(zaglavlje, "Naziv dobavljaca"),

    izabrano = Table.SelectColumns(zaglavlje,
        List.RemoveNulls({cArt, cNaz, cZal, cDob, cDobN}), MissingField.Ignore),
    preimenovano = Table.RenameColumns(izabrano,
        List.RemoveNulls({
            if cArt  <> null then {cArt , "Artikal"}          else null,
            if cNaz  <> null then {cNaz , "NazivArtikla"}     else null,
            if cZal  <> null then {cZal , "Zaliha"}           else null,
            if cDob  <> null then {cDob , "SifraDobavljaca"}  else null,
            if cDobN <> null then {cDobN, "NazivDobavljaca"}  else null
        })),

    // Artikal MUST stay text. 00111930 as a number becomes 111930 and the
    // join silently drops those rows.
    tipovi = Table.TransformColumnTypes(preimenovano, {
        {"Artikal", type text}, {"NazivArtikla", type text},
        {"SifraDobavljaca", type text}, {"NazivDobavljaca", type text}
    }, "hr-HR"),
    zalihaBroj = Table.TransformColumns(tipovi,
        {{"Zaliha", uBroj, type number}}),
    ocisceno = Table.TransformColumns(zalihaBroj,
        {{"Artikal", each Text.Trim(Text.From(_)), type text}}),
    bezPraznih = Table.SelectRows(ocisceno, each
        [Artikal] <> null and [Artikal] <> "" and imaNaziv(_))
in
    bezPraznih


/*@ query: qStanjePrethodno | load: connection | derive-from: qStanje | replace-once: idx = 0 => idx = 1 */

/* ════════════════════════════════════════════════════════════════
   6c. qProdaja  — rolling average across ALL Analiza prodaje files
   ════════════════════════════════════════════════════════════════

   This is the year-rollover answer. Every Analiza prodaje file in the
   folder is read, each month becomes a row tagged with its year, and the
   average is taken over the last RollingN COMPLETED months. In January
   2027 the window is Feb-Dec 2026, so the average never divides by one.
   Keeping the 2026 file in the folder is the only manual step.
*/
/*@ query: qProdaja | load: connection */
let
    mjeseci = {"Sijecanj","Veljaca","Ozujak","Travanj","Svibanj","Lipanj",
               "Srpanj","Kolovoz","Rujan","Listopad","Studeni","Prosinac"},
    // same helpers as qStanje. A real .xlsx gives numbers as numbers. The legacy
    // reader (a renamed .xls)
    // gives them as text with a decimal POINT, "82.5", which hr-HR would read
    // as 825. So: a comma means Croatian format, otherwise a point is decimal.
    uBroj = (v) as number =>
        if v = null then 0
        else if v is number then v
        else let t = Text.Trim(Text.From(v)) in
            if t = "" then 0
            else if Text.Contains(t, ",") then Number.From(t, "hr-HR")
            else Number.From(t, "en-US"),
    // the export ends with a totals row: a count in Artikal and no name
    imaNaziv = (r as record) as logical =>
        Text.Trim(Text.From(Record.FieldOrDefault(r, "NazivArtikla", "?") ?? "")) <> "",

    datoteke = fnDatoteke("analiza"),
    provjera = if Table.RowCount(datoteke) = 0
               then error "U mapi nema datoteke 'Analiza prodaje...'"
               else datoteke,

    // the newest export defines "now"; months before it are completed
    najnoviji  = provjera{0}[DatumIzvoza],
    tekuciIdx  = Date.Year(najnoviji) * 12 + Date.Month(najnoviji),

    ucitaj = Table.AddColumn(provjera, "Tab", each
        let
            ime    = [Name],
            knjiga = Excel.Workbook([Content], null, true),
            // same guard as qStanje: a renamed .xls/HTML export has no Kind
            listovi= if Table.HasColumns(knjiga, "Kind")
                     then Table.SelectRows(knjiga, each [Kind] = "Sheet")
                     else if Table.HasColumns(knjiga, "Data") then knjiga
                     else error ("Datoteka '" & ime & "' nije ispravna "
                          & "Excel datoteka. Otvorite je u Excelu i spremite kao .xlsx."),
            prvi   = listovi{0}[Data],
            zag    = if fnStupac(prvi, "Artikal") <> null then prvi
                     else Table.PromoteHeaders(prvi, [PromoteAllScalars=true]),
            cArt   = fnStupac(zag, "Artikal"),
            // two columns are called Naziv: the group's, then the article's.
            // Take the first Naziv AFTER Artikal.
            sviZag = Table.ColumnNames(zag),
            nakonArt = if cArt = null then sviZag
                       else List.Skip(sviZag, List.PositionOf(sviZag, cArt) + 1),
            cNaz   = List.First(List.Select(nakonArt,
                        each Text.StartsWith(fnNorm(_), "naziv")), fnStupac(zag, "Naziv")),
            imena  = List.RemoveNulls(
                        List.Transform(mjeseci, each fnStupac(zag, _))),
            uzmi   = Table.SelectColumns(zag,
                        List.RemoveNulls({cArt, cNaz}) & imena, MissingField.Ignore),
            preim  = Table.RenameColumns(uzmi,
                        List.RemoveNulls({
                            if cArt <> null then {cArt, "Artikal"} else null,
                            if cNaz <> null then {cNaz, "NazivArtikla"} else null
                        })),
            tekst  = Table.TransformColumnTypes(preim, {{"Artikal", type text}}, "hr-HR"),
            cisto  = Table.SelectRows(tekst, each
                        [Artikal] <> null and [Artikal] <> "" and imaNaziv(_))
        in
            cisto),

    // one file per export date; keep only the newest file per year
    poGodini = Table.Group(
        Table.AddColumn(ucitaj, "Godina", each Date.Year([DatumIzvoza]), Int64.Type),
        {"Godina"}, {{"Naj", each Table.FirstN(Table.Sort(_,
            {{"DatumIzvoza", Order.Descending}}), 1), type table}}),
    razvuci = Table.Combine(Table.TransformRows(poGodini, (g) =>
        Table.AddColumn(g[Naj]{0}[Tab], "Godina", each g[Godina], Int64.Type))),

    unpivot = Table.UnpivotOtherColumns(razvuci,
        {"Artikal","NazivArtikla","Godina"}, "Mjesec", "Kolicina"),
    brojcano = Table.TransformColumns(unpivot,
        {{"Kolicina", uBroj, type number}}),
    indeks = Table.AddColumn(brojcano, "Idx", each
        [Godina] * 12 + List.PositionOf(
            List.Transform(mjeseci, fnNorm), fnNorm([Mjesec])) + 1, Int64.Type),

    // completed months only: the current, partial month is excluded because
    // a 3-day month would drag every average down
    uProzoru = (i as number) as logical =>
        i < tekuciIdx and i >= tekuciIdx - Postavke[RollingN],
    godina   = Date.Year(najnoviji),

    // every article sold in the window OR this year, so the month block on
    // NABAVA also shows an article whose only sales are in the current month
    relevantno = Table.SelectRows(indeks, each uProzoru([Idx]) or [Godina] = godina),
    grupirano = Table.Group(relevantno, {"Artikal"}, {
        {"NazivArtikla", each List.First([NazivArtikla]), type text},
        {"Zbroj", each List.Sum(List.Transform(
            Table.SelectRows(_, each uProzoru([Idx]))[Kolicina], each _ ?? 0)), type number},
        {"BrojMjeseci", each List.Count(List.Distinct(
            Table.SelectRows(_, each uProzoru([Idx]))[Idx])), Int64.Type}
    }),
    prosjek = Table.AddColumn(grupirano, "Prosjek", each
        if [BrojMjeseci] = 0 then 0
        else Number.Round([Zbroj] / [BrojMjeseci], 2), type number),

    // raw monthly sales of the newest export's year, Mj1..Mj12, for the
    // Sij..Pro block to the right of STATUS (the partial month included,
    // as in the original model)
    mjKol    = List.Transform({1..12}, each "Mj" & Text.From(_)),
    oveGod   = Table.SelectRows(indeks, each [Godina] = godina),
    kljuc    = Table.AddColumn(Table.SelectColumns(oveGod, {"Artikal","Idx","Kolicina"}),
                   "Mj", each "Mj" & Text.From([Idx] - godina * 12), type text),
    samoMj   = Table.SelectRows(kljuc, each List.Contains(mjKol, [Mj])),
    pivot    = Table.Pivot(Table.RemoveColumns(samoMj, {"Idx"}),
                   mjKol, "Mj", "Kolicina", List.Sum),
    spojeno  = Table.NestedJoin(prosjek, {"Artikal"}, pivot, {"Artikal"}, "Mj", JoinKind.LeftOuter),
    sMjesecima = Table.ExpandTableColumn(spojeno, "Mj", mjKol)
in
    sMjesecima


/* ════════════════════════════════════════════════════════════════
   6d. qDolasci  — the manual ROBA_U_DOLASKU table, ranked
   ════════════════════════════════════════════════════════════════ */
/*@ query: qDolasci | load: connection */
let
    izvor = Excel.CurrentWorkbook(){[Name="tblDolasci"]}[Content],
    tipovi = Table.TransformColumnTypes(izvor, {
        {"Artikal", type text}, {"Sifra dobavljaca", type text}
    }, "hr-HR"),
    brojevi = Table.TransformColumns(tipovi, {
        {"Kolicina", each Number.From(_, "hr-HR") ?? 0, type number},
        {"Mjeseci do dolaska", each Number.From(_, "hr-HR") ?? 0, type number}
    }),
    cisto = Table.SelectRows(brojevi, each
        [Artikal] <> null and Text.Trim([Artikal]) <> "" and [Kolicina] > 0),
    trim = Table.TransformColumns(cisto, {{"Artikal", Text.Trim, type text}}),

    // chronological rank within each article: 1, 2, 3
    sortirano = Table.Sort(trim, {{"Artikal", Order.Ascending},
                                  {"Mjeseci do dolaska", Order.Ascending}}),
    grupe = Table.Group(sortirano, {"Artikal"}, {{"Redci", each
        Table.AddIndexColumn(_, "Rb", 1, 1, Int64.Type), type table}}),
    spojeno = Table.Combine(Table.Column(grupe, "Redci")),
    samoTri = Table.SelectRows(spojeno, each [Rb] <= 3)
in
    samoTri


/* ════════════════════════════════════════════════════════════════
   7.  qNabava   — the main screen. LOAD THIS ONE TO THE SHEET.
   ════════════════════════════════════════════════════════════════

   Column order is fixed here because import_queries.ps1 binds the colour
   rules to G (Izlaz), I (Zaliha traje) and Q (STATUS), as in
   NABAVA_model_ver03.xlsx. R is an empty spacer, Sij..Pro are S..AD.
   If you reorder columns, move the rules in Set-NabavaLayout too.
*/
/*@ query: qNabava | load: sheet NABAVA!$A$4 */
let
    prag = Postavke[Prag],

    // master list = every article seen in ANY source. The stock export is
    // filtered to warehouse 04 and misses ~139 articles that do have sales.
    // Buffered once. Power Query does not cache query references: every mention
    // of qStanje / qProdaja re-runs its whole chain, which means re-opening the
    // export files, and every mention of qDolasci re-reads this workbook.
    bufStanje    = Table.Buffer(qStanje),
    bufPrethodno = Table.Buffer(qStanjePrethodno),
    bufProdaja   = Table.Buffer(qProdaja),
    bufDolasci   = Table.Buffer(qDolasci),

    sviArtikli = Table.Distinct(Table.Combine({
        Table.SelectColumns(bufStanje,  {"Artikal"}),
        Table.SelectColumns(bufProdaja, {"Artikal"}),
        Table.SelectColumns(bufDolasci, {"Artikal"})
    })),

    sStanjem = Table.NestedJoin(sviArtikli, {"Artikal"}, bufStanje, {"Artikal"}, "S", JoinKind.LeftOuter),
    sPrethodnim = Table.NestedJoin(sStanjem, {"Artikal"}, bufPrethodno, {"Artikal"}, "P", JoinKind.LeftOuter),
    sProdajom = Table.NestedJoin(sPrethodnim, {"Artikal"}, bufProdaja, {"Artikal"}, "R", JoinKind.LeftOuter),
    // shipments joined like every other source instead of re-scanned per row
    sDolascima = Table.NestedJoin(sProdajom, {"Artikal"}, bufDolasci, {"Artikal"}, "D", JoinKind.LeftOuter),

    polja = Table.AddColumn(sDolascima, "X", each
        let
            s = try [S]{0} otherwise null,
            p = try [P]{0} otherwise null,
            r = try [R]{0} otherwise null,
            zaliha = if s = null then 0 else s[Zaliha],
            prethodno = if p = null then null else p[Zaliha],
            prosjek = if r = null then 0 else r[Prosjek],
            naziv = if s <> null and s[NazivArtikla] <> null then s[NazivArtikla]
                    else if r <> null then r[NazivArtikla] else "",
            // [D] already holds only this article's shipments, from the join.
            // Filtering qDolasci here re-evaluated it once per article row, and
            // because it reads Excel.CurrentWorkbook() that cost grew as the
            // output table itself grew.
            dolasci = Table.Sort([D], {{"Rb", Order.Ascending}}),
            d1 = try dolasci{0} otherwise null,
            d2 = try dolasci{1} otherwise null,
            d3 = try dolasci{2} otherwise null,
            m1 = if d1 = null then 0 else d1[#"Mjeseci do dolaska"],
            q1 = if d1 = null then 0 else d1[Kolicina],
            m2 = if d2 = null then 0 else d2[#"Mjeseci do dolaska"],
            q2 = if d2 = null then 0 else d2[Kolicina],
            m3 = if d3 = null then 0 else d3[#"Mjeseci do dolaska"],
            q3 = if d3 = null then 0 else d3[Kolicina],

            traje = if prosjek <= 0 then 999 else Number.Round(zaliha / prosjek, 1),

            // walk the stock forward; a negative level before a shipment is a gap
            n1 = zaliha - prosjek * m1,
            n2 = zaliha - prosjek * m2 + q1,
            n3 = zaliha - prosjek * m3 + q1 + q2,
            ostatak = if q3 > 0 then n3 + q3
                      else if q2 > 0 then n2 + q2
                      else if q1 > 0 then n1 + q1
                      else zaliha,
            pokrivenost = if prosjek <= 0 then 999
                          else Number.Round(ostatak / prosjek, 1),

            status =
                if prosjek <= 0 then "Nema prodaje"
                else if q1 = 0 and q2 = 0 and q3 = 0 then
                    (if traje < prag then "Naruči odmah" else "Sve u redu")
                else if n1 < 0 then "Rupa za ~" & Number.ToText(traje, "F1", "hr-HR") & " mj."
                else if q2 > 0 and n2 < 0 then "Rupa nakon 1. dolaska"
                else if q3 > 0 and n3 < 0 then "Rupa nakon 2. dolaska"
                else if pokrivenost < prag then "Količina nije dovoljna"
                else if traje >= prag then "Sve u redu"
                else "Stiže na vrijeme"
        in
            [ NazivArtikla = naziv,
              SifraDob = if s = null then "" else s[SifraDobavljaca],
              NazivDob = if s = null then "" else s[NazivDobavljaca],
              ZalihaPrethodno = prethodno,
              Zaliha = zaliha,
              Izlaz = if prethodno = null then null else prethodno - zaliha,
              Prosjek = prosjek,
              Traje = traje,
              M1 = m1, Q1 = q1, M2 = m2, Q2 = q2, M3 = m3, Q3 = q3,
              Pokrivenost = pokrivenost,
              Status = status ]),

    prosireno = Table.ExpandRecordColumn(polja, "X",
        {"NazivArtikla","SifraDob","NazivDob","ZalihaPrethodno","Zaliha","Izlaz",
         "Prosjek","Traje","M1","Q1","M2","Q2","M3","Q3","Pokrivenost","Status"}),

    // the Sij..Pro block: this year's raw monthly sales, 0 where none
    mjKol    = List.Transform({1..12}, each "Mj" & Text.From(_)),
    mjNazivi = {"Sij","Vel","Ožu","Tra","Svi","Lip","Srp","Kol","Ruj","Lis","Stu","Pro"},
    sMjesecima = Table.ExpandTableColumn(prosireno, "R", mjKol),
    mjNule   = Table.TransformColumns(sMjesecima,
        List.Transform(mjKol, (k) => {k, each _ ?? 0, type number})),

    // an empty column R between STATUS and the months, as in ver03; the
    // header is a single space because a table header cannot be empty
    razmak   = Table.AddColumn(mjNule, " ", each null),

    konacno = Table.SelectColumns(razmak,
        {"Artikal","NazivArtikla","SifraDob","NazivDob","ZalihaPrethodno","Zaliha",
         "Izlaz","Prosjek","Traje","M1","Q1","M2","Q2","M3","Q3","Pokrivenost","Status"," "}
        & mjKol),

    // worst first
    redoslijed = Table.AddColumn(konacno, "H", each
        if Text.StartsWith([Status], "Rupa") then 1
        else if [Status] = "Naruči odmah" then 2
        else if [Status] = "Količina nije dovoljna" then 3
        else if [Status] = "Stiže na vrijeme" then 4
        else if [Status] = "Sve u redu" then 5
        else 6, Int64.Type),
    sortirano = Table.Sort(redoslijed, {{"H", Order.Ascending},
                                        {"Traje", Order.Ascending}}),
    bezPomocnog = Table.RemoveColumns(sortirano, {"H"}),

    // "Zaliha 31.08." / "Izlaz od 31.08.": the date of the export that
    // qStanjePrethodno read (the newest one when there is only one)
    stanja    = fnDatoteke("stanje"),
    datPreth  = try stanja{List.Min({1, Table.RowCount(stanja) - 1})}[DatumIzvoza]
                otherwise null,
    dd        = (n as number) as text => Text.PadStart(Text.From(n), 2, "0"),
    oznaka    = if datPreth = null then "prethodno"
                else dd(Date.Day(datPreth)) & "." & dd(Date.Month(datPreth)) & ".",

    imena = Table.RenameColumns(bezPomocnog, {
        {"Artikal","Šifra"}, {"NazivArtikla","Naziv artikla"},
        {"SifraDob","Šifra dob."}, {"NazivDob","Naziv dobavljača"},
        {"ZalihaPrethodno","Zaliha " & oznaka}, {"Zaliha","Zaliha danas"},
        {"Izlaz","Izlaz od " & oznaka}, {"Prosjek","Prosjek/mj."},
        {"Traje","Zaliha traje (mj.)"},
        {"M1","Stiže 1"}, {"Q1","Količina 1"},
        {"M2","Stiže 2"}, {"Q2","Količina 2"},
        {"M3","Stiže 3"}, {"Q3","Količina 3"},
        {"Pokrivenost","Pokrivenost ukupno"}, {"Status","STATUS"}
    } & List.Zip({mjKol, mjNazivi}))
in
    imena

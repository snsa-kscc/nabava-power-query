# NABAVA — upute za rad

Ova datoteka svaki put iznova izračuna što treba naručiti, na temelju izvoza iz
ERP-a koje spremite u istu mapu. Ne diramo ništa unutar nje — samo dodajete nove
izvoze i osvježite.

---

## 1. Što vam treba (provjerite prije svega ostalog)

**Excel 2021 ili Microsoft 365.**

Na Excelu 2019 i 2016 model neće raditi — javit će grešku pri osvježavanju.
Ako ne znate koju verziju imate: `Datoteka → Račun → O programu Excel`.

---

## 2. Prvo otvaranje

### 2.1. Žuta traka na vrhu — obavezno kliknite

Kada datoteku dobijete mailom ili preuzmete, Excel će je otvoriti s upozorenjem
na vrhu:

> **ZAŠTIĆENI PRIKAZ** — `Omogući uređivanje`
>
> ili
>
> **SIGURNOSNO UPOZORENJE** Vanjske podatkovne veze onemogućene su —
> `Omogući sadržaj`

**Kliknite na to.** Dok ne kliknete, osvježavanje neće raditi — i neće javiti
nikakvu grešku, jednostavno se ništa neće dogoditi. Ovo je najčešći razlog
zbog kojeg se "ništa ne mijenja".

Kliknete samo jednom, prvi put.

### 2.2. Stavite datoteku u mapu s izvozima

Datoteka mora biti u istoj mapi u kojoj su ERP izvozi (ili u mapi iznad njih —
vidi točku 4.2).

### 2.3. Spremite datoteku (`Ctrl+S`)

Ovo nije formalnost. Dok datoteka nije spremljena na disk, ne može prepoznati
u kojoj je mapi, i polje `B4` na listu `POSTAVKE` ostaje prazno. Model će vam
u tom slučaju javiti:

> Spremite datoteku (Save) prije osvjezavanja.

Ako datoteku kasnije premjestite u drugu mapu, putanja se sama ažurira —
ne trebate ništa mijenjati.

### 2.4. Osvježite — `Ctrl+Alt+F5`

Isto možete i preko `Podaci → Osvježi sve`.

Prvo osvježavanje traje nekoliko sekundi. Nakon toga list `NABAVA` je popunjen
i obojen.

---

## 3. Svakodnevni rad

1. Izvezite iz ERP-a `Stanje skladišta` i po potrebi `Analiza prodaje artikala`
2. Spremite ih u mapu — **bez preimenovanja**
3. Otvorite NABAVA datoteku i pritisnite `Ctrl+Alt+F5`
4. Radite po listu `NABAVA`, sortiranom tako da je najgore na vrhu

To je sve. Nema koraka koji se mogu preskočiti ili zabuniti.

---

## 4. Mapa i datoteke

### 4.1. Nazivi se ne smiju mijenjati

Model prepoznaje datoteke po početku naziva i **datum čita iz naziva datoteke**:

| Datoteka | Model je prepoznaje kao |
|---|---|
| `Stanje skladišta 03.09.2026.xlsx` | trenutno stanje, datum 03.09.2026. |
| `Analiza prodaje artikala 03.09.2026.xlsx` | prodaja |
| `Narudžba dobavljaču ND 7814…xlsx` | **ignorira se** |

Zato ostavite nazive kakve ERP daje. Ako preimenujete datoteku i izgubite datum
iz naziva, model će uzeti datum zadnje izmjene datoteke, što je manje pouzdano.

Datoteke narudžbi dobavljačima smiju ostati u mapi — model ih preskače.
Isto tako preskače i samu sebe.

### 4.2. Podmape su dopuštene

Izvozi mogu biti u podmapi (npr. `data\`) — model pretražuje i podmape.

**Ali:** ista datoteka smije postojati **samo na jednom mjestu**. Ako je isti
`Stanje skladišta` u mapi i u podmapi, model će misliti da ima dva izvoza
istog datuma i izlaz robe će svugdje pokazivati nulu, bez ikakve greške.

### 4.3. Čuvajte stare izvoze

**`Stanje skladišta` — čuvajte barem dva.** Najnoviji je "danas", prethodni je
ono s čime se poredi. Iz te razlike nastaju stupci `Zaliha prethodno` i
`Izlaz od prethodnog`.

Dok u mapi imate samo jedan izvoz stanja, stupac `Izlaz od prethodnog` bit će
nula na svim artiklima. To nije greška — model nema s čime usporediti.

Koliko često izvozite, toliko gust je i prikaz izlaza:

- izvoz svaki dan → dnevni izlaz
- zadržite izvoz s 1. u mjesecu i jedan svježi → izlaz od početka mjeseca

**`Analiza prodaje` — čuvajte i prošlogodišnju.** Prosjek se računa preko
zadnjih 12 **završenih** mjeseci. U siječnju 2027. to znači velik dio 2026.,
pa datoteka iz 2026. mora ostati u mapi. Ako se u mapi nađe više izvoza prodaje
iz iste godine, uzima se najnoviji.

Tekući, nedovršeni mjesec se uvijek izbacuje iz prosjeka — inače bi tri dana
rujna spustila prosjek svim artiklima.

---

## 5. List `ROBA_U_DOLASKU` — ovdje vi upisujete

Jedan redak = jedna pošiljka. Do tri pošiljke po artiklu (ako upišete više,
uzimaju se tri najbliže). Prazne retke ne ostavljajte između upisanih.

| Stupac | Obavezno? | Čemu služi |
|---|---|---|
| `Artikal` | **da** | mora se točno podudarati sa šifrom iz ERP-a |
| `Kolicina` | **da** | količina u pošiljci |
| `Mjeseci do dolaska` | **da** | **ovo model koristi za izračun** |
| `Ocekivani datum` | ne | samo vaša bilješka, model ga ne čita |
| `Sifra dobavljaca` | ne | bilješka |
| `Narudzba` | ne | bilješka, npr. `ND_7858` |
| `Napomena` | ne | bilješka |

### 5.1. Najvažnije pravilo na cijelom listu

**`Mjeseci do dolaska` je broj koji se ne ažurira sam.**

Model računa isključivo s tim brojem. `Ocekivani datum` je samo tekst za vaše
oko — ako promijenite datum, izračun se neće promijeniti.

To znači: pošiljka koju ste upisali kao "2,2 mjeseca" ostat će "2,2 mjeseca" i
za tri mjeseca, i model će i dalje vjerovati da roba tek dolazi — iako je
odavno u skladištu.

Zato, svaki put kad radite s listom:

- **pošiljka je stigla** → izbrišite cijeli redak (roba je od tada u `Zaliha danas`)
- **pošiljka još nije stigla** → smanjite `Mjeseci do dolaska` na stvarno stanje

Praktično: prođite kroz ovaj list jednom mjesečno. Nije velik — trenutno ima
oko 70 redaka.

---

## 6. List `POSTAVKE`

Žuta polja mijenjate po potrebi, sivo se popunjava samo.

| Polje | Sada | Što znači |
|---|---|---|
| `B4` Mapa modela | automatski | datoteka sama prepozna svoju mapu — ne dirajte |
| `B6` Sigurnosni prag (mjeseci) | **5** | ispod ovoliko mjeseci zaliha se pali alarm |
| `B8` Kotrljajući prosjek (mjeseci) | **12** | duljina prozora za prosjek prodaje |
| `B7` Tranzit iz Kine (mjeseci) | 4 | **samo podsjetnik** — model ga ne koristi |

Ako povisite prag s 5 na 6, više artikala dobit će `Naruci odmah`. Ako smanjite
prozor prosjeka s 12 na 6, prosjek će brže reagirati na sezonu, ali i na
slučajne mjesece.

Nakon svake promjene: `Ctrl+Alt+F5`.

---

## 7. List `NABAVA` — što piše u stupcima

Sortirano je tako da je najhitnije na vrhu.

| Stupac | Značenje |
|---|---|
| `Sifra`, `Naziv artikla` | artikal |
| `Sifra dob.`, `Naziv dobavljaca` | dobavljač iz izvoza stanja |
| `Zaliha prethodno` | zaliha u prethodnom izvozu stanja |
| `Zaliha danas` | zaliha u najnovijem izvozu |
| `Izlaz od prethodnog` | koliko je otišlo između dva izvoza |
| `Prosjek/mj.` | prosječna mjesečna prodaja (12 završenih mjeseci) |
| `Zaliha traje (mj.)` | koliko mjeseci zaliha traje bez novih dolazaka |
| `Stize 1/2/3`, `Kolicina 1/2/3` | pošiljke iz `ROBA_U_DOLASKU` |
| `Pokrivenost ukupno` | koliko mjeseci pokrivate kad se uračunaju dolasci |
| `STATUS` | zaključak, vidi niže |

Artikli koji nisu u izvozu stanja (drugo skladište, novi artikl) svjesno su
uključeni: `Zaliha danas` im je 0, a `Sifra dob.` i `Izlaz od prethodnog` ostaju
prazni. Takvih je 158 — 139 ih se pojavljuje samo u prodaji, a 19 samo u
`ROBA_U_DOLASKU`.

---

## 8. Što koji `STATUS` znači

| STATUS | Znači |
|---|---|
| `Sve u redu` | zaliha pokriva potrošnju dulje od praga |
| `Naruci odmah` | zaliha pada pod prag, a ništa nije u dolasku |
| `Rupa za ~X mj.` | ostat ćete bez robe za približno X mjeseci, prije prve pošiljke |
| `Rupa nakon 1. dolaska` | prva pošiljka ne pokriva do druge |
| `Rupa nakon 2. dolaska` | druga pošiljka ne pokriva do treće |
| `Kolicina nije dovoljna` | pošiljke stižu na vrijeme, ali ih je ukupno premalo |
| `Stize na vrijeme` | pošiljka stiže prije nego zaliha padne pod prag |
| `Nema prodaje` | artikl nema zabilježenu prodaju — prosjek je 0 |

Boje na stupcima `STATUS` i `Zaliha traje (mj.)` pale se same.

---

## 9. Ako nešto ne radi

| Što vidite | Što je |
|---|---|
| Pritisnete osvježi i **ništa se ne promijeni** | žuta traka iz točke 2.1. nije potvrđena. Zatvorite i ponovno otvorite datoteku, kliknite `Omogući sadržaj` |
| `Spremite datoteku (Save) prije osvjezavanja.` | datoteka nije spremljena na disk — `Ctrl+S` |
| `U mapi nema datoteke 'Stanje skladista...'` | izvoz nije u mapi ili mu je naziv promijenjen |
| `Izlaz od prethodnog` je svugdje 0 | u mapi je samo jedan izvoz stanja, ili je isti izvoz na dva mjesta (točka 4.2) |
| Prosjek izgleda prenizak | provjerite je li prošlogodišnja `Analiza prodaje` još u mapi |
| Artikl ima zalihu, a `Naziv artikla` je prazan | artikl je samo u prodaji, nije u izvozu stanja |
| Osvježavanje traje neuobičajeno dugo | provjerite koliko izvoza je u mapi; svaki se čita pri osvježavanju |

Ako se pojavi greška koju ovdje ne vidite, prepišite tekst greške i pošaljite —
poruke su namjerno napisane da se mogu prepisati.

---

## 10. Što ovaj alat ne radi

Da ne bude nesporazuma:

- **ne čita narudžbe dobavljačima** — dolaske vodite ručno na `ROBA_U_DOLASKU`
- **ne smanjuje sam `Mjeseci do dolaska`** — vidi točku 5.1.
- **ne šalje ništa nikome** i ne mijenja ništa u ERP-u
- **ne pamti povijest** — svako osvježavanje računa sve iznova iz datoteka u mapi

Zadnje je namjerno: ne postoji stanje koje se može pokvariti. Ako nešto izgleda
čudno, obrišite ili dodajte izvoz u mapu i osvježite ponovno.

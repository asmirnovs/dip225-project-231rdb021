# CityBee rēķinu analīze

## Problēmas apraksts

Pagajušajā vasarā es veiksmīgi nokārtoju CSDD eksāmenu un dabūju vadītāja tiesības. Mani vecāki, lai motivētu mani iemācīties braukt vēl labāk, nolēma apmaksāt manus braucienus ar CityBee mašīnām viena gada laikā. Protams, visas izmaksas es skaitu pats, vecākiem atsūtot tikai kopējo summu. Lai iegūtu izmaksas par vienu nedeļu, es izmantoju CityBee rēķinus, kur ir iekļauti dati par visām mašīnām, ko es īrēju nedeļas laikā, kā arī kopēja summa, ko samaksāju - dokumentus uzņēmums automātiski sūta uz epastu katru nedeļu. Izmantojot iegūtas zināšanas kursa laikā, nolēmu izveidot Python scriptu, kas automatizēs to.

## Projekta apraksts

Manis izveidotais projekts analizē epastā nosūtītas CityBee rēķinus. Izmantojot imaplib bibliotēku programma automātiski pieslēdzas GMAIL serveram, nolasa no vēstulēm rēķinu linku un paroli un, izmantojot selenium bibliotēku, nolasa no katra rēķina kopējo izmaksāto summu par noteikto periodu. Beidzot, visu informāciju skripts pieraksta Excel tabulā, lai lietotājam ar to būtu ērti darboties nākotnē.

## Izmantotas bibliotēkas

Programmā ir izmantotas sēkojošās bibliotēkas:
- `imaplib` - ļauj scriptam pievienoties GMAIL servisam un nolasīt vēstuļu datus
- `email` - bibliotēka, kas pārveido nolasīto vēstuli klasē, ar kuru var darboties
- `selenium` - parlūkprogrammas automatizēšana, scriptā izmantota lai automātiski atvērt rēķinu, ievadīt paroli un nolasīt datus
- `openpyxl` - ļauj strādāt ar tabulām, palīdz attēlot iegūto informāciju Excel tabulā
- `re` - bibliotēka, atbildīga par RegEx izmantošanu. Izmantota, lai viegli atrast informāciju gan vēstulēs, gan rēķinā
- `datetime` - bibliotēka, ko izmanto pašā programmas sākumā, lai pārbaudītu, vai ievadītais datums ir pareizajā formatā

## Programmas uzbūve

Programma ir sadalīta divās daļās.

### 1. daļa - informācijas iegūšana no epasta

Sākumā, izmantojot lietotāja ievadīto datumu (ko programma pārbauda, izmantojot `datetime` bibliotēku), scripts pieslēdzās GMAIL serveram un pieprasa visas vēstules no `info@citybee.lv` ar frāzi `par sniegtajiem pakalpojumiem` vēstules tekstā. Tālāk programma izmanto ciklu, lai, izmantojot RegEx formulas, no katras vēstules iegūt rēķina linku, paroli, kā arī nosūtīšanas laiku. Tos programma saglabā dubultlīmeņu sarakstā, pēc kā aizver savienojumu ar serveri.

### 2. daļa - rēķinu apstrāde

Šajā solī programma izveido jauno webdriver instanci, kā arī tukšo Excel tabulu. Tālāk, izmantojot ciklu, programma atver katru rēķinu pēc kārtas (links un parole ir iegūti no vēstules). Izmantojot `find_elements` funkciju, programma iegūst kopējo rēķina summu, kā arī informāciju par katru pakalpojumu atsevišķi (mašīnas tipu, numuru, datumu, kad tā bija izīrēta, un izmaksas). Datus programma pieraksta tabulā, sadalot tos pa stabiņām. Katrā ciklas iteracijā tabula tiek saglabāta, lai samazinātu iespēju pazaudēt datus. Tā kā rēķinu weblapā bieži vien parādās kļūdas (kurus var viegli novērst, pārlādējot lappusi), cikla iekšā ir vēl viens cikls, kas nodrošina programmai iespēju nolasīt datus trīs reizes (izmantojot `range(3)` funkciju). Praksē tas ļauj pārlādēt lapu katru reizi kad parādās kļūda (jeb html elements nav atrasts). Beigās programma pievieno tabulā kopējo summu par visu lietotāja pieprasīto periodu un aizver dokumentus.

## Papildus resursi

- [Links](https://drive.google.com/file/d/1z5zhgP6DY23nWsITECMcE7M0U6SbpBOA/view?usp=sharing) uz video ar programmas darbības paraugu
- [Google Sheet](https://docs.google.com/spreadsheets/d/1taRcFp6s7BUwJKRaFS4WjLX7NXKLE2rRN8C54KcANOM/edit?usp=sharing) ar izveidotas tabulas piemēru

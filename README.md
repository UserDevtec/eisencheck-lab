# Eisencheck Lab

Eisencheck Lab is een webapp om twee Excel-bestanden te vergelijken op basis van een sleutelkolom en een gekozen tekstkolom. De app negeert volgorde, normaliseert zichtbare tekst (spaties/returns), en exporteert de verschillen naar Excel met een legenda en een apart tabblad voor vervallen eisen.

## Features
- Upload twee .xlsx bestanden (oud en nieuw).
- Kolomkoppeling: kies Eiscode en EisTekst per bestand.
- Vergelijking op zichtbare tekst: dubbele spaties en returns worden genegeerd.
- Resultaten met statuskleuren: groen (ongewijzigd), geel (toegevoegd), oranje (gewijzigd), rood (vervallen).
- Excel-export met tabs: Resultaat, Vervallen eisen, Legenda.

## Gebruik
1) Upload bestand A (oud) en bestand B (nieuw).
2) Kies de Eiscode kolommen en de EisTekst kolommen.
3) Klik "Vergelijk bestanden".
4) Download het Excel-overzicht.

## Ontwikkelen
```bash
npm install
npm run dev
```

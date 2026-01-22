# Eisencheck Lab

[![Deploy to GitHub Pages](https://github.com/UserDevtec/eisencheck-lab/actions/workflows/deploy.yml/badge.svg)](https://github.com/UserDevtec/eisencheck-lab/actions/workflows/deploy.yml)

<img width="1840" height="1023" alt="image" src="https://github.com/user-attachments/assets/e1b24965-9d3f-4663-80d8-ce0b2b21d200" />
<img width="1820" height="801" alt="image" src="https://github.com/user-attachments/assets/68732b45-47af-4f95-920e-9241365bd83a" />
<img width="1843" height="171" alt="image" src="https://github.com/user-attachments/assets/473c6c62-1672-476e-8aa0-e0bba5e6b729" />
<img width="1842" height="1022" alt="image" src="https://github.com/user-attachments/assets/99d9c365-c7b5-44ed-9f06-6913ef4d864d" />

Eisencheck Lab is een webapp om twee Excel-bestanden te vergelijken op basis van een sleutelkolom en een of meer gekozen tekstkolommen. De app negeert volgorde, normaliseert zichtbare tekst (spaties/returns), en exporteert de verschillen naar Excel met een legenda en een apart tabblad voor vervallen eisen.

## Features
- Upload twee .xlsx bestanden (bestand 1 = oud, bestand 2 = nieuw).
- Kolomkoppeling: kies Eiscode als sleutel en een of meer EisTekst kolommen per bestand.
- Vergelijking op zichtbare tekst: dubbele spaties en returns worden genegeerd.
- Resultaten met statuskleuren: groen (ongewijzigd), geel (toegevoegd), oranje (gewijzigd), rood (vervallen).
- Excel-export met tabs: Resultaat, Vervallen eisen, Legenda.

## Gebruik
1) Upload bestand 1 (oud) en bestand 2 (nieuw).
2) Kies de Sleutel kolommen (Eiscode) en de EisTekst kolommen (meerdere mogelijk).
3) Klik "Vergelijk bestanden".
4) Download het Excel-overzicht.

## Ontwikkelen
```bash
npm install
npm run dev
```

# PnP - Extensibility Library Demo Project

This project shows how to implement custom web component for the 'Search Results' Web Part.

## Documentation

-nvm use 10.24.0 (Alessio: testata compatibile con le versioni usate di spfx 1.11.0)

- npm install @pnp/sp@2.0.3 --save //for SharePoint operations
- npm i xlsx //To parser and writer for spreadsheet formats (.xlsx file)
- npm i file-saver //For saving files on the client-side

## TODO

Parametrizzare scelta se esportare file .xlsx (attuale) o .csv

## Build

gulp clean; gulp bundle --ship; gulp package-solution --ship  
